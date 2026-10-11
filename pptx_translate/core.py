#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx_translate.core — in-place, format-preserving PPTX translation engine.

Design
------
The package is treated as a ZIP of parts.  Every translatable run's text is
replaced in place, so paragraph/run/rPr, bullets, indents, fills, effects,
animations, theme, masters, layouts, hyperlinks and media are all preserved
byte-for-byte.  Nothing is ever rebuilt.

Translation units
-----------------
A ``Unit`` is one logical text block (a paragraph, a table cell paragraph, a
notes paragraph, a diagram run).  Its ``text`` may contain embedded line-break
characters (``\x0b`` vertical tab, ``\n``, ``\r``); those survive round-trip.
Emoji and non-linguistic symbols are swapped for ``[[E1]]``-style tokens before
they reach the model and restored afterwards, so they can never be dropped or
"helpfully" translated.

Applying a translation back into a text frame
---------------------------------------------
The original paragraphs each own a list of runs.  The translated block is split
into exactly ``len(paragraphs)`` pieces so the one-paragraph-per-line structure
of the source is kept.  Within a paragraph:
  * 1 run  -> the whole line is written into that run (its rPr is untouched);
  * N runs -> the line is split proportionally to the original run lengths, so
    per-run overrides (bold/colour/hyperlink) stay on the words they decorated.
A paragraph that ends up with no text keeps its paragraph properties (``pPr``),
so bullets and spacing stay.
"""

from __future__ import annotations

import copy
import hashlib
import io
import json
import logging
import os
import re
import sys
import threading
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

from .simplify import to_simplified

log = logging.getLogger("pptx_translate")

EMU_PER_INCH = 914400
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
A_T = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
A_R = "{http://schemas.openxmlformats.org/drawingml/2006/main}r"
A_BR = "{http://schemas.openxmlformats.org/drawingml/2006/main}br"
A_Fld = "{http://schemas.openxmlformats.org/drawingml/2006/main}fld"
A_P = "{http://schemas.openxmlformats.org/drawingml/2006/main}p"

LANG_CJK_RE = re.compile(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")


# --------------------------------------------------------------------------- #
# text helpers
# --------------------------------------------------------------------------- #
BREAK_CHARS = ("\r\n", "\x0b", "\n", "\r")
VB_ESCAPE = "_x000B_"
# Models echo a line break back in every spelling imaginable; all of these mean
# one soft break, and none of them should ever reach the slide as literal text.
_VB_ALIASES = ("\\u000b", "\\u000B", "\\v", "\\x0b", "\\011", "\x0b", VB_ESCAPE)


def normalize_breaks(text: str) -> str:
    """Collapse every spelling of a manual line break into ``\\x0b``.

    Two different things are handled here: an ``<a:br/>`` soft break that
    python-pptx reads back as the ``_x000B_`` escape, and an escaped sequence the
    model typed out (``\\v``, ``\\u000b``) instead of emitting a real newline.
    """
    for alias in _VB_ALIASES:
        if alias in text:
            text = text.replace(alias, "\x0b")
    return text


def split_breaks(text: str) -> list[str]:
    """Split a block on line-break characters, keeping the separators."""
    return re.split(r"(\r\n|\x0b|\n|\r)", text)


def text_parts(text: str) -> list[str]:
    """The text segments of a block, with the break separators removed.

    ``split_breaks`` keeps its separators, so they must be filtered by identity;
    a naive ``p != ""`` leaves the ``\\x0b`` separators behind and they reappear
    as stray breaks when the pieces are joined again.
    """
    return [p for p in split_breaks(normalize_breaks(text)) if p not in BREAK_CHARS]


def is_blank(text: str) -> bool:
    """True when there is nothing worth sending to the model."""
    stripped = re.sub(r"[\s\u200b\ufeff\xa0]+", "", text)
    if not stripped:
        return True
    # pure punctuation / digits / symbols (e.g. "->", "2024", "#3")
    if not re.search(r"[A-Za-z\u00c0-\u024f\u0400-\u04ff]", stripped):
        return True
    if LANG_CJK_RE.search(stripped):
        return True  # already Chinese; nothing to do
    return False


def chinese_ratio(text: str) -> float:
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return 0.0
    return sum(1 for c in letters if LANG_CJK_RE.match(c)) / len(letters)


# --------------------------------------------------------------------------- #
# placeholder protection (emoji + symbols)
# --------------------------------------------------------------------------- #
EMOJI_RANGES = (
    (0x1F300, 0x1FAFF),
    (0x1F000, 0x1F2FF),
    (0x2600, 0x27BF),
    (0x2B00, 0x2BFF),
    (0x2190, 0x21FF),
    (0x2300, 0x23FF),
    (0x25A0, 0x25FF),
    (0xFE0F, 0xFE0F),
    (0x200D, 0x200D),
    (0x20E3, 0x20E3),
    (0x1F1E6, 0x1F1FF),
    (0x2700, 0x27BF),
    (0x1F900, 0x1F9FF),
    (0x1FA70, 0x1FAFF),
)


def _is_emoji(ch: str) -> bool:
    o = ord(ch)
    return any(lo <= o <= hi for lo, hi in EMOJI_RANGES)


def _emoji_groups(text: str) -> list[str]:
    """Split text into emoji "clusters" and everything else.

    A single visible emoji can be several code points (base + variation selector,
    ZWJ sequence, skin tone, regional-indicator pair, keycap). Grouping them keeps
    one token per user-visible emoji instead of three or four.
    """
    joiner = {"\u200d", "\ufe0f", "\ufe0e", "\u20e3"}
    out: list[str] = []
    i = 0
    while i < len(text):
        ch = text[i]
        if _is_emoji(ch):
            group = ch
            i += 1
            while i < len(text) and (_is_emoji(text[i]) or text[i] in joiner):
                group += text[i]
                i += 1
            out.append(group)
        else:
            out.append(ch)
            i += 1
    return out


class PlaceholderMap:
    """Replace emoji (and box-drawing/arrow glyphs) with stable ASCII tokens.

    The mapping is derived from the text itself — emoji are numbered in the
    order they first appear in *that* string — so it is deterministic and needs
    no shared state. That matters for the cache: a cached translation is stored
    still-tokenised, and decoding it later cannot be misled by a different run's
    numbering.
    """

    def mapping_for(self, text: str) -> dict[str, str]:
        """emoji cluster -> token, in order of first appearance."""
        mapping: dict[str, str] = {}
        for piece in _emoji_groups(text):
            if _is_emoji(piece[0]) and piece not in mapping:
                mapping[piece] = f"[[E{len(mapping) + 1}]]"
        return mapping

    def protect(self, text: str) -> str:
        mapping = self.mapping_for(text)
        if not mapping:
            return text
        return "".join(mapping.get(piece, piece) for piece in _emoji_groups(text))

    def restore(self, text: str, mapping: dict[str, str]) -> str:
        """Put the emoji back, tolerating the spacing the model may insert."""
        if not mapping or "[[" not in text:
            return text
        reverse = {token: emoji for emoji, token in mapping.items()}
        pattern = re.compile(r"\[{1,2}\s*(?:E|e|3)?\s*_?\s*(\d{1,2})\s*\]{1,2}")

        def put(match: re.Match) -> str:
            token = f"[[E{match.group(1)}]]"
            return reverse.get(token, match.group(0))

        return pattern.sub(put, text)


# --------------------------------------------------------------------------- #
# units
# --------------------------------------------------------------------------- #
@dataclass
class Unit:
    """One translatable block of text."""

    uid: str
    text: str
    kind: str
    slide: int
    where: str
    source: str  # original text, verbatim
    index: int = 0  # paragraph index inside its frame
    target: str | None = None
    translated: str | None = None
    note: str = ""


@dataclass
class Inspection:
    units: list[Unit] = field(default_factory=list)
    slides: int = 0
    media: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    custom_parts: list[str] = field(default_factory=list)
    part_names: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# read side
# --------------------------------------------------------------------------- #
def _walk_shapes(shapes, slide_no: int, where: str, sink: Callable):
    """Recursively visit shapes, including group members."""
    for shape in shapes:
        if getattr(shape, "has_group", False):
            _walk_shapes(shape.shapes, slide_no, f"{where}/{shape.name}", sink)
            continue
        sink(shape, slide_no, where)


def _collect_text_frame(tf, slide_no: int, where: str, kind: str, bucket: list[Unit]) -> None:
    for index, para in enumerate(tf.paragraphs):
        raw = para._p
        pieces: list[str] = []
        for child in raw:
            tag = child.tag
            if tag == A_R:
                t = child.find(A_T)
                if t is not None and t.text:
                    pieces.append(t.text)
            elif tag == A_BR:
                pieces.append("\x0b")
        text = normalize_breaks("".join(pieces))
        if not text.strip():
            continue
        bucket.append(
            Unit(
                uid="",
                text=text,
                kind=kind,
                slide=slide_no,
                where=where,
                source=text,
                index=index,
            )
        )


# --------------------------------------------------------------------------- #
# part (chart / diagram) XML handling
# --------------------------------------------------------------------------- #
C_V = "{http://schemas.openxmlformats.org/drawingml/2006/chart}v"
C_STR_CACHE = "{http://schemas.openxmlformats.org/drawingml/2006/chart}strCache"


def extract_part_text(xml_bytes: bytes, part_name: str, part_bucket: dict) -> list[Unit]:
    """Pull every translatable string out of a chart or diagram part.

    Two shapes of text live in a chart part:
      * ``<a:t>`` inside ``c:rich`` — titles, axis titles, data labels;
      * ``<c:v>`` inside ``c:strCache`` — series names and string category labels.
    Numeric caches (``c:numCache``) hold data values and are never touched.
    """
    from lxml import etree

    units: list[Unit] = []
    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError as exc:
        log.warning("unparseable part %s: %s", part_name, exc)
        return units

    kind = "chart" if "/charts/" in part_name else "diagram"
    records: list[dict] = []

    def consider(node, index: int, label: str) -> None:
        text = node.text or ""
        if not text.strip() or is_blank(text):
            return
        records.append({"t": node, "text": text})
        units.append(
            Unit(uid="", text=text, kind=kind, slide=0,
                 where=f"{part_name}#{label}[{index}]", source=text)
        )

    for idx, t in enumerate(root.iter(A_T)):
        consider(t, idx, "a:t")
    for idx, v in enumerate(root.iter(C_V)):
        parent = v.getparent()
        if parent is None or parent.getparent() is None:
            continue
        if parent.getparent().tag != C_STR_CACHE:
            continue  # numCache / literals: data values, not language
        consider(v, idx, "c:v/strCache")

    part_bucket[part_name] = {"root": root, "records": records, "kind": "xml"}
    return units


def apply_part_text(part_bucket: dict, part_name: str, mapping: dict[str, str]) -> int:
    entry = part_bucket.get(part_name)
    if not entry:
        return 0
    changed = 0
    for rec in entry["records"]:
        new = mapping.get(rec["text"])
        if new and new != rec["text"]:
            rec["t"].text = new
            changed += 1
    return changed


# --------------------------------------------------------------------------- #
# apply side (text frames)
# --------------------------------------------------------------------------- #
# Matches any placeholder this pipeline can emit: [[E1]] emoji or [[G1]] term.
TOKEN_RE = re.compile(r"\[{1,2}\s*(?:E|e|G|g|3|6)?\s*_?\s*\d{1,2}\s*\]{1,2}")


def strip_unknown_tokens(text: str, known: set[str]) -> tuple[str, list[str]]:
    """Delete placeholders that were never in the source. Returns (text, removed).

    A model shown ``[[E1]]`` in its instructions will sometimes *invent* one in a
    paragraph that contained no emoji at all: it imitates the format rather than
    the meaning. Any token outside ``known`` is therefore a fabrication and is
    deleted, not "restored". The pattern is deliberately permissive — a local
    model produced ``[[U1]]`` and ``[[E12]]`` from the same habit — but because
    ``known`` holds every token the source could legitimately carry, being
    permissive cannot damage a real one.

    A fabricated token usually lands next to the punctuation it was standing in
    for (``方法：[[E1]]：攻击者``), so when a token is removed the punctuation on
    both sides is merged instead of left doubled.
    """
    if "[" not in text:
        return text, []

    pattern = re.compile(
        r"[ \t]*\[{1,2}\s*[A-Za-z]?\s*_?\s*\d{1,2}\s*\]{1,2}[ \t]*"
    )
    removed: list[str] = []

    def drop(match: re.Match) -> str:
        token = match.group(0).strip()
        if token in known:
            return match.group(0)
        removed.append(token)
        return ""

    cleaned = pattern.sub(drop, text)
    if not removed:
        return text, []

    # merge punctuation left doubled by the deletion, e.g. "：：" or ": 数据"
    for _ in range(4):
        merged = re.sub(r"([：:，,。.；;！!？?、])\s*([：:，,。.；;！!？?、])", r"\2", cleaned)
        if merged == cleaned:
            break
        cleaned = merged
    cleaned = re.sub(r"[ \t]{2,}", " ", cleaned)
    cleaned = re.sub(r"^[：:，,。.；;！!？?、]\s*", "", cleaned)
    return cleaned.strip(), removed


_QUOTE_PAIRS = {'"': '"', "'": "'", "“": "”", "‘": "’", "「": "」", "『": "』"}


def strip_redundant_quotes(text: str) -> str:
    """Drop one quote pair that wraps the whole line.

    Small translation models like to wrap a whole paragraph in quotation marks
    (``“输入：”`` for a source that had none). On a slide a label is not a
    quotation, so the wrapping pair is noise — but only when it truly encloses
    the entire line, and only when the source itself was not quoted.
    """
    stripped = text.strip()
    if len(stripped) < 3:
        return stripped
    opener = stripped[0]
    closer = _QUOTE_PAIRS.get(opener)
    if not closer:
        return stripped
    if not stripped.endswith(closer):
        return stripped
    inner = stripped[1:-1]
    if not inner.strip():
        return stripped
    # only collapse when no stray matching quote remains inside
    return inner.strip()


OPEN_BRACKETS = "([{（【《「『"
CLOSE_BRACKETS = ")]}）】》」』"


def repair_brackets(translation: str, source: str) -> tuple[str, str]:
    """Balance brackets against the source's counts. Returns (text, note).

    A small MT model keeps producing ``（ATT&CK>）`` or ``（ATLAS））`` for a source
    that had exactly one balanced pair each: it garbles or duplicates the closing
    bracket. Prompt instructions did not fix it reliably, so the counts are
    compared instead.

    Two repairs, both conservative:

    * a mangled closer — ``>`` sitting inside an unclosed ``（…`` — is restored to
      ``）`` (and ``<`` likewise);
    * surplus closers are deleted, choosing the ones with no opener to their
      left, so a properly matched pair never loses its bracket.

    Nothing is ever inserted, so a repair cannot invent structure the source did
    not have.
    """
    if not translation:
        return translation, ""
    src_open = sum(source.count(c) for c in OPEN_BRACKETS)
    src_close = sum(source.count(c) for c in CLOSE_BRACKETS)
    if src_open == 0 and src_close == 0:
        return translation, ""          # brackets the model added out of nowhere
    text = translation
    notes: list[str] = []

    def counts(s: str) -> tuple[int, int]:
        return (sum(s.count(c) for c in OPEN_BRACKETS),
                sum(s.count(c) for c in CLOSE_BRACKETS))

    # 1) a closer mangled into an angle bracket inside an otherwise open pair
    if ">" in text or "<" in text:
        after = None      # remember the char we replaced, for the note
        for idx, ch in enumerate(text):
            if ch not in "<>":
                continue
            before = text[:idx]
            # it must sit after an opener that has no closer yet, and before a
            # real closer:  （ATT&CK>）   ->   （ATT&CK）
            if (sum(before.count(c) for c in OPEN_BRACKETS)
                    > sum(before.count(c) for c in CLOSE_BRACKETS)
                    and any(c in CLOSE_BRACKETS for c in text[idx + 1:])):
                text = text[:idx] + "）" + text[idx + 1:]
                after = ch
                break
        if after:
            notes.append(f"restored a mangled closer ({after} -> ）)")

    # 2) delete surplus closers, preferring the ones with no opener to their left
    out_open, out_close = counts(text)
    surplus = out_close - src_close
    if surplus > 0:
        removed = 0
        while removed < surplus:
            depth = 0
            orphan = None
            for i, ch in enumerate(text):
                if ch in OPEN_BRACKETS:
                    depth += 1
                elif ch in CLOSE_BRACKETS:
                    if depth == 0:
                        orphan = i
                        break
                    depth -= 1
            if orphan is not None:
                text = text[:orphan] + text[orphan + 1:]
            else:
                for i in range(len(text) - 1, -1, -1):
                    if text[i] in CLOSE_BRACKETS:
                        text = text[:i] + text[i + 1:]
                        break
            removed += 1
        notes.append(f"removed {removed} surplus closer(s)")

    return text, "; ".join(notes)


def _xml_safe(value: str) -> str:
    """Encode a translated line as valid, PowerPoint-renderable DrawingML text.

    Two constraints meet here:

    * a manual line break inside one DrawingML run is written as the ``_x000B_``
      escape — a bare newline inside ``<a:t>`` is not a soft break and leaves the
      part marginally non-conformant, so soft breaks are emitted as the escape;
    * lxml refuses C0 control characters, so a stray one must not reach the tree.

    Whitespace-and-structure characters that legacy decks really do use — tab
    (``_x0009_``) and carriage return (``_x000D_``) — are escaped the same way.
    """
    value = value.replace("\x0b", VB_ESCAPE)
    value = value.replace("\t", "_x0009_").replace("\r", "_x000D_")
    return "".join(
        ch if ch in "\n" or ord(ch) >= 0x20 else " " for ch in value
    )


def _split_line_to_runs(line: str, weights: Sequence[int]) -> list[str]:
    """Split a translated line across n runs, proportional to the source weights.

    Every run gets at least one character whenever the text is long enough, so a
    bold or coloured run never ends up empty (which would silently drop that
    formatting from the slide).
    """
    n = len(weights)
    if n == 0:
        return []
    if n == 1:
        return [line]
    total_weight = sum(weights) or n
    sizes = [1] * n
    leftover = len(line) - n
    if leftover > 0:
        for i, weight in enumerate(weights):
            sizes[i] += leftover * weight // total_weight
    # hand out the rounding remainder, longest original run first
    order = sorted(range(n), key=lambda i: -weights[i])
    idx = 0
    while sum(sizes) < len(line):
        sizes[order[idx % n]] += 1
        idx += 1
    out: list[str] = []
    cursor = 0
    for size in sizes:
        out.append(line[cursor : cursor + size])
        cursor += size
    if cursor < len(line):
        out[-1] += line[cursor:]
    return out


def _write_line_into_group(group: Sequence, line: str) -> None:
    """Write one translated line into the text nodes of one original line.

    Working on the ``<a:t>`` nodes directly (rather than on ``Paragraph.runs``)
    keeps multiple ``<a:t>`` children of one run, and any other child such as a
    field, exactly where they were.
    """
    if not group:
        return
    targets = [node for node in group if (node.text or "").strip()]
    if not targets:
        targets = list(group)
    pieces = _split_line_to_runs(line, [max(1, len(node.text or "")) for node in targets])
    for node, piece in zip(targets, pieces):
        node.text = _xml_safe(piece)
    for node in group:
        if node not in targets:
            node.text = ""


def _paragraph_lines(para) -> list[list]:
    """Group a paragraph's runs into visual lines, split by ``<a:br>`` children.

    A manual line break lives *between* runs, so a paragraph that looks like
    ``run1 <br> run2`` renders as two lines but exposes two runs.  Writing the
    second translated line into ``run2`` (the naive positional approach) would
    overwrite the first line's text, so the break structure has to be honoured.
    """
    groups: list[list] = [[]]
    for child in para._p:
        if child.tag == A_R:
            t = child.find(A_T)
            if t is None:
                from lxml import etree

                t = etree.SubElement(child, A_T)
            groups[-1].append(t)
        elif child.tag == A_BR:
            groups.append([])
    if not groups[-1] and len(groups) > 1:
        groups.pop()
    return groups


def _read_group(group: Sequence) -> str:
    return normalize_breaks("".join(node.text or "" for node in group))


def _fit_lines_to_groups(lines: Sequence[str], count: int) -> list[str]:
    """Normalise a translated block to exactly ``count`` line slots."""
    out = list(lines)
    if len(out) < count:
        out += [""] * (count - len(out))
    elif len(out) > count:
        out = out[: count - 1] + ["\x0b".join(out[count - 1 :])]
    return out


def apply_block(tf, lines: Sequence[str]) -> bool:
    """Distribute translated lines over a frame, one line per visual line.

    The paragraph / ``<a:br>`` skeleton of the frame is never modified: every
    original line group keeps its runs, and therefore its formatting.
    """
    changed = False
    text_lines: list[str] = []
    for line in lines:
        text_lines.extend(text_parts(line))
    for para in tf.paragraphs:
        groups = _paragraph_lines(para)
        if not groups or not any(groups):
            continue
        inner = _fit_lines_to_groups(text_lines, len(groups))
        for group, line in zip(groups, inner):
            if not group:
                continue
            before = _read_group(group)
            _write_line_into_group(group, line)
            if _read_group(group) != before:
                changed = True
    return changed


def _set_paragraph_text(para, text: str) -> None:
    """Set a single paragraph's text, preserving its ``<a:br>`` structure."""
    groups = _paragraph_lines(para)
    if not groups:
        return
    inner = _fit_lines_to_groups(text_parts(text), len(groups))
    for group, line in zip(groups, inner):
        if group:
            _write_line_into_group(group, line)

def apply_units_to_frame(tf, by_index: dict[int, str]) -> bool:
    """Write translations into specific paragraphs of a frame.

    ``by_index`` maps a paragraph index to its translated text.  Paragraphs that
    are absent keep their original text; every paragraph keeps its ``<a:br>``
    skeleton and its runs, so formatting never moves.
    """
    changed = False
    for index, para in enumerate(tf.paragraphs):
        if index not in by_index:
            continue
        groups = _paragraph_lines(para)
        if not groups or not any(groups):
            continue
        inner = _fit_lines_to_groups(text_parts(by_index[index]), len(groups))
        for group, line in zip(groups, inner):
            if not group:
                continue
            before = _read_group(group)
            _write_line_into_group(group, line)
            if _read_group(group) != before:
                changed = True
    return changed


# --------------------------------------------------------------------------- #
# providers
# --------------------------------------------------------------------------- #
SYSTEM_PROMPT = """You are a professional technical translator localising cybersecurity \
training slide decks from English into {target_language}.

You receive ONE slide paragraph at a time. It may contain:
- EMOJI TOKENS such as [[E1]] or [[E12]]: these stand for emoji. Copy every token \
into your output unchanged, in the same position. Never drop, merge, translate or \
invent one.
- LOCKED TERMS such as [[G1]]: copy these tokens unchanged as well. They are \
placeholders whose real wording is restored afterwards.
- a literal \\v means a manual line break inside the paragraph. Keep exactly the same \
number of \\v markers, in the same places, so the paragraph keeps its line structure.
- leading bullet characters (@, -, •, digits) and tabs: keep them.

OUTPUT: only the translated paragraph, on its own, with no numbering, no quotes \
around it, no explanation and no markdown. Never add items that were not in the input.

Rules:
1. Translate meaning, not words. Write natural, fluent Simplified Chinese that a \
security practitioner would actually read. Keep sentences short and punchy, as \
befit slides.
2. Write Simplified characters only. Never emit Traditional forms: 规则 not 規則, \
范围 not 範圍, 严格 not 嚴格, 无 not 無, 误解 not 誤解, 红队 not 紅隊.
3. Every bracket must balance. The source's parentheses are part of its labels: \
``(ATT&CK)`` becomes ``（ATT&CK）``, never ``（ATT&CK）]`` or ``（ATLAS））``. Never add \
or drop a closing bracket, and never leave a stray ``]`` or ``)`` behind.
4. These terms must stay in English, verbatim, never translated and never followed by \
a Chinese gloss: {keep_terms}
   Most of them arrive as [[G…]] tokens; if one appears as plain text anyway, leave it \
in English.
5. Use full-width Chinese punctuation (，。：；？！) except inside kept English terms, \
code and URLs. Prefer Chinese quotation marks “ ” over ASCII quotes.
6. Numbers, versions, CVE ids, URLs, file paths, command lines and code stay exactly \
as they are.
7. If the input is only symbols or numbers, return it unchanged."""


# Only unambiguous proper nouns, acronyms and product names belong here. General
# concepts are deliberately absent: "penetration testing", "red teaming" and
# "token" all have natural Chinese renderings, and forcing them to English makes
# a sentence read worse rather than better.
DEFAULT_GLOSSARY = {
    "LLM": "LLM",
    "LLMs": "LLMs",
    "GPT": "GPT",
    "RAG": "RAG",
    "MITRE ATT&CK": "MITRE ATT&CK",
    "ATT&CK": "ATT&CK",
    "ATLAS": "ATLAS",
    "OWASP": "OWASP",
    "NIST": "NIST",
    "CVE": "CVE",
    "CWE": "CWE",
    "TTP": "TTP",
    "TTPs": "TTPs",
    "Shodan": "Shodan",
    "Burp Suite": "Burp Suite",
    "Metasploit": "Metasploit",
    "Nmap": "Nmap",
    "Wireshark": "Wireshark",
    "CrowdStrike": "CrowdStrike",
    "Mandiant": "Mandiant",
    "MITRE": "MITRE",
    "API": "API",
    "SDK": "SDK",
}

DEFAULT_KEEP_TERMS = (
    "LLM, LLMs, GPT, RAG, MITRE ATT&CK, ATT&CK, ATLAS, OWASP, NIST, CVE, CWE, "
    "TTP, TTPs, Shodan, Burp Suite, Metasploit, Nmap, Wireshark, API, SDK"
)

# The local provider's default checkpoint, re-exported so the CLI can advertise
# it without importing torch (only ``local`` needs the heavy dependencies).
DEFAULT_LOCAL_MODEL = "DALabCommunity/Haidass-Translate-143M"


@dataclass
class ProviderConfig:
    provider: str = "deepseek"
    base_url: str = "https://api.deepseek.com/v1"
    api_key: str = ""
    model: str = "deepseek-chat"
    target_language: str = "Simplified Chinese (简体中文)"
    timeout: float = 180.0
    retries: int = 4
    keep_terms: str = DEFAULT_KEEP_TERMS
    temperature: float = 0.2
    # How long a local Ollama server should keep the model in memory after a
    # request. Long enough to outlast a whole deck translation.
    ollama_keep_alive: str = "60m"


class TranslationError(RuntimeError):
    pass


class BaseProvider:
    name = "base"

    def __init__(self, cfg: ProviderConfig) -> None:
        self.cfg = cfg
        self.failed_items: list[str] = []
        self.lost_placeholders: list[str] = []
        self.system = SYSTEM_PROMPT.format(
            target_language=cfg.target_language,
            keep_terms=cfg.keep_terms,
        )

    def translate(self, items: Sequence[str]) -> list[str]:  # pragma: no cover
        raise NotImplementedError


def _http_post_json(url: str, payload: dict, headers: dict, timeout: float) -> dict:
    import urllib.error
    import urllib.request

    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode("utf-8", "replace")
    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise TranslationError(f"non-JSON reply: {body[:400]}") from exc


_QUOTE_RE = re.compile(r'"([^"\n]*)"')


def neutralize_quotes(text: str) -> str:
    """Replace ASCII double quotes with typographic ones.

    The transport is JSON, and this MT model reliably emits bare ``"`` inside a
    translation (``"忘记你的规则"``), which breaks the line's JSON. Curly quotes
    mean the same thing typographically and cannot break the envelope.
    """
    if '"' not in text:
        return text
    text = _QUOTE_RE.sub("\u201c\\1\u201d", text)
    return text.replace('"', "\u201d")


def parse_lines(raw: str) -> list[str]:
    """Parse an answer into one translation per emitted line, in order.

    The model is asked for ``{"i": n, "t": "..."}`` lines, but it also merges
    items and renumbers them, so its own ``i`` values are not trustworthy: an
    answer for of 8 items that comes back as 7 lines cannot say *which* line it
    merged. Line order is trustworthy, so translations are taken positionally and
    the caller treats a short answer as "the last N items are missing" rather
    than silently binding a translation to the wrong paragraph.
    """
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)

    out: list[str] = []
    for line in text.splitlines():
        line = line.strip().rstrip(",")
        if not line or line[0] not in "{[":
            continue
        if line[0] == "[":
            continue
        value: str | None = None
        try:
            obj = json.loads(line)
            if isinstance(obj, dict):
                for key in ("t", "text", "translation", "target"):
                    if isinstance(obj.get(key), str):
                        value = obj[key]
                        break
        except json.JSONDecodeError:
            m = re.match(r'^\{\s*"i"\s*:\s*\d+\s*,\s*"t"\s*:\s*"(.*)"\s*\}?$', line, re.S)
            if m:
                value = m.group(1)
        if value is not None and value.strip():
            out.append(value)
    return out


def parse_positional(raw: str) -> list[str] | None:
    """Whole-text replies: a JSON array, or a legacy ``{"translations": [...]}``."""
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)
    m = re.search(r"\[.*\]", text, re.S)
    if m:
        try:
            candidate = json.loads(m.group(0))
            if isinstance(candidate, list) and all(isinstance(x, str) for x in candidate):
                return candidate
        except json.JSONDecodeError:
            pass
    try:
        obj = json.loads(text)
    except json.JSONDecodeError:
        return None
    if isinstance(obj, dict):
        for key in ("translations", "translation", "result", "results", "data"):
            if isinstance(obj.get(key), list) and all(isinstance(x, str) for x in obj[key]):
                return obj[key]
    return None


def parse_single(raw: str) -> str:
    """Extract the one translation from a single-item reply.

    This model answers in whatever shape it likes (a JSON array, a quoted
    string, a bare sentence), so the reply is unwrapped rather than trusted.
    """
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)
    text = text.strip()
    arr = parse_positional(text)
    if arr:
        return arr[0]
    if len(text) >= 2 and text[0] == text[-1] and text[0] in "\"'“”「」":
        text = text[1:-1].strip()
    text = re.sub(r"^\s*\{\s*\"i\"\s*:\s*\d+\s*,\s*\"t\"\s*:\s*\"?", "", text)
    text = re.sub(r"\"?\s*\}\s*$", "", text)
    text = re.sub(r"^\s*@?\d+[@.]\s*", "", text)
    text = re.sub(r"\s*@?\d+@\s*$", "", text)
    return text.strip()


class OpenAICompatProvider(BaseProvider):
    """Any OpenAI-compatible /chat/completions endpoint (incl. DeepSeek)."""

    name = "openai"

    def _url(self) -> str:
        base = self.cfg.base_url.rstrip("/")
        if base.endswith("/chat/completions"):
            return base
        return base + "/chat/completions"
    def _payload(self, numbered: Sequence[tuple[int, str]]) -> dict:
        """One item per request — see ``_translate_items`` for why."""
        text = numbered[0][1]
        payload = {
            "model": self.cfg.model,
            "messages": [
                {"role": "system", "content": self.system},
                {
                    "role": "user",
                    "content": (
                        "Translate this single slide paragraph. Return only the "
                        f"translation, with no numbering and no explanation.\n{json.dumps(text, ensure_ascii=False)}"
                    ),
                },
            ],
            "temperature": self.cfg.temperature,
            "stream": False,
        }
        if self.cfg.provider == "deepseek":
            payload["thinking"] = {"type": "disabled"}
        if self.cfg.provider == "ollama":
            # Keep the model resident for the whole run. The default 5-minute
            # idle timer is shorter than a long deck takes to translate, and a
            # mid-run unload costs a full reload on the next paragraph.
            payload["keep_alive"] = self.cfg.ollama_keep_alive
        return payload

    def _headers(self) -> dict:
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "pptx-translate/1.0",
        }
        # A local server (Ollama) needs no credential; sending an empty bearer
        # token would be rejected by some gateways.
        if self.cfg.api_key:
            headers["Authorization"] = f"Bearer {self.cfg.api_key}"
        return headers

    def _translate_items(self, items: Sequence[str]) -> list[str]:
        """Translate items, one request per item.

        MT models of this class merge adjacent segments: given eight paragraphs
        they answer seven lines, and their own numbering cannot be trusted
        (they renumber whatever they did emit, and moved markers are relocated
        or dropped).  A merged answer is indistinguishable from a dropped one,
        so *no* positional scheme is safe for a batch — every later translation
        would shift onto the wrong paragraph.

        One item per request removes the ambiguity structurally.  The extra
        requests are cheap and slow-string caching means a mostly-translated
        deck re-runs almost for free.

        Raises ``TranslationError`` if an item cannot be translated, so the
        caller can leave it out of the cache and retry it on the next run.  A
        failure must never be recorded as if it were a translation.
        """
        items = list(items)
        out: list[str] = []
        failures: list[int] = []
        for position, item in enumerate(items, 1):
            try:
                value, last = self._translate_one(item, position, len(items))
            except Exception as exc:  # noqa: BLE001 - one bad item must not sink the batch
                value, last = None, f"provider error: {exc}"
            if value is None:
                # keep the source text so the deck is still usable, but do not
                # pretend it was translated: it stays out of the cache
                failures.append(position)
                out.append(item)
                log.warning("item %d/%d untranslated after %d attempts (%s)",
                            position, len(items), self.cfg.retries, last)
            else:
                out.append(value)
        if failures:
            failed_texts = [items[i - 1] for i in failures]
            self.failed_items.extend(failed_texts)
            if len(failures) == len(items):
                raise TranslationError(f"all {len(items)} item(s) failed")
            # A partial failure returns the successes so they are cached and
            # never re-requested; the failures are retried on the next run.
            log.warning("%d/%d item(s) in this batch failed and will be retried "
                        "on the next run", len(failures), len(items))
        return out

    def _translate_one(self, item: str, position: int, total: int) -> tuple[str | None, str]:
        """Translate a single paragraph. Returns (translation or None, reason)."""
        payload = self._payload([(1, item)])
        last = ""
        placeholder_retried = False
        for attempt in range(1, self.cfg.retries + 1):
            try:
                reply = _http_post_json(self._url(), payload, self._headers(), self.cfg.timeout)
                content = reply["choices"][0]["message"]["content"]
            except Exception as exc:  # noqa: BLE001
                last = str(exc)
                wait = self._backoff(attempt, exc)
                log.warning("request %d/%d failed (attempt %d/%d): %s — waiting %.0fs",
                            position, total, attempt, self.cfg.retries, exc, wait)
                time.sleep(wait)
                continue
            candidate = parse_single(normalize_breaks(content))
            if not candidate:
                last = "empty translation"
                log.warning("empty translation for item %d/%d; retrying", position, total)
                time.sleep(1.0)
                continue
            if self._placeholder_loss(item, candidate):
                if not placeholder_retried:
                    # one re-ask is worth it; looping on it is not, and a dropped
                    # emoji must not cost the whole translation
                    placeholder_retried = True
                    last = "placeholder dropped"
                    log.info("emoji token lost for item %d/%d; re-asking once", position, total)
                    continue
                self.lost_placeholders.append(item)
                log.warning("emoji token missing from the translation of item %d/%d",
                            position, total)
            return candidate, ""
        return None, last

    def _placeholder_loss(self, source: str, reply: str) -> bool:
        """True when the reply dropped or invented an ``[[E…]]`` emoji token."""
        src = set(re.findall(r"\[\[E\d+\]\]", source))
        if not src:
            return False
        got = set(re.findall(r"\[\[E\d+\]\]", reply))
        return not src.issubset(got)

    def _backoff(self, attempt: int, exc: Exception) -> float:
        """Wait longer for rate limiting, which this endpoint returns in bursts."""
        text = str(exc)
        if "429" in text or "Too Many" in text:
            return min(45.0, 6.0 * attempt)
        return min(20.0, 2.0 ** attempt)

    def translate(self, items: Sequence[str]) -> list[str]:
        return self._translate_items(list(items))


class MockProvider(BaseProvider):
    """Offline provider: verifies the whole pipeline without any network call."""

    name = "mock"

    def translate(self, items: Sequence[str]) -> list[str]:
        return [f"【{item}】" for item in items]


def build_provider(cfg: ProviderConfig) -> BaseProvider:
    kind = (cfg.provider or "").strip().lower()
    if kind == "mock":
        return MockProvider(cfg)
    if kind == "local":
        from .local import build_local_provider

        return build_local_provider(cfg)
    # 'ollama' exposes an OpenAI-compatible /chat/completions endpoint
    return OpenAICompatProvider(cfg)


# --------------------------------------------------------------------------- #
# translator
# --------------------------------------------------------------------------- #
class Translator:
    def __init__(
        self,
        provider: BaseProvider,
        batch_size: int = 20,
        concurrency: int = 4,
        cache_path: Path | None = None,
        use_cache: bool = True,
        glossary: dict[str, str] | None = None,
        placeholders: "PlaceholderMap | None" = None,
        simplify: bool = True,
    ) -> None:
        self.provider = provider
        self.batch_size = max(1, batch_size)
        self.concurrency = max(1, concurrency)
        self.cache_path = Path(cache_path) if cache_path else None
        self.use_cache = use_cache
        # Convert Traditional output to Simplified (a no-op for other targets).
        self.simplify = simplify
        self.placeholders = placeholders or PlaceholderMap()
        self.glossary = {k: v for k, v in (glossary or {}).items() if k and k.strip()}
        # Global, term-derived token numbering: [[Gn]] always means the same term.
        # Longest source first, so "MITRE ATT&CK" wins over "ATT&CK".
        self._ordered_terms = sorted(
            self.glossary.items(), key=lambda kv: (-len(kv[0]), kv[0].lower())
        )
        self._terms_by_token: list[str] = [dst for _src, dst in self._ordered_terms]
        self._patterns: list[tuple[re.Pattern, str]] = []
        for i, (src, _dst) in enumerate(self._ordered_terms):
            token = f"[[G{i + 1}]]"
            self._patterns.append(
                (
                    re.compile(
                        r"(?<![A-Za-z0-9])" + re.escape(src) + r"(?![A-Za-z0-9])",
                        re.IGNORECASE,
                    ),
                    token,
                )
            )
        self.cache: dict[str, str] = {}
        self.failed_sources: set[str] = set()
        self._lock = threading.Lock()
        digest_src = json.dumps(sorted(self.glossary.items()), ensure_ascii=False)
        self._glossary_digest = hashlib.sha256(digest_src.encode("utf-8")).hexdigest()[:8]

        if self.cache_path and self.use_cache and self.cache_path.exists():
            try:
                self.cache = json.loads(self.cache_path.read_text("utf-8"))
                log.info("cache: %d remembered string(s) from %s",
                         len(self.cache), self.cache_path)
            except Exception as exc:
                log.warning("unreadable cache %s: %s", self.cache_path, exc)
        elif self.cache_path and self.use_cache:
            log.info("cache: %s does not exist yet; starting empty", self.cache_path)

    # -- glossary ----------------------------------------------------------- #
    def lock_terms(self, text: str) -> str:
        """Replace glossary terms with ``[[Gn]]`` so the model cannot translate them.

        Token numbers are global and derived from the term itself, so ``[[G7]]``
        means the same term in every string. Numbering them per-string by
        position would make one token mean different terms in different
        paragraphs, and restoring it would then substitute the wrong word —
        silently and only in strings that happen to contain a subset of terms.
        """
        for pattern, token in self._patterns:
            text = pattern.sub(token, text)
        return text

    def unlock_terms(self, text: str) -> str:
        """Restore locked terms, tolerating however the model rewrote a token.

        The model routinely reformats ``[[G9]]`` into ``[[ G9 ]]`` or drops a
        bracket, and an exact-match restore would then leave the raw token
        visible on the slide. Any ``G<number>`` left inside brackets is mapped
        back to its term.
        """
        if not self._terms_by_token:
            return text
        pattern = re.compile(
            r"\[{1,2}\s*(?:G|g|6)?\s*(\d{1,2})\s*\]{1,2}"  # [[G9]], [[ 9 ]], [G9], a misread [[69]]
        )

        def restore(match: re.Match) -> str:
            index = int(match.group(1))
            if 1 <= index <= len(self._terms_by_token):
                return self._terms_by_token[index - 1]
            return match.group(0)

        text = pattern.sub(restore, text)
        # drop bracket residue the pattern could not consume (e.g. "[[G1 ]")
        return re.sub(r"\[{1,2}\s*G\s*\d{1,2}\s*\]{0,2}", "", text)
        if self.cache_path and self.use_cache and self.cache_path.exists():
            try:
                self.cache = json.loads(self.cache_path.read_text("utf-8"))
                log.info("cache: %d remembered string(s) from %s",
                         len(self.cache), self.cache_path)
            except Exception as exc:
                log.warning("unreadable cache %s: %s", self.cache_path, exc)
        elif self.cache_path and self.use_cache:
            log.info("cache: %s does not exist yet; starting empty", self.cache_path)

    # -- cache -------------------------------------------------------------- #
    def _save_cache(self) -> None:
        if not self.cache_path:
            return
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.cache_path.with_suffix(self.cache_path.suffix + ".tmp")
            tmp.write_text(json.dumps(self.cache, ensure_ascii=False, indent=0, sort_keys=True), "utf-8")
            tmp.replace(self.cache_path)
        except Exception as exc:
            log.warning("could not write cache: %s", exc)

    def cache_key(self, text: str) -> str:
        digest = hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]
        return f"{self.provider.cfg.model}|{self._glossary_digest}|{digest}"

    # -- one batch ---------------------------------------------------------- #
    def _translate_batch(self, texts: list[str]) -> list[str]:
        # emoji are swapped for [[En]] tokens so the model physically cannot drop
        # or "translate" them, then restored from the same per-string map
        prepared: list[tuple[str, dict[str, str]]] = []
        for text in texts:
            locked = self.lock_terms(text)
            prepared.append((self.placeholders.protect(locked), self.placeholders.mapping_for(locked)))
        try:
            result = self.provider.translate([p[0] for p in prepared])
        except Exception as exc:  # noqa: BLE001
            if len(prepared) <= 1:
                # a lone failing string keeps its source text; it is never cached
                log.warning("a string could not be translated (%s); keeping its source text", exc)
                result = [None] * len(prepared)
            else:
                # A single unmendable string must not take its whole batch down,
                # or every other string in it would be re-requested next run.
                log.warning("a batch of %d failed (%s); retrying its items one by one",
                            len(prepared), exc)
                result = []
                for text in [p[0] for p in prepared]:
                    try:
                        result.append(self.provider.translate([text])[0])
                    except Exception:  # noqa: BLE001
                        result.append(None)
        out: list[str] = []
        for src, translated, (_locked, emoji_map) in zip(texts, result, prepared):
            value = "" if translated is None else str(translated)
            value = self.placeholders.restore(value, emoji_map)
            value = self.unlock_terms(value)
            value, phantoms = strip_unknown_tokens(value, set())
            if phantoms:
                log.info("removed %d invented placeholder(s) %s from a translation",
                         len(phantoms), sorted(set(phantoms))[:3])
            value = normalize_breaks(value)
            value = value.replace("\r\n", "\n").replace("\n", "\x0b").strip()
            value, bracket_note = repair_brackets(value, src)
            if bracket_note:
                log.info("bracket repair on %r: %s", src[:40], bracket_note)
            if self.simplify:
                value = to_simplified(value)
            # only unwrap when the source carried no quotes of its own
            if not any(q in src for q in _QUOTE_PAIRS):
                value = strip_redundant_quotes(value)
            leftovers = re.findall(r"\[\[[A-Za-z]?\d+\]\]", value)
            if leftovers:
                log.warning("unrestored placeholder(s) %s left in a translation: %r",
                            sorted(set(leftovers))[:4], value[:90])
            out.append(value or src)
        return out

    # -- entry -------------------------------------------------------------- #
    def translate_texts(self, texts: Sequence[str], progress: Callable[[int, int], None] | None = None) -> dict[str, str]:
        """Map source text -> translation, using cache and parallel batches."""
        unique: list[str] = []
        seen: set[str] = set()
        for text in texts:
            if text in seen:
                continue
            seen.add(text)
            unique.append(text)

        todo: list[str] = []
        resolved: dict[str, str] = {}
        for text in unique:
            key = self.cache_key(text)
            hit = self.cache.get(key)
            if hit is not None:
                # the cache holds still-tokenised text: undo both layers, then
                # drop anything the model invented that was never in the source
                locked = self.lock_terms(text)
                decoded = self.placeholders.restore(hit, self.placeholders.mapping_for(locked))
                cleaned, phantoms = strip_unknown_tokens(self.unlock_terms(decoded), set())
                if phantoms:
                    log.info("removed %d invented placeholder(s) from a cached translation",
                             len(phantoms))
                cleaned = normalize_breaks(cleaned)
                cleaned, _note = repair_brackets(cleaned, text)
                if self.simplify:
                    cleaned = to_simplified(cleaned)
                if not any(q in text for q in _QUOTE_PAIRS):
                    cleaned = strip_redundant_quotes(cleaned)
                resolved[text] = cleaned
            else:
                todo.append(text)

        log.info("strings: %d unique, %d cached, %d to translate",
                 len(unique), len(resolved), len(todo))
        if not todo:
            if progress:
                progress(len(unique), len(unique))
            return resolved

        batches = [todo[i : i + self.batch_size] for i in range(0, len(todo), self.batch_size)]
        done_count = len(resolved)
        total = len(unique)
        sem = threading.Semaphore(self.concurrency)
        errors: list[str] = []

        def work(batch: list[str]) -> None:
            nonlocal done_count
            with sem:
                try:
                    results = self._translate_batch(batch)
                except Exception as exc:
                    log.error("batch of %d failed: %s", len(batch), exc)
                    with self._lock:
                        errors.append(str(exc))
                    return
                with self._lock:
                    for src, dst in zip(batch, results):
                        if dst is None or not str(dst).strip():
                            dst = src
                        resolved[src] = str(dst)
                        if str(dst) != src:
                            # Remember the success. A failure is never stored —
                            # and a failure must never *delete* an earlier
                            # success either, or re-running the same command
                            # would keep discarding good translations.
                            locked = self.lock_terms(str(dst))
                            self.cache[self.cache_key(src)] = self.placeholders.protect(locked)
                            self.failed_sources.discard(src)
                        else:
                            # "translation equals source" is how a failure and a
                            # deliberate keep-as-is both look. Record it as failed
                            # and let the caller report it honestly; a later
                            # request that succeeds clears it again.
                            self.failed_sources.add(src)
                    done_count += len(batch)
                    if progress:
                        progress(done_count, total)
                    self._save_cache()

        threads = []
        for batch in batches:
            t = threading.Thread(target=work, args=(batch,), daemon=True)
            t.start()
            threads.append(t)
        for t in threads:
            t.join()
        if errors:
            log.warning("%d batch(es) failed; those strings keep their original text", len(errors))
        return resolved

# --------------------------------------------------------------------------- #
# dry-run helpers
# --------------------------------------------------------------------------- #
def flatten(prs) -> Inspection:
    """Collect every translatable text unit from a presentation object."""
    insp = Inspection()
    insp.slides = len(prs.slides._sldIdLst)  # noqa: SLF001
    counter = 0

    def add(unit: Unit) -> None:
        nonlocal counter
        counter += 1
        unit.uid = f"u{counter}"
        insp.units.append(unit)

    def sink(shape, slide_no: int, where: str) -> None:
        label = f"{where}/{shape.name}"
        if shape.has_text_frame:
            bucket: list[Unit] = []
            _collect_text_frame(shape.text_frame, slide_no, label, "shape", bucket)
            for u in bucket:
                add(u)
        if getattr(shape, "has_table", False) and shape.has_table:
            for r, row in enumerate(shape.table.rows):
                for c, cell in enumerate(row.cells):
                    bucket = []
                    _collect_text_frame(cell.text_frame, slide_no, f"{label}/table[{r}][{c}]", "table", bucket)
                    for u in bucket:
                        add(u)
        if getattr(shape, "has_chart", False) and shape.has_chart:
            insp.warnings.append(f"slide {slide_no}: chart '{shape.name}' found (translated via its XML data part)")
        if shape.shape_type is not None and "PICTURE" in str(shape.shape_type):
            insp.media.append({
                "slide": slide_no,
                "name": shape.name,
                "width_in": round(shape.width / EMU_PER_INCH, 2),
                "height_in": round(shape.height / EMU_PER_INCH, 2),
            })

    for slide_no, slide in enumerate(prs.slides, 1):
        _walk_shapes(slide.shapes, slide_no, f"slide{slide_no}", sink)
        if slide.has_notes_slide:
            tf = slide.notes_slide.notes_text_frame
            bucket = []
            _collect_text_frame(tf, slide_no, f"slide{slide_no}/notes", "notes", bucket)
            for u in bucket:
                add(u)
    return insp


def part_units(part_name: str, xml_bytes: bytes) -> list[Unit]:
    bucket: dict = {}
    return extract_part_text(xml_bytes, part_name, bucket)


def rewrite_package(src: Path, dest: Path, part_updates: dict[str, bytes]) -> None:
    """Copy the source package, replacing only the given part payloads."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zin:
        infos = zin.infolist()
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                data = zin.read(info.filename)
                new = part_updates.get(info.filename)
                if new is not None:
                    data = new
                zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zi.internal_attr = info.internal_attr
                zi.create_system = info.create_system
                zout.writestr(zi, data)


def serialize_part(root) -> bytes:
    from lxml import etree

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
