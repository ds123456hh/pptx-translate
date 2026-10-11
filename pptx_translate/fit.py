# -*- coding: utf-8 -*-
"""Conservative auto-shrink so Chinese text keeps fitting its original shape.

Chinese glyphs are ~1 em wide where Latin lowercase is ~0.5 em, so a translated
line can be markedly wider than the English it replaces.  Text boxes whose frame
is set to *shrink text on overflow* are handled by PowerPoint itself; fixed-size
frames and table cells are not, so this module reduces the font size of those
runs just enough to keep the text inside the original frame — the box geometry is
never touched, because moving or resizing shapes is exactly the "layout change"
we are trying to avoid.
"""

from __future__ import annotations

import logging

log = logging.getLogger("pptx_translate.fit")

EMU_PER_PT = 12700
EMU_PER_INCH = 914400
DEFAULT_FONT_PT = 18.0
MIN_SCALE = 0.70

CJK_EXTRA = set("，。：；？！、（）【】《》“”‘’…—·％￥")


def _char_em(ch: str) -> float:
    if ch in CJK_EXTRA:
        return 1.0
    o = ord(ch)
    if o >= 0x2E80:  # CJK, kana, hangul, fullwidth forms
        return 1.0
    if ch in " ":
        return 0.28
    if ch.isdigit():
        return 0.52
    if not ch.isalpha():
        return 0.30
    if ch.isupper():
        return 0.66
    return 0.52


def est_lines(text: str, width_emu: int, font_pt: float) -> tuple[float, int]:
    """Estimate rendered width (emu) and wrapped line count."""
    if not text:
        return 0.0, 0
    width_emu = max(1, width_emu)
    if "\x0b" in text or "\n" in text:
        parts = text.replace("\n", "\x0b").split("\x0b")
        widths = [sum(_char_em(c) for c in p) * font_pt * EMU_PER_PT for p in parts]
        return max(widths or [0.0]), max(1, len(parts))
    chars = sum(_char_em(c) for c in text)
    total = chars * font_pt * EMU_PER_PT
    if total <= width_emu:
        return total, 1
    return total, max(1, int(total / width_emu + 0.999))


def _frame_geometry(shape) -> tuple[int, int, int] | None:
    """(width, height, insets) with insets applied, or None when unusable."""
    try:
        width = int(shape.width)
        height = int(shape.height)
    except Exception:
        return None
    tf = getattr(shape, "text_frame", None)
    left = right = top = bottom = 0
    if tf is not None:
        try:
            left = int(tf.margin_left or 0)
            right = int(tf.margin_right or 0)
            top = int(tf.margin_top or 0)
            bottom = int(tf.margin_bottom or 0)
        except Exception:
            left = right = top = bottom = 0
    usable_w = width - left - right
    usable_h = height - top - bottom
    if usable_w <= 0 or usable_h <= 0:
        return None
    return usable_w, usable_h, 0


def _block_text(tf) -> str:
    return "\x0b".join("".join(r.text for r in p.runs) for p in tf.paragraphs)


def _font_size(tf, shape) -> float:
    for para in tf.paragraphs:
        for run in para.runs:
            size = run.font.size
            if size is not None:
                return size.pt
    # placeholder-inherited size: use the layout default as a rough guess
    try:
        if shape.is_placeholder:
            ph = shape.placeholder_format
            if ph.type is not None and int(ph.type) in (13, 14, 15, 16, 17):
                return 32.0
    except Exception:
        pass
    return DEFAULT_FONT_PT


def _sizes_need_shrink(text: str, usable_w: int, usable_h: int, font_pt: float) -> float | None:
    """Return the scale to apply when the text clearly overflows, else None."""
    if font_pt <= 1:
        return None
    total, lines = est_lines(text, usable_w, font_pt)
    line_h = font_pt * 1.22 * EMU_PER_PT
    need_w = total > usable_w * 0.98
    need_h = lines * line_h > usable_h * 0.98
    if not (need_w or need_h):
        return None
    ratio = min(
        (usable_w * 0.98 / total) if need_w and total else 1.0,
        (usable_h * 0.98 / (lines * line_h)) if need_h and lines else 1.0,
    )
    if ratio >= 0.95 or ratio < MIN_SCALE:
        return None
    return ratio


def _apply_scale(tf, ratio: float) -> bool:
    """Scale every explicitly sized run down by ``ratio`` (points, not EMU)."""
    from pptx.util import Pt

    applied = False
    for para in tf.paragraphs:
        for run in para.runs:
            size = run.font.size
            if size is None:
                continue
            new_pt = max(6.0, round(size.pt * ratio, 1))
            if new_pt < size.pt:
                run.font.size = Pt(new_pt)
                applied = True
    return applied


def _autofit_mode(tf) -> str:
    """'grow' (frame follows text), 'shrink' (PowerPoint shrinks), 'none' (fixed)."""
    try:
        mode = tf.auto_size
    except Exception:
        return "none"
    from pptx.enum.text import MSO_AUTO_SIZE

    if mode == MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT:
        return "grow"
    if mode == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE:
        return "shrink"
    return "none"


def _explicit_sizes(tf) -> bool:
    return any(r.font.size is not None for p in tf.paragraphs for r in p.runs)


def _is_title(shape) -> bool:
    try:
        if shape.is_placeholder:
            t = shape.placeholder_format.type
            if t is not None and "TITLE" in str(t):
                return True
            if t is not None and "CENTER_TITLE" in str(t):
                return True
    except Exception:
        pass
    return str(getattr(shape, "name", "")).startswith("Title")


def _shrink_shapes(shapes, stats: dict, include_titles: bool = True, dry_run: bool = False) -> None:
    for shape in shapes:
        if getattr(shape, "has_group", False):
            _shrink_shapes(shape.shapes, stats, include_titles, dry_run)
            continue

        if getattr(shape, "has_table", False) and shape.has_table:
            _shrink_table(shape.table, stats, dry_run)
            continue

        if not getattr(shape, "has_text_frame", False) or not shape.has_text_frame:
            continue
        tf = shape.text_frame
        text = _block_text(tf)
        if not text.strip():
            continue
        if tf.word_wrap is False:
            continue
        if not include_titles and _is_title(shape):
            stats["titles_kept"] += 1
            continue
        mode = _autofit_mode(tf)
        if mode != "none":
            # 'grow' lets the frame expand; 'shrink' is PowerPoint's own autofit.
            # Touching either would be an unnecessary format change.
            stats["handled_by_powerpoint"] += 1
            continue
        geo = _frame_geometry(shape)
        if geo is None:
            continue
        usable_w, usable_h, _ = geo
        font_pt = _font_size(tf, shape)
        ratio = _sizes_need_shrink(text, usable_w, usable_h, font_pt)
        if ratio is None:
            stats["skip"] += 1
            continue
        if dry_run:
            stats["would_shrink"].append(
                {"shape": shape.name, "ratio": round(ratio, 3),
                 "font_pt": font_pt, "chars": len(text)})
            continue
        if _apply_scale(tf, ratio):
            stats["shapes"] += 1
            log.debug("fit: %s x%.2f", shape.name, ratio)
        else:
            # no explicit run sizes: nothing we can safely scale
            stats["skip"] += 1


def _shrink_table(table, stats: dict, dry_run: bool = False) -> None:
    col_widths = []
    try:
        grid = table._tbl.find("{http://schemas.openxmlformats.org/drawingml/2006/main}tblGrid")
        for col in grid:
            col_widths.append(int(col.get("w") or 0))
    except Exception:
        col_widths = []
    for r, row in enumerate(table.rows):
        try:
            row_h = int(row.height or 0)
        except Exception:
            row_h = 0
        for c, cell in enumerate(row.cells):
            tf = cell.text_frame
            text = _block_text(tf)
            if not text.strip():
                continue
            col_w = int(cell.width or 0) or (col_widths[c] if c < len(col_widths) else 0)
            if col_w <= 0:
                continue
            usable_w = col_w - int(tf.margin_left or 0) - int(tf.margin_right or 0)
            usable_h = (row_h or int(cell.height or 0)) - int(tf.margin_top or 0) - int(tf.margin_bottom or 0)
            if usable_w <= 0 or usable_h <= 0:
                continue
            font_pt = _font_size(tf, cell)
            ratio = _sizes_need_shrink(text, usable_w, usable_h, font_pt)
            if ratio is None:
                stats["skip"] += 1
                continue
            if dry_run:
                stats["would_shrink"].append(
                    {"shape": f"table cell r{r}c{c}", "ratio": round(ratio, 3),
                     "font_pt": font_pt, "chars": len(text)})
                continue
            if _apply_scale(tf, ratio):
                stats["cells"] += 1
            else:
                stats["skip"] += 1
    return None


def shrink_to_fit(prs, include_titles: bool = True, dry_run: bool = False) -> dict:
    """Walk the whole deck and shrink overflowing explicit-size text.

    ``dry_run`` reports what would change without touching the presentation.
    """
    stats = {"shapes": 0, "cells": 0, "skip": 0,
             "handled_by_powerpoint": 0, "titles_kept": 0, "would_shrink": []}
    for slide in prs.slides:
        try:
            _shrink_shapes(slide.shapes, stats, include_titles, dry_run)
        except Exception as exc:  # never let cosmetic fitting break the run
            log.warning("fit: slide skipped: %s", exc)
    return stats
