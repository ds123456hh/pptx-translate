# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2026-11-06

### Added

- `--qa-mode commercial` blocks detectable translation and package failures,
  records shape geometry and run-style comparisons, and marks outputs for
  independent language and visual review.
- Generic `scripts/final_qa.py` binds the report to the source/output hashes
  and rechecks package structure and placeholder residues.
- Cross-domain model regression examples and a repeatable evaluation script.
- Measurement-unit change warnings for common medical and engineering units.

### Changed

- README and package description no longer promise pixel-identical rendering.

## [1.1.0] - 2026-11-06

### Changed

- General-purpose translation prompt and empty default glossary; the original
  cybersecurity terms remain available with `--glossary-profile cybersecurity`.
- English paragraphs with short Chinese glosses are translated; Chinese-led
  teaching notes and multiple-choice labels are skipped.
- Translation memory keys include the prompt, target language and term settings.
- Output copies the original PPTX/PPTM package and replaces only edited XML,
  preserving all unrelated package parts byte-for-byte.

### Added

- Quality warnings for missing numbers or list labels, invented Latin terms,
  unresolved markers and suspiciously short translations. Leading list labels
  are restored automatically.
- Optional PowerPoint measurement of source automatic table row heights on
  Windows (`pip install -e ".[windows]"`), plus a portability warning when
  those rows cannot be stabilized.
- SHA-256 media comparison and broader grouped/table-text verification.

### Fixed

- Shape IDs now distinguish text boxes that share a display name, preventing
  one translated paragraph from overwriting another.
- Source tabs remain real tabs in PowerPoint instead of rendering as the
  literal `_x0009_` escape.
- Table auto-fit preview no longer tries to read unavailable cell width/height
  attributes, and wrapped lines are counted across manual line breaks.


## [1.0.0] - 2025

First release. The engine rewrites text nodes in place, so layout, styles,
media and animations survive byte-for-byte.

### Added

- **Format-preserving translation** of `.pptx` / `.pptm`: only `<a:t>` text is
  replaced; `rPr`, `pPr`, `<a:br/>`, shapes, media and timing are untouched.
- **Local model provider** (`--provider local`) running a small HuggingFace
  translation model in-process — no API key, no per-string cost, deck content
  never leaves the machine.
- **OpenAI-compatible provider** for DeepSeek, Tencent tokenhub, SiliconFlow and
  any other `/chat/completions` endpoint.
- **Offline provider** (`mock`) and **offline modes**: `--dry-run`,
  `--dump-strings`, `--strings-json`.
- **Term locking that cannot be ignored**: glossary hits are replaced by stable
  `[[Gn]]` placeholders before the request and restored afterwards, with
  tolerant recovery when the model reformats a token. Numbering is global and
  term-derived, so one token never means two different terms.
- **Emoji protection**: emoji clusters (variation selectors, ZWJ sequences,
  flags, keycaps) become one `[[En]]` token each and are restored exactly.
- **Soft-break preservation**: `<a:br/>` skeletons and `_x000B_` escapes survive;
  every spelling a model emits for a line break (`\v`, `\u000b`, `\\v`) is
  normalised, and output is written back as the valid escape.
- **Chart and diagram text** translation via their XML parts (`<a:t>` plus
  `<c:v>` inside `strCache`), leaving numeric caches intact.
- **Translation memory** that is safe to reuse: failures are never cached, and a
  failure never evicts an earlier success.
- **Rate-limit handling**: 429-aware backoff, per-item retry, and a per-item
  fallback so one bad string cannot take its whole batch down.
- **Optional auto-fit** (`--autofit`) that shrinks only fixed-size frames whose
  explicit font sizes overflow, plus `--autofit-preview` to inspect first.
- **Verification tooling**: per-run structure report and
  `scripts/verify_pptx.py` for shape-set, geometry and style diffing.
- Unit tests covering the run/paragraph/soft-break round-trip, chart parts,
  glossary tokens, placeholder protection and cache semantics.

[1.0.0]: https://github.com/ds123456hh/pptx-translate/releases/tag/v1.0.0
