"""Conservative, reviewable checks for model output.

These checks cannot prove a translation correct.  They catch regressions that
otherwise look like a successful format-preserving export: invented names,
dropped paragraph numbers, missing figures, and suspiciously short answers.
"""

from __future__ import annotations

import re
from collections import Counter

from .core import chinese_ratio

LATIN_WORD = re.compile(r"[A-Z]{2,}(?:[&/-][A-Z0-9]+)+|[A-Za-z]{4,}")
NUMBER = re.compile(r"(?<![A-Za-z\d])\d+(?:[.,]\d+)*(?:%|％)?(?![A-Za-z\d])")
LEADING_LABEL = re.compile(r"^\s*((?:\d{1,3}|[A-F])(?:[.)、．])?)\s+")
PLACEHOLDER = re.compile(r"\[\[[A-Za-z]?\d+\]\]")
MEASUREMENT = re.compile(
    r"(?<![A-Za-z0-9])\d+(?:[.,]\d+)?\s*"
    r"(mg/kg|µg/kg|μg/kg|mmHg|kPa|MPa|kWh|MWh|km/h|m/s|N[·.]m|"
    r"µg|μg|mg|kg|mL|ml|GB|MB|TB|GHz|MHz|kHz|Hz|°C|℃|°F)"
    r"(?=$|[\s,.;:，。；）)])"
)


def issues_for(source: str, target: str, target_language: str) -> list[str]:
    issues: list[str] = []
    if not target.strip() or target == source:
        issues.append("unchanged_or_empty")
    if PLACEHOLDER.search(target) or "\ufffd" in target:
        issues.append("unresolved_marker")
    match = LEADING_LABEL.match(source)
    if match and not target.lstrip().startswith(match.group(1)):
        issues.append("leading_label_missing")
    missing_numbers = Counter(NUMBER.findall(source)) - Counter(NUMBER.findall(target))
    if missing_numbers:
        issues.append("numbers_missing")
    source_units = Counter(unit for unit in MEASUREMENT.findall(source))
    target_units = Counter(unit for unit in MEASUREMENT.findall(target))
    if source_units - target_units:
        issues.append("measurement_units_changed")
    if "chinese" in target_language.lower() or "中文" in target_language:
        introduced = {w.lower() for w in LATIN_WORD.findall(target)} - {
            w.lower() for w in LATIN_WORD.findall(source)
        }
        if introduced:
            issues.append("latin_term_added")
        if len(source) >= 80 and len(target) < len(source) * 0.12:
            issues.append("suspiciously_short")
        if len(source) >= 40 and chinese_ratio(target) < 0.05 and target != source:
            issues.append("mostly_not_chinese")
    return issues


def audit_units(units, resolved: dict[str, str], target_language: str) -> list[dict]:
    """Report each suspect source once with its first slide location."""
    found: list[dict] = []
    seen: set[str] = set()
    for unit in units:
        source = unit.source
        if source in seen:
            continue
        seen.add(source)
        target = resolved.get(source, source)
        issues = issues_for(source, target, target_language)
        if issues:
            found.append({
                "slide": unit.slide,
                "where": unit.where,
                "issues": issues,
                "source": source,
                "translation": target,
            })
    return found
