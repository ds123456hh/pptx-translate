"""Conservative delivery assessment for translated presentations.

Automatic checks can reject an unsafe export or identify review work.  They
cannot establish that a sentence is faithful or that every slide looks good.
The result therefore never claims that a deck is ready for a client without
independent linguistic and visual review.
"""

from __future__ import annotations

from collections import Counter


CRITICAL_TRANSLATION_ISSUES = {
    "unchanged_or_empty", "unresolved_marker", "leading_label_missing",
    "numbers_missing", "latin_term_added", "mostly_not_chinese",
    "suspiciously_short", "measurement_units_changed",
}


def assess(report: dict) -> dict:
    """Classify machine-detectable blockers and the remaining review queue."""
    blockers: list[str] = []
    review: list[str] = []
    verification = report.get("verification") or {}

    if report.get("provider") == "mock":
        blockers.append("mock_provider_is_not_a_translation")
    if int(report.get("untranslated_strings") or 0):
        blockers.append("untranslated_strings")
    if int(report.get("api_failures") or 0):
        blockers.append("api_failures")
    if verification:
        if verification.get("slides") != report.get("slides"):
            blockers.append("slide_count_changed")
        if not verification.get("media_identical"):
            blockers.append("media_changed")
        if not verification.get("timing_same"):
            blockers.append("animation_changed")
        if verification.get("parts_added") or verification.get("parts_removed"):
            blockers.append("package_parts_changed")
    else:
        blockers.append("structural_verification_missing")
    visual_structure = report.get("visual_structure") or {}
    if visual_structure:
        if int(visual_structure.get("shape_set_differences") or 0):
            blockers.append("shape_set_changed")
        if int(visual_structure.get("geometry_differences") or 0):
            blockers.append("shape_geometry_changed")
        if int(visual_structure.get("run_style_differences") or 0):
            if report.get("autofit"):
                review.append("run_styles_changed_by_autofit")
            else:
                blockers.append("run_styles_changed")
    else:
        blockers.append("visual_structure_check_missing")

    counts: Counter[str] = Counter(report.get("quality_issue_counts") or {})
    if not counts:
        for warning in report.get("quality_warnings") or []:
            counts.update(warning.get("issues") or [])
    for issue in sorted(counts):
        if issue in CRITICAL_TRANSLATION_ISSUES:
            blockers.append(f"translation_{issue}:{counts[issue]}")
        else:
            review.append(f"translation_{issue}:{counts[issue]}")

    if (report.get("quality_warning_count") or 0) > len(report.get("quality_warnings") or []):
        review.append("quality_warnings_truncated")
    table = report.get("table_layout") or {}
    if int(table.get("unresolved") or 0):
        review.append(f"automatic_table_rows:{table['unresolved']}")
    if report.get("autofit_preview"):
        review.append(f"possible_text_overflow:{len(report['autofit_preview'])}")
    if report.get("autofit") and (report["autofit"].get("shapes") or report["autofit"].get("cells")):
        review.append("font_sizes_changed_by_autofit")
    if int(report.get("untranslated_note_paragraphs") or 0):
        review.append(f"speaker_notes_not_translated:{report['untranslated_note_paragraphs']}")
    if int(report.get("embedded_objects") or 0):
        review.append(f"embedded_objects_not_translated:{report['embedded_objects']}")
    if report.get("image_text_review", {}).get("candidates"):
        review.append("pictures_may_contain_untranslated_text")

    return {
        "status": "blocked" if blockers else "manual_review_required",
        "blockers": blockers,
        "review_items": review,
        "translation_accuracy_verified": False,
        "visual_layout_verified": False,
        "client_ready": False,
        "next_step": "Review meaning, terminology, every rendered slide, and image text before delivery.",
    }
