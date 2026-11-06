"""End-to-end checks for the client-delivery gate, without any live API."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from pptx_translate.cli import main
from pptx_translate.quality import issues_for
from pptx_translate.structure import compare
from scripts.final_qa import check
from scripts.evaluate_model import evaluate


def make_deck(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    text = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    text.text = "The patient received 5 mg of the drug."
    table = slide.shapes.add_table(1, 2, Inches(1), Inches(2), Inches(8), Inches(1)).table
    table.cell(0, 0).text = "Total revenue"
    table.cell(0, 1).text = "18% growth"
    prs.save(path)


def test_commercial_gate_and_matching_report() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        source, output = root / "source.pptx", root / "translated.pptx"
        mapping = root / "translations.json"
        make_deck(source)
        mapping.write_text(json.dumps({
            "The patient received 5 mg of the drug.": "患者接受了 5 mg 该药物。",
            "Total revenue": "总收入",
            "18% growth": "增长 18%",
        }, ensure_ascii=False), "utf-8")
        code = main([str(source), "-o", str(output), "--strings-json", str(mapping),
                     "--qa-mode", "commercial", "--quiet"])
        assert code == 0
        report_path = Path(str(output) + ".report.json")
        assessment = check(source, output, report_path)
        assert assessment["status"] == "manual_review_required"
        assert not assessment["client_ready"]
        assert not assessment["blockers"]

        report = json.loads(report_path.read_text("utf-8"))
        report["output_sha256"] = "stale"
        report_path.write_text(json.dumps(report), "utf-8")
        assert "output_does_not_match_report" in check(source, output, report_path)["blockers"]


def test_commercial_gate_blocks_untranslated_text() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        source, output = root / "source.pptx", root / "translated.pptx"
        mapping = root / "incomplete.json"
        make_deck(source)
        mapping.write_text(json.dumps({"Total revenue": "总收入"}, ensure_ascii=False), "utf-8")
        code = main([str(source), "-o", str(output), "--strings-json", str(mapping),
                     "--qa-mode", "commercial", "--quiet"])
        assert code == 3
        report = json.loads(Path(str(output) + ".report.json").read_text("utf-8"))
        assert report["delivery_assessment"]["status"] == "blocked"
        assert "untranslated_strings" in report["delivery_assessment"]["blockers"]


def test_domain_measurement_must_survive_translation() -> None:
    changed = issues_for("Give 5 mg/kg for 14 days.", "连续14天给予5毫克/公斤。",
                         "Simplified Chinese")
    preserved = issues_for("Set torque to 45 N·m.", "将扭矩设为45 N·m。",
                           "Simplified Chinese")
    assert "measurement_units_changed" in changed
    assert "measurement_units_changed" not in preserved


def test_domain_benchmark_exposes_terminology_error() -> None:
    samples = [{"domain": "finance", "source": "Net margin fell to 6.2%.",
                "must_contain": ["净利率"], "must_preserve": ["6.2%"]}]
    result = evaluate(samples, {samples[0]["source"]: "净利润降至6.2%。"})
    assert result["passed_checks"] == 0
    assert result["results"][0]["missing_terms"] == ["净利率"]


def test_visual_structure_catches_shape_movement() -> None:
    with tempfile.TemporaryDirectory() as temp:
        source, moved = Path(temp) / "source.pptx", Path(temp) / "moved.pptx"
        make_deck(source)
        prs = Presentation(source)
        prs.slides[0].shapes[0].left += Inches(0.5)
        prs.save(moved)
        result = compare(source, moved)
        assert result["shape_set_differences"] == 0
        assert result["geometry_differences"] == 1


def test_commercial_mode_requires_explicit_provider() -> None:
    with tempfile.TemporaryDirectory() as temp:
        source = Path(temp) / "source.pptx"
        make_deck(source)
        assert main([str(source), "--qa-mode", "commercial", "--quiet"]) == 2
