"""Check a translated PPTX/PPTM and its matching report before client review.

Usage:
    python scripts/final_qa.py source.pptx translated.pptx
    python scripts/final_qa.py source.pptx translated.pptx --report run.report.json

Exit 3 means machine-detectable blockers. Exit 0 means the files passed these
checks, *not* that linguistic accuracy or rendered layout has been approved.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pptx_translate import commercial_qa  # noqa: E402
from pptx_translate import structure  # noqa: E402
from pptx_translate.cli import verify_output  # noqa: E402

TOKEN = re.compile(rb"\[\[[A-Za-z]?\d+\]\]")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def leftover_tokens(path: Path) -> list[str]:
    hits: list[str] = []
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.endswith(".xml"):
                for token in sorted(set(TOKEN.findall(package.read(name)))):
                    hits.append(f"{name}: {token.decode('ascii', 'replace')}")
    return hits


def check(source: Path, output: Path, report_path: Path) -> dict:
    report = json.loads(report_path.read_text("utf-8"))
    assessment = commercial_qa.assess(report)
    blockers = assessment["blockers"]
    if report.get("input_sha256") != sha256(source):
        blockers.append("source_does_not_match_report")
    if report.get("output_sha256") != sha256(output):
        blockers.append("output_does_not_match_report")
    current = verify_output(source, output)
    for key in ("slides", "media_identical", "timing_same", "parts_added", "parts_removed"):
        if current.get(key) != (report.get("verification") or {}).get(key):
            blockers.append(f"verification_stale:{key}")
    if structure.compare(source, output) != report.get("visual_structure"):
        blockers.append("visual_structure_stale_or_missing")
    tokens = leftover_tokens(output)
    if tokens:
        blockers.append(f"unresolved_tokens_in_package:{len(tokens)}")
    assessment["status"] = "blocked" if blockers else "manual_review_required"
    assessment["unresolved_token_examples"] = tokens[:20]
    assessment["source"] = str(source)
    assessment["output"] = str(output)
    assessment["report"] = str(report_path)
    return assessment


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)
    report = args.report or Path(str(args.output) + ".report.json")
    for path in (args.source, args.output, report):
        if not path.is_file():
            parser.error(f"file not found: {path}")
    try:
        assessment = check(args.source, args.output, report)
    except (ValueError, zipfile.BadZipFile, OSError) as exc:
        parser.error(str(exc))
    print(json.dumps(assessment, ensure_ascii=False, indent=2))
    return 3 if assessment["blockers"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
