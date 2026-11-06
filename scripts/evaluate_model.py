"""Run a small cross-domain translation regression; no PPT files are changed.

This catches known terminology and literal-preservation mistakes. Passing is
not a semantic accuracy certificate or a substitute for domain review.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pptx_translate import core, quality  # noqa: E402


def evaluate(samples: list[dict], translations: dict[str, str]) -> dict:
    rows = []
    for sample in samples:
        source = sample["source"]
        target = translations.get(source, source)
        missing_terms = [term for term in sample.get("must_contain", []) if term not in target]
        missing_literals = [term for term in sample.get("must_preserve", []) if term not in target]
        warnings = quality.issues_for(source, target, "Simplified Chinese")
        rows.append({"domain": sample["domain"], "source": source, "translation": target,
                     "missing_terms": missing_terms, "missing_literals": missing_literals,
                     "quality_warnings": warnings,
                     "passed_checks": not (missing_terms or missing_literals or warnings)})
    return {"cases": len(rows), "passed_checks": sum(row["passed_checks"] for row in rows),
            "semantic_accuracy_verified": False, "results": rows}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--samples", type=Path, default=Path(__file__).resolve().parents[1] /
                        "benchmarks" / "general_en_zh.json")
    parser.add_argument("--host", default="http://localhost:11434/v1")
    parser.add_argument("--model", default="hy-mt-local")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    samples = json.loads(args.samples.read_text("utf-8"))
    cfg = core.ProviderConfig(provider="ollama", base_url=args.host, api_key="",
                              model=args.model, target_language="Simplified Chinese (简体中文)",
                              retries=1)
    translator = core.Translator(core.build_provider(cfg), glossary={}, use_cache=False)
    result = evaluate(samples, translator.translate_texts([item["source"] for item in samples]))
    output = args.output or Path(str(args.samples) + ".result.json")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
    print(f"{result['passed_checks']}/{result['cases']} passed mechanical checks; details: {output}")
    return 0 if result["passed_checks"] == result["cases"] else 3


if __name__ == "__main__":
    raise SystemExit(main())
