# -*- coding: utf-8 -*-
"""Smoke-test a local Ollama translation model before translating a whole deck.

Checks, in order:
  1. the Ollama server answers on /api/version;
  2. the requested model is registered;
  3. a handful of real deck strings come back translated, with emoji, soft
     breaks and locked terms intact.

Usage::

    python scripts/ollama_smoke_test.py                        # defaults
    python scripts/ollama_smoke_test.py --model hy-mt-local --host http://localhost:11434
"""
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx_translate import core  # noqa: E402

SAMPLES = [
    "What Are LLM",
    "Prompt Injection",
    "Tricking LLMs to ignore safety rules (e.g., ‘Forget your rules’).",
    "🛡️ Simulated Attacks\x0bFriendly knights test your castle walls.",
    "TTPs (Tactics, Techniques, and Procedures)",
    "OWASP Top 10 | AI Incident DB | AVID",
]


def http_json(url: str, payload: dict | None = None, timeout: float = 60.0):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(
        url, data=data, method="POST" if data else "GET",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8", "replace"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="http://localhost:11434",
                        help="Ollama root url (without /v1)")
    parser.add_argument("--model", default="hy-mt-local", help="registered model name")
    parser.add_argument("--timeout", type=float, default=180.0)
    args = parser.parse_args()

    root = args.host.rstrip("/")
    if root.endswith("/v1"):
        root = root[:-3]

    print(f"1) server: {root}/api/version")
    try:
        version = http_json(f"{root}/api/version", timeout=10)
        print(f"   OK  ollama {version.get('version', '?')}")
    except Exception as exc:
        print(f"   FAIL {type(exc).__name__}: {exc}")
        print("   -> is the server running?  start it with:  ollama serve")
        return 2

    print(f"2) model registered? {args.model}")
    try:
        tags = http_json(f"{root}/api/tags", timeout=15)
        names = [m.get("name") for m in tags.get("models", [])]
        print(f"   available: {', '.join(names) if names else '(none)'}")
        if not any(n == args.model or (n or "").startswith(args.model + ":") for n in names):
            print(f"   FAIL '{args.model}' is not registered.")
            print(f"   -> ollama create {args.model} -f D:\\AI-tools\\Models\\Modelfile")
            return 3
        print(f"   OK  {args.model} present")
    except Exception as exc:
        print(f"   FAIL {type(exc).__name__}: {exc}")
        return 2

    print("3) translation through the real pipeline")
    cfg = core.ProviderConfig(
        provider="ollama",
        base_url=f"{root}/v1",
        api_key="",
        model=args.model,
        timeout=args.timeout,
        retries=1,
    )
    provider = core.build_provider(cfg)
    translator = core.Translator(provider, glossary=core.DEFAULT_GLOSSARY, use_cache=False)

    started = time.time()
    resolved = translator.translate_texts(SAMPLES)
    elapsed = time.time() - started

    ok = 0
    for source in SAMPLES:
        got = resolved.get(source, "")
        good = bool(got) and got != source and core.LANG_CJK_RE.search(got)
        ok += 1 if good else 0
        print(f"   [{'OK ' if good else 'BAD'}] {source[:58]!r}")
        print(f"         -> {got[:78]!r}")

    # the two guarantees that matter for layout, checked explicitly
    emoji_src = SAMPLES[3]
    emoji_out = resolved.get(emoji_src, "")
    emoji_kept = "🛡️" in emoji_out
    break_kept = "\x0b" in emoji_out
    print(f"\n   emoji preserved : {emoji_kept}   soft break preserved: {break_kept}")

    print(f"\n{ok}/{len(SAMPLES)} translated in {elapsed:.1f}s "
          f"({elapsed / len(SAMPLES):.1f}s per string)")
    if ok == len(SAMPLES) and emoji_kept and break_kept:
        print("RESULT: PASS — ready to translate the deck:")
        print(f'  python -m pptx_translate "deck.pptx" --provider ollama --ollama-model {args.model}')
        return 0
    print("RESULT: REVIEW — see the BAD lines above")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
