# -*- coding: utf-8 -*-
"""Find which Hy-MT GGUF repos actually exist and list their files.

The user asked for `tencent/Hy-MT2-1.8B-GGUF`; several spellings and mirrors are
probed so the download step targets a repository that truly exists, with the real
Q4_K_M filename, instead of a guessed URL.
"""
import json
import urllib.request

CANDIDATES = [
    "tencent/Hy-MT2-1.8B-GGUF",
    "tencent/Hy-MT1.5-1.8B-GGUF",
    "tencent/Hy-MT1.5-1.8B",
    "tencent/Hy-MT2-1.8B",
    "zanish-labs/HY-MT1.5-1.8B-GGUF",
    "AngelSlim/Hy-MT1.5-1.8B-2bit-GGUF",
    "mradermacher/HY-MT1.5-1.8B-i1-GGUF",
    "mradermacher/HY-MT1.5-1.8B-GGUF",
    "demonbyron/HY-MT1.5-1.8B-GGUF",
    "tencent/Hunyuan-MT-7B-GGUF",
    "mradermacher/Hunyuan-MT-7B-GGUF",
]
BASES = ["https://hf-mirror.com", "https://huggingface.co"]


def fetch(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "pptx-translate/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def probe(repo):
    for base in BASES:
        try:
            raw = fetch(f"{base}/api/models/{repo}")
            data = json.loads(raw)
        except Exception as exc:
            print(f"  [{base}] {type(exc).__name__}: {str(exc)[:70]}")
            continue
        files = [s.get("rfilename"), s.get("size")] if False else [
            (s.get("rfilename"), s.get("size")) for s in data.get("siblings", [])
        ]
        print(f"  [OK via {base}] {len(files)} file(s)")
        for name, size in files:
            if name and (name.lower().endswith(".gguf") or name in ("config.json", "README.md")):
                size_txt = f"{size / 1048576:.0f} MB" if isinstance(size, int) else "?"
                print(f"      {name:52} {size_txt}")
        return True
    return False


print("=== probing candidate repositories ===")
found = []
for repo in CANDIDATES:
    print(f"\n-- {repo}")
    if probe(repo):
        found.append(repo)
print("\n=== repos that exist ===")
for repo in found:
    print("  ", repo)
