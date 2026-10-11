# -*- coding: utf-8 -*-
"""Final QA on a translated deck: text integrity, tokens, alignment, samples."""
import argparse
import json
import re
import sys
import zipfile

from pptx import Presentation

from pptx_translate import core

SRC = r"D:\大模型渗透测试：大语言模型安全防护教程\1 - Introduction\1 -LLM PENTEST.pptx"
DST = sys.argv[1] if len(sys.argv) > 1 else r"D:\AI-tools\out\zh-final.pptx"
REPORT = sys.argv[2] if len(sys.argv) > 2 else r"D:\AI-tools\out\zh-final.report.json"

report = json.load(open(REPORT, encoding="utf-8"))
print("report: translated=%s untranslated=%s api_failures=%s"
      % (report.get("translated_strings"), report.get("untranslated_strings"),
         report.get("api_failures")))
print("still English:", report.get("untranslated"))

print("\n--- leftover placeholder tokens in the package ---")
with zipfile.ZipFile(DST) as zf:
    bad = []
    for name in zf.namelist():
        if not name.endswith(".xml"):
            continue
        data = zf.read(name).decode("utf-8", "replace")
        for tok in set(re.findall(r"\[\[[A-Za-z]?\d+\]\]", data)):
            bad.append((name, tok))
    print("  leftover tokens:", len(bad))
    for item in bad[:12]:
        print("   ", item)

print("\n--- paragraph alignment source vs translated ---")
src_prs, dst_prs = Presentation(SRC), Presentation(DST)
def units(prs):
    out = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                t = "".join(r.text for r in p.runs)
                if t.strip():
                    out.append((i, sh.name, t))
    return out
a, b = units(src_prs), units(dst_prs)
print(f"  paragraphs: source {len(a)} vs translated {len(b)}", "OK" if len(a) == len(b) else "MISMATCH")
same = [(i, n, t) for (i, n, t), (_j, _m, u) in zip(a, b) if t == u]
print(f"  identical to source: {len(same)}")
for row in same[:8]:
    print("   ", row[0], repr(row[2][:70]))

print("\n--- soft-break structure (slide 3 style shapes) ---")
shown = 0
for i, s in enumerate(dst_prs.slides, 1):
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        groups = [g for p in sh.text_frame.paragraphs for g in core._paragraph_lines(p)]
        if len(groups) >= 2:
            lines = ["".join(n.text or "" for n in g) for g in groups]
            if any("\x0b" in x for x in lines) or True:
                print(f"  slide {i} {sh.name}: {len(lines)} lines")
                for ln in lines[:4]:
                    print(f"     | {ln}")
                shown += 1
                break
    if shown >= 2:
        break

print("\n--- sample translations (slide 1, 2) ---")
for i, s in enumerate(dst_prs.slides, 1):
    if i > 2:
        break
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            for p in sh.text_frame.paragraphs[:6]:
                t = "".join(r.text for r in p.runs)
                if t.strip():
                    print(f"  s{i} | {t}")
