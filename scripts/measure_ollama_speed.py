# -*- coding: utf-8 -*-
"""Measure local Ollama inference speed in detail, to explain the per-string cost."""
import json
import time
import urllib.request

ROOT = "http://localhost:11434"


def post(path, payload, timeout=600):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(ROOT + path, data=data,
                                 headers={"Content-Type": "application/json"})
    started = time.time()
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = json.loads(resp.read().decode("utf-8", "replace"))
    return body, time.time() - started


print("=== model residency (/api/ps) ===")
try:
    with urllib.request.urlopen(ROOT + "/api/ps", timeout=15) as resp:
        ps = json.loads(resp.read().decode())
    for m in ps.get("models", []):
        print(f"  {m.get('name')}  size={m.get('size',0)/2**30:.2f} GiB  "
              f"vram={m.get('size_vram',0)/2**30:.2f} GiB  expires={m.get('expires_at')}")
    if not ps.get("models"):
        print("  (nothing loaded — the first request will pay a cold load)")
except Exception as exc:
    print("  unavailable:", exc)

print("\n=== raw /api/generate with timings ===")
payload = {
    "model": "hy-mt-local",
    "prompt": "将下面的文本翻译成中文，只输出译文：Prompt Injection is a common attack against LLM applications.",
    "stream": False,
    "options": {"temperature": 0},
}
body, wall = post("/api/generate", payload)
ns = lambda k: body.get(k, 0) / 1e9  # noqa: E731
print(f"  wall clock       : {wall:.1f}s")
print(f"  load_duration    : {ns('load_duration'):.1f}s   (model load; 0 if already resident)")
print(f"  prompt_eval      : {ns('prompt_eval_duration'):.1f}s for {body.get('prompt_eval_count')} tokens")
print(f"  eval (generate)  : {ns('eval_duration'):.1f}s for {body.get('eval_count')} tokens")
if body.get("eval_duration"):
    print(f"  generation speed : {body['eval_count'] / (body['eval_duration'] / 1e9):.1f} tokens/s")
print(f"  output           : {body.get('response', '')[:90]!r}")

print("\n=== second identical call (should skip the load) ===")
body2, wall2 = post("/api/generate", payload)
print(f"  wall clock       : {wall2:.1f}s   load_duration: {body2.get('load_duration',0)/1e9:.1f}s")
if body2.get("eval_duration"):
    print(f"  generation speed : {body2['eval_count'] / (body2['eval_duration'] / 1e9):.1f} tokens/s")

print("\n=== a short paragraph (typical slide bullet) ===")
payload3 = dict(payload, prompt="将下面的文本翻译成中文，只输出译文：Why LLMs Are Vulnerable")
body3, wall3 = post("/api/generate", payload3)
print(f"  wall clock       : {wall3:.1f}s for {body3.get('eval_count')} output tokens")
print(f"  output           : {body3.get('response', '')[:60]!r}")
