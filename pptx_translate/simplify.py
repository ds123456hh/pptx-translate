# -*- coding: utf-8 -*-
"""Traditional -> Simplified conversion for model output.

The local MT model occasionally emits Traditional forms (規則, 嚴格, 無, 誤解…).
Two layers handle that:

1. **OpenCC** (``pip install opencc-python-reimplemented``) when present — the
   standard converter, covering the whole character set.
2. A small built-in table for the handful of characters this model actually
   produces, so the default install is not broken without the extra package.

Explicit conversion beats asking the model nicely: instructions to "use
Simplified" reduced but did not eliminate Traditional output.
"""

from __future__ import annotations

import logging
import threading

log = logging.getLogger("pptx_translate.simplify")

# Characters seen in this model's output, plus the common high-frequency set.
_FALLBACK = {
    "誤": "误", "規": "规", "則": "则", "嚴": "严", "範": "范", "圍": "围",
    "無": "无", "亂": "乱", "線": "线", "隊": "队", "擊": "击", "體": "体",
    "應": "应", "這": "这", "個": "个", "們": "们", "來": "来", "說": "说",
    "對": "对", "時": "时", "會": "会", "實": "实", "現": "现", "點": "点",
    "類": "类", "數": "数", "據": "据", "網": "网", "絡": "络", "為": "为",
    "與": "与", "從": "从", "進": "进", "過": "过", "開": "开", "關": "关",
    "發": "发", "場": "场", "題": "题", "計": "计", "劃": "划", "學": "学",
    "習": "习", "國": "国", "語": "语", "電": "电", "話": "话", "報": "报",
    "導": "导", "護": "护", "險": "险", "組": "组", "織": "织", "構": "构",
    "設": "设", "備": "备", "統": "统", "權": "权", "證": "证", "驗": "验",
    "錯": "错", "讀": "读", "寫": "写", "選": "选", "擇": "择", "風": "风",
    "雲": "云", "庫": "库", "儲": "储", "執": "执", "業": "业", "務": "务",
    "員": "员", "團": "团", "領": "领", "訓": "训", "練": "练", "錢": "钱",
    "財": "财", "費": "费", "買": "买", "賣": "卖", "價": "价", "當": "当",
    "還": "还", "讓": "让", "認": "认", "識": "识", "講": "讲", "書": "书",
    "筆": "笔", "產": "产", "稱": "称", "標": "标", "籤": "签", "擊": "击",
}

_lock = threading.Lock()
_converter = None
_checked = False


def _opencc():
    """The OpenCC converter, or None when the package is not installed."""
    global _converter, _checked
    if _checked:
        return _converter
    with _lock:
        if _checked:
            return _converter
        try:
            from opencc import OpenCC

            _converter = OpenCC("t2s")
            log.info("simplify: using OpenCC (t2s)")
        except Exception:
            _converter = None
            log.info("simplify: OpenCC not installed; using the built-in table "
                     "(install with: pip install opencc-python-reimplemented)")
        _checked = True
    return _converter


def to_simplified(text: str) -> str:
    """Convert Traditional characters to Simplified. Safe for any input."""
    if not text:
        return text
    converter = _opencc()
    if converter is not None:
        try:
            return converter.convert(text)
        except Exception as exc:  # fall through to the table
            log.debug("OpenCC failed (%s); using the built-in table", exc)
    if any(c in _FALLBACK for c in text):
        return "".join(_FALLBACK.get(c, c) for c in text)
    return text


def has_traditional(text: str) -> bool:
    """Whether any character in ``text`` has a distinct Simplified form."""
    return any(c in _FALLBACK for c in text)
