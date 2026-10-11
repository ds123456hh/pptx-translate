# -*- coding: utf-8 -*-
"""Report which optional local-model dependencies are installed."""
import importlib

for name in ("torch", "transformers", "huggingface_hub", "sentencepiece", "tokenizers"):
    try:
        module = importlib.import_module(name)
        print("{:16} OK      {}".format(name, getattr(module, "__version__", "?")))
    except Exception:
        print("{:16} not installed".format(name))
