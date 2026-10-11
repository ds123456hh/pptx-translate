# -*- coding: utf-8 -*-
"""Local, in-process translation with a small HuggingFace translation model.

Why this exists
---------------
An 0.1B-class translation model runs comfortably on CPU, costs nothing per
string, and keeps the deck inside the machine — no slide text leaves the host.
That makes a full 500-string deck free to re-translate as often as needed.

Model
-----
Default is `DALabCommunity/Haidass-Translate-143M
<https://huggingface.co/DALabCommunity/Haidass-Translate-143M>`_, a ~143M
parameter Chinese/English translation model. The instruction-tuned sibling is
selected by default because this pipeline talks to the model in instructions
(keep these terms in English, preserve these tokens); pass
``--local-model`` to use another repo.

Design notes
------------
* The model is loaded lazily, once, and reused for the whole run.
* ``prompt_style`` exists because small MT models differ in what they expect:
  ``instruction`` (default) wraps the text in a translate command, ``chat`` uses
  the tokenizer's own chat template when it has one, and ``direct`` feeds the
  bare text — which is what a pure text-to-text MT checkpoint wants.
* Every other layer of the pipeline (term locking, emoji tokens, soft breaks,
  JSON contracts) is bypassed here on purpose: those exist to constrain a remote
  chat model through a lossy text protocol. Locally we pass strings straight in.
"""

from __future__ import annotations

import logging
import os
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

log = logging.getLogger("pptx_translate.local")

# Re-exported from core so there is a single source of truth for the default.
from .core import DEFAULT_LOCAL_MODEL  # noqa: E402

INSTRUCTION_MODEL = "umeiko/Haidass-Translate-143M-Instruction"


@dataclass
class LocalConfig:
    """Everything the local provider needs to load and drive a model."""

    model_id: str = DEFAULT_LOCAL_MODEL
    device: str = "auto"          # auto | cpu | cuda | cuda:0
    dtype: str = "auto"           # auto | float32 | float16 | bfloat16
    prompt_style: str = "auto"    # auto | instruction | chat | direct
    target_language: str = "中文"
    max_input_tokens: int = 480
    max_new_tokens: int = 400
    num_beams: int = 1
    cache_dir: str | None = None
    local_files_only: bool = False
    trust_remote_code: bool = False
    extra: dict = field(default_factory=dict)


def hf_endpoint() -> str:
    """The HuggingFace endpoint in use (honours ``HF_ENDPOINT`` / ``HF_MIRROR``)."""
    return (
        os.environ.get("HF_ENDPOINT")
        or os.environ.get("HF_MIRROR")
        or "https://huggingface.co"
    )


def is_available() -> tuple[bool, str]:
    """Whether the local stack can run, and why not when it cannot."""
    missing = []
    for module in ("torch", "transformers"):
        try:
            __import__(module)
        except Exception:
            missing.append(module)
    if missing:
        return False, "missing dependency: " + ", ".join(missing)
    return True, "ok"


class LocalProvider:
    """A translation provider backed by a locally loaded HF seq2seq model."""

    name = "local"

    def __init__(self, cfg, local: LocalConfig | None = None) -> None:
        # ``cfg`` is the shared ProviderConfig; only target_language is reused.
        self.cfg = cfg
        self.local = local or LocalConfig()
        self.system = None          # unused locally, kept for interface parity
        self.failed_items: list[str] = []
        self.lost_placeholders: list[str] = []
        self._lock = threading.Lock()
        self._model = None
        self._tokenizer = None
        self._device = "cpu"
        self._style = self.local.prompt_style
        self.extra_stats = {"calls": 0, "generated_tokens": 0}

        if not self.local.target_language or self.local.target_language == "中文":
            self.local.target_language = self._target_from_shared(cfg)

    @staticmethod
    def _target_from_shared(cfg) -> str:
        text = (getattr(cfg, "target_language", "") or "").lower()
        if "chinese" in text or "中文" in text:
            return "中文"
        if "english" in text or "英文" in text:
            return "English"
        return getattr(cfg, "target_language", "中文") or "中文"

    # -- loading ------------------------------------------------------------ #
    def load(self) -> None:
        """Load the tokenizer and model once (idempotent, thread-safe)."""
        if self._model is not None:
            return
        with self._lock:
            if self._model is not None:
                return
            ok, why = is_available()
            if not ok:
                raise RuntimeError(
                    f"local translation needs torch and transformers ({why}). "
                    "Install them with:  pip install -e .[local]"
                )
            import torch
            from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

            endpoint = hf_endpoint()
            log.info("loading local model %s (endpoint %s)", self.local.model_id, endpoint)
            kwargs = {"trust_remote_code": self.local.trust_remote_code}
            if self.local.cache_dir:
                kwargs["cache_dir"] = self.local.cache_dir
            if self.local.local_files_only:
                kwargs["local_files_only"] = True

            self._tokenizer = AutoTokenizer.from_pretrained(self.local.model_id, **kwargs)
            model_cls = AutoModelForSeq2SeqLM
            try:
                self._model = model_cls.from_pretrained(self.local.model_id, **kwargs)
            except Exception as exc:  # some checkpoints are causal LMs
                log.warning("seq2seq load failed (%s); trying causal LM", exc)
                from transformers import AutoModelForCausalLM

                self._model = AutoModelForCausalLM.from_pretrained(self.local.model_id, **kwargs)

            self._device = self._resolve_device(torch)
            dtype = self._resolve_dtype(torch)
            if dtype is not None:
                self._model = self._model.to(dtype)
            self._model = self._model.to(self._device)
            self._model.eval()
            if self._tokenizer.pad_token_id is None:
                self._tokenizer.pad_token = self._tokenizer.eos_token
            self._style = self._resolve_style()
            log.info("local model ready: device=%s dtype=%s style=%s",
                     self._device, dtype or "default", self._style)

    def _resolve_device(self, torch):
        want = (self.local.device or "auto").lower()
        if want == "auto":
            return "cuda" if torch.cuda.is_available() else "cpu"
        if want.startswith("cuda") and not torch.cuda.is_available():
            log.warning("device %s requested but CUDA is unavailable; using cpu", want)
            return "cpu"
        return want

    def _resolve_dtype(self, torch):
        want = (self.local.dtype or "auto").lower()
        if want == "auto":
            return None  # let the checkpoint decide
        return {
            "float32": torch.float32,
            "float16": torch.float16,
            "bfloat16": torch.bfloat16,
        }.get(want)

    def _resolve_style(self) -> str:
        want = (self.local.prompt_style or "auto").lower()
        if want != "auto":
            return want
        # an instruction-tuned checkpoint wants a command; a bare MT checkpoint
        # wants the text alone
        if "instruction" in self.local.model_id.lower() or "instruct" in self.local.model_id.lower():
            return "instruction"
        if getattr(self._tokenizer, "chat_template", None):
            return "chat"
        return "direct"

    @property
    def loaded(self) -> bool:
        return self._model is not None

    # -- prompting ---------------------------------------------------------- #
    def _instruction(self) -> str:
        return (
            f"将下面的文本翻译成{self.local.target_language}，"
            "只输出译文，不要解释，不要添加任何额外内容：\n"
        )

    def _build_inputs(self, texts: Sequence[str]):
        """Tokenize a batch according to the resolved prompt style."""
        tok = self._tokenizer
        if self._style == "chat" and getattr(tok, "chat_template", None):
            prompts = []
            for text in texts:
                messages = [
                    {"role": "system", "content": self.cfg.system or ""},
                    {"role": "user", "content": self._instruction() + text},
                ]
                prompts.append(tok.apply_chat_template(messages, tokenize=False, add_generation_prompt=True))
            return tok(prompts, return_tensors="pt", padding=True, truncation=True,
                       max_length=self.local.max_input_tokens)
        if self._style == "instruction":
            prompts = [self._instruction() + text for text in texts]
        else:
            prompts = list(texts)
        return tok(prompts, return_tensors="pt", padding=True, truncation=True,
                   max_length=self.local.max_input_tokens)

    # -- translation -------------------------------------------------------- #
    def translate(self, items: Sequence[str]) -> list[str]:
        self.load()
        import torch

        batch = self._build_inputs(list(items))
        batch = {k: v.to(self._device) for k, v in batch.items()}
        gen_kwargs = {
            "max_new_tokens": self.local.max_new_tokens,
            "num_beams": max(1, self.local.num_beams),
            "do_sample": False,
        }
        with torch.inference_mode():
            output = self._model.generate(**batch, **gen_kwargs)
        prompt_len = batch["input_ids"].shape[1]
        decoded = self._tokenizer.batch_decode(output[:, prompt_len:], skip_special_tokens=True)
        self.extra_stats["calls"] += 1
        self.extra_stats["generated_tokens"] += int(output.shape[1] - prompt_len) * len(items)
        return [text.strip() for text in decoded]

    def warmup(self, sample: str = "Hello world") -> str:
        """Load and translate one string, so mistakes surface before a full run."""
        return self.translate([sample])[0]


def build_local_provider(cfg, model_id: str | None = None, device: str = "auto",
                         prompt_style: str = "auto", cache_dir: str | None = None,
                         max_new_tokens: int = 400, local_files_only: bool = False,
                         target_language: str | None = None):
    """Construct a LocalProvider from CLI-level options."""
    local = LocalConfig(
        model_id=model_id or DEFAULT_LOCAL_MODEL,
        device=device,
        prompt_style=prompt_style,
        cache_dir=cache_dir,
        max_new_tokens=max_new_tokens,
        local_files_only=local_files_only,
    )
    if target_language:
        local.target_language = target_language
    return LocalProvider(cfg, local)


def default_cache_dir() -> Path:
    """Where downloaded weights live by default."""
    return Path(os.environ.get("HF_HOME") or Path.home() / ".cache" / "huggingface")
