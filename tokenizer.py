"""Tokenizer wrapper for paper-aligned compression metrics.

The original encoder reported "tokens" using ``len(json.dumps(...))``, which
counts characters of a JSON string and is not comparable to the paper's
tokenizer-based ratios. This module provides a thin wrapper that uses
``tiktoken`` when available and a deterministic char-based fallback otherwise.
"""
from __future__ import annotations

import logging
from typing import Dict, Optional

logger = logging.getLogger(__name__)

try:  # pragma: no cover - import guard
    import tiktoken  # type: ignore

    _TIKTOKEN_AVAILABLE = True
except Exception:  # pragma: no cover - import guard
    tiktoken = None  # type: ignore
    _TIKTOKEN_AVAILABLE = False


DEFAULT_MODEL = "gpt-4"
_ENCODER_CACHE: dict = {}
_FALLBACK_CHARS_PER_TOKEN = 4
_FALLBACK_WARNED = False


def _get_encoder(model: str):
    if model in _ENCODER_CACHE:
        return _ENCODER_CACHE[model]
    try:
        enc = tiktoken.encoding_for_model(model)  # type: ignore[union-attr]
    except Exception:
        enc = tiktoken.get_encoding("cl100k_base")  # type: ignore[union-attr]
    _ENCODER_CACHE[model] = enc
    return enc


def is_tiktoken_available() -> bool:
    """Return ``True`` if ``tiktoken`` was importable."""
    return _TIKTOKEN_AVAILABLE


def tokenizer_metadata(model: Optional[str] = None, force_fallback: bool = False) -> Dict[str, object]:
    """Return JSON-serializable metadata for the tokenizer path.

    ``force_fallback`` is primarily for tests and for callers that need to
    record hypothetical fallback behavior without monkeypatching imports.
    """
    model_name = model or DEFAULT_MODEL
    fallback = force_fallback or not _TIKTOKEN_AVAILABLE
    return {
        "model": model_name,
        "backend": "char_approximation" if fallback else "tiktoken",
        "fallback": fallback,
        "fallback_chars_per_token": _FALLBACK_CHARS_PER_TOKEN if fallback else None,
    }


def count_tokens(text: str, model: Optional[str] = None) -> int:
    """Return the tokenizer token count for ``text``.

    Uses ``tiktoken``'s encoder for ``model`` when available, with a
    deterministic char-based approximation otherwise.
    """
    global _FALLBACK_WARNED
    if not isinstance(text, str):
        text = str(text)
    if _TIKTOKEN_AVAILABLE:
        try:
            return len(_get_encoder(model or DEFAULT_MODEL).encode(text))
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("tiktoken encode failed (%s); using char fallback", exc)

    if not _FALLBACK_WARNED:
        logger.warning(
            "tiktoken is not installed; compression ratios will use a "
            "char/%d approximation. pip install tiktoken for paper-aligned "
            "metrics.",
            _FALLBACK_CHARS_PER_TOKEN,
        )
        _FALLBACK_WARNED = True
    return max(1, len(text) // _FALLBACK_CHARS_PER_TOKEN) if text else 0
