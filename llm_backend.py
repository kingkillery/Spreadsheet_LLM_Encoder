"""LLM backend protocol and reference adapters for SpreadsheetLLM.

The Chain-of-Spreadsheet pipeline needs an LLM. This module provides a
minimal protocol any callable can satisfy plus three concrete adapters:

* ``EchoBackend`` — deterministic test fixture.
* ``CallableBackend`` — wraps any ``Callable[[str], str]``.
* ``OpenAIBackend`` — adapter for the official ``openai`` Python SDK
  (chat-completions style). Only imported when actually used so the package
  remains an optional dependency.
"""
from __future__ import annotations

import logging
from typing import Any, Callable, List, Optional, Protocol, runtime_checkable

logger = logging.getLogger(__name__)


@runtime_checkable
class LLMBackend(Protocol):
    """Anything callable as ``backend(prompt: str) -> str``."""

    def __call__(self, prompt: str) -> str:  # pragma: no cover - protocol
        ...


class EchoBackend:
    """Deterministic test backend.

    Returns ``response`` for every call (or cycles through ``responses`` if
    a list is provided). Records every prompt seen on ``calls``.
    """

    def __init__(
        self,
        response: str = "",
        responses: Optional[List[str]] = None,
    ) -> None:
        self.response = response
        self.responses = list(responses) if responses else None
        self.calls: List[str] = []
        self._idx = 0

    def __call__(self, prompt: str) -> str:
        self.calls.append(prompt)
        if self.responses:
            out = self.responses[self._idx % len(self.responses)]
            self._idx += 1
            return out
        return self.response


class CallableBackend:
    """Adapt a plain callable to the ``LLMBackend`` protocol."""

    def __init__(self, fn: Callable[[str], str]) -> None:
        self.fn = fn

    def __call__(self, prompt: str) -> str:
        return self.fn(prompt)


class OpenAIBackend:
    """Adapter for the OpenAI Python SDK (``openai>=1.0``)."""

    def __init__(
        self,
        model: str = "gpt-4o-mini",
        system_prompt: Optional[str] = None,
        client: Any = None,
        **completion_kwargs: Any,
    ) -> None:
        self.model = model
        self.system_prompt = system_prompt
        self._client = client
        self._completion_kwargs = completion_kwargs

    def _ensure_client(self) -> Any:
        if self._client is not None:
            return self._client
        try:
            from openai import OpenAI  # type: ignore
        except ImportError as exc:  # pragma: no cover - exercised only when openai missing
            raise RuntimeError(
                "openai package is not installed; pip install openai or "
                "supply your own backend."
            ) from exc
        self._client = OpenAI()
        return self._client

    def __call__(self, prompt: str) -> str:
        client = self._ensure_client()
        messages: List[dict] = []
        if self.system_prompt:
            messages.append({"role": "system", "content": self.system_prompt})
        messages.append({"role": "user", "content": prompt})
        resp = client.chat.completions.create(
            model=self.model, messages=messages, **self._completion_kwargs
        )
        choice = resp.choices[0]
        content = getattr(choice.message, "content", None)
        return content or ""


__all__ = [
    "LLMBackend",
    "EchoBackend",
    "CallableBackend",
    "OpenAIBackend",
]
