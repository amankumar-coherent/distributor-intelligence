from __future__ import annotations

import os
from typing import Any

from vendor_intel.config import Settings
from vendor_intel.placeholders import llm as ph
from vendor_intel.placeholders.claude import DEFAULT_COMPILER_MODEL


class ClaudeClient:
    """LLM client (Anthropic / Gemini / Groq / OpenCode Zen via LLM_PROVIDER)."""

    def __init__(self, settings: Settings):
        self._settings = settings
        self._model = settings.compiler_model

    @property
    def available(self) -> bool:
        if self._settings.use_mock_data or self._settings.mock_mode:
            return False
        return ph.is_configured()

    def _resolve_model(self, model: str | None = None) -> str:
        """Route to provider-specific model (OpenAI / OpenCode / Gemini / Groq)."""
        if self._settings.llm_provider == "openai":
            req = (model or "").strip()
            # Ignore leftover OpenCode/Anthropic model ids from yaml defaults
            if req and not any(
                x in req.lower() for x in ("claude", "deepseek", "big-pickle", "ling-3", "free")
            ):
                return req
            return (
                getattr(self._settings, "openai_model", None)
                or os.getenv("OPENAI_MODEL")
                or "gpt-4o-mini"
            ).strip() or "gpt-4o-mini"
        if self._settings.llm_provider == "opencode":
            req = (model or "").strip()
            if req and "claude" not in req.lower():
                return req
            return self._settings.opencode_model
        if self._settings.llm_provider == "chatgpt_scraper":
            return "chatgpt-web-guest"
        if model:
            return model
        if self._settings.llm_provider == "gemini":
            return self._settings.gemini_model
        if self._settings.llm_provider == "groq":
            return self._settings.groq_model
        return self._model

    def complete(self, system: str, user: str, model: str | None = None, max_tokens: int = 8192) -> str:
        return ph.llm_complete(
            system, user, model=self._resolve_model(model), max_tokens=max_tokens
        )

    def complete_json(
        self,
        system: str,
        user: str,
        model: str | None = None,
        *,
        max_tokens: int = 4096,
    ) -> Any:
        return ph.llm_complete_json(
            system, user, model=self._resolve_model(model), max_tokens=max_tokens
        )

    def complete_json_with_usage(
        self,
        system: str,
        user: str,
        model: str | None = None,
        *,
        max_tokens: int = 4096,
    ) -> tuple[Any, dict[str, Any]]:
        """Like complete_json, but also returns the real token-usage dict
        ({"input_tokens", "output_tokens", "provider", "model"}) for callers
        that record real LLM cost (pipeline/llm_meter.py's record())."""
        return ph.llm_complete_json_with_usage(
            system, user, model=self._resolve_model(model), max_tokens=max_tokens
        )
