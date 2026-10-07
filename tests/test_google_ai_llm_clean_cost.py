"""Cost controls on the DeepSeek cleanup of Google AI Overview text.

The cleanup ran on EVERY Overview response and sent up to 12,000 chars each
time. Since every downstream consumer parses the result with regex rather
than reading prose, shortening an already-short answer bought nothing —
measured on live traffic, ~50% of calls were avoidable with no loss of
extraction quality.
"""
from __future__ import annotations

import pytest

from vendor_intel.clients import google_ai_scraper as gs


class TestSkipThreshold:
    def test_default_threshold(self, monkeypatch):
        monkeypatch.delenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", raising=False)
        assert gs._llm_clean_min_chars() == 900

    def test_configurable(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", "1500")
        assert gs._llm_clean_min_chars() == 1500

    def test_zero_disables_skipping(self, monkeypatch):
        """0 means "never skip" — every response goes to the LLM again."""
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", "0")
        assert gs._llm_clean_min_chars() == 0

    def test_garbage_falls_back_to_default(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", "not-a-number")
        assert gs._llm_clean_min_chars() == 900


class TestMaxInputChars:
    def test_default_is_far_below_the_old_12000(self, monkeypatch):
        monkeypatch.delenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MAX_CHARS", raising=False)
        assert gs._llm_clean_max_input_chars() == 4000

    def test_configurable(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MAX_CHARS", "8000")
        assert gs._llm_clean_max_input_chars() == 8000

    def test_has_a_floor_so_context_is_never_useless(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MAX_CHARS", "10")
        assert gs._llm_clean_max_input_chars() == 500


class TestShortResponsesSkipTheLlm:
    def test_short_text_is_left_alone(self, monkeypatch):
        """A response under the threshold must not be sent to DeepSeek."""
        import asyncio

        called = {"n": 0}

        async def fake_llm_clean(query, raw):
            called["n"] += 1
            return "should not be used"

        monkeypatch.setattr(gs, "llm_clean_scrape_markdown", fake_llm_clean)
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_ENABLED", "true")
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN", "true")
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", "900")

        data = {"markdown": "Acme is a distributor of medical devices.", "citations": []}
        out = asyncio.run(gs._apply_deepseek_sentence_clean("q", data, llm_clean=None))

        assert called["n"] == 0, "short text must not reach the LLM"
        assert out.get("llm_clean_skipped") == "short_enough"
        assert out["markdown"].startswith("Acme is a distributor")

    def test_long_text_still_gets_cleaned(self, monkeypatch):
        """Long, chrome-heavy pages are where the LLM measurably helps."""
        import asyncio

        called = {"n": 0}

        async def fake_llm_clean(query, raw):
            called["n"] += 1
            return "Acme is a distributor."

        monkeypatch.setattr(gs, "llm_clean_scrape_markdown", fake_llm_clean)
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_ENABLED", "true")
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN", "true")
        monkeypatch.setenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", "900")

        long_text = "Acme distributes devices. " * 100  # ~2600 chars
        data = {"markdown": long_text, "citations": []}
        out = asyncio.run(gs._apply_deepseek_sentence_clean("q", data, llm_clean=None))

        assert called["n"] == 1
        assert out.get("llm_cleaned") is True
