"""LLM router usage-surfacing (src/vendor_intel/placeholders/llm.py) —
verifies llm_complete_with_usage/llm_complete_json_with_usage extract real
token counts per provider, and that the legacy llm_complete/llm_complete_json
functions are unaffected (same text-only behavior as before the refactor).
All HTTP is mocked via llm._post_json — no real network calls.
"""
from __future__ import annotations

import httpx

from vendor_intel.placeholders import llm


def test_anthropic_usage_extraction(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "anthropic")
    monkeypatch.setattr(llm, "ANTHROPIC_API_KEY", "sk-ant-real")

    def fake_post_json(url, headers, body):
        assert "anthropic.com" in url
        return {
            "content": [{"type": "text", "text": '{"ok": true}'}],
            "usage": {"input_tokens": 123, "output_tokens": 45},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)

    text, usage = llm.llm_complete_with_usage("sys", "user")
    assert text == '{"ok": true}'
    assert usage == {
        "input_tokens": 123,
        "output_tokens": 45,
        "provider": "anthropic",
        "model": llm.DEFAULT_COMPILER_MODEL,
    }

    # Legacy function unaffected — same text, no usage in the return value.
    assert llm.llm_complete("sys", "user") == '{"ok": true}'


def test_openai_usage_extraction(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "openai")
    monkeypatch.setenv("OPENAI_API_KEY", "sk-real")

    def fake_post_json(url, headers, body):
        assert "chat/completions" in url
        return {
            "choices": [{"message": {"content": '{"a": 1}'}}],
            "usage": {"prompt_tokens": 500, "completion_tokens": 80, "total_tokens": 580},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)

    text, usage = llm.llm_complete_with_usage("sys", "user", model="gpt-4o-mini")
    assert text == '{"a": 1}'
    assert usage == {"input_tokens": 500, "output_tokens": 80, "provider": "openai", "model": "gpt-4o-mini"}


def test_groq_usage_extraction(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "groq")
    monkeypatch.setattr(llm, "GROQ_API_KEY", "gsk-real")

    def fake_post_json(url, headers, body):
        return {
            "choices": [{"message": {"content": "ok"}}],
            "usage": {"prompt_tokens": 10, "completion_tokens": 5},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)
    _, usage = llm.llm_complete_with_usage("sys", "user")
    assert usage["input_tokens"] == 10
    assert usage["output_tokens"] == 5
    assert usage["provider"] == "groq"


def test_gemini_usage_extraction(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "gemini")
    monkeypatch.setattr(llm, "GEMINI_API_KEY", "real-key")

    def fake_post_json(url, headers, body):
        return {
            "candidates": [{"content": {"parts": [{"text": "hello"}]}}],
            "usageMetadata": {"promptTokenCount": 7, "candidatesTokenCount": 3},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)
    text, usage = llm.llm_complete_with_usage("sys", "user")
    assert text == "hello"
    assert usage == {
        "input_tokens": 7,
        "output_tokens": 3,
        "provider": "gemini",
        "model": llm.DEFAULT_GEMINI_MODEL,
    }


def test_opencode_usage_extraction(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "opencode")
    monkeypatch.setattr(llm, "OPENCODE_API_KEY", "oc-real")

    def fake_post_json(url, headers, body):
        return {
            "choices": [{"message": {"content": '{"x":1}'}, "finish_reason": "stop"}],
            "usage": {"prompt_tokens": 20, "completion_tokens": 8},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)
    text, usage = llm.llm_complete_with_usage("sys", "user")
    assert text == '{"x":1}'
    assert usage["provider"] == "opencode"
    assert usage["input_tokens"] == 20
    assert usage["output_tokens"] == 8


def test_anthropic_quota_fallback_reports_opencode_usage(monkeypatch):
    """When Anthropic hits a quota error and falls back to OpenCode, the
    returned usage/provider/model must reflect the FALLBACK call, not the
    failed Anthropic attempt — otherwise cost tracking would misattribute
    the spend to the wrong provider."""
    monkeypatch.setattr(llm, "LLM_PROVIDER", "anthropic")
    monkeypatch.setattr(llm, "ANTHROPIC_API_KEY", "sk-ant-real")
    monkeypatch.setattr(llm, "OPENCODE_API_KEY", "oc-real")
    monkeypatch.setattr(llm, "LLM_FALLBACK_TO_OPENCODE", True)

    request = httpx.Request("POST", "https://api.anthropic.com/v1/messages")
    response = httpx.Response(429, request=request, text="rate limit exceeded, quota hit")
    calls = {"n": 0}

    def fake_post_json(url, headers, body):
        calls["n"] += 1
        if "anthropic.com" in url:
            raise httpx.HTTPStatusError("429", request=request, response=response)
        return {
            "choices": [{"message": {"content": "fallback text"}, "finish_reason": "stop"}],
            "usage": {"prompt_tokens": 11, "completion_tokens": 4},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)

    text, usage = llm.llm_complete_with_usage("sys", "user", model="claude-sonnet-4-6")
    assert text == "fallback text"
    assert usage["provider"] == "opencode"
    assert usage["input_tokens"] == 11
    assert usage["output_tokens"] == 4
    assert calls["n"] == 2  # anthropic attempt + opencode fallback


def test_error_response_has_zero_usage(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "openai")
    monkeypatch.setenv("OPENAI_API_KEY", "sk-real")

    request = httpx.Request("POST", "https://api.openai.com/v1/chat/completions")
    response = httpx.Response(500, request=request, text="server error")

    def fake_post_json(url, headers, body):
        raise httpx.HTTPStatusError("500", request=request, response=response)

    monkeypatch.setattr(llm, "_post_json", fake_post_json)
    text, usage = llm.llm_complete_with_usage("sys", "user")
    assert "llm_failed" in text
    assert usage == {"input_tokens": 0, "output_tokens": 0, "provider": "openai", "model": "gpt-4o-mini"}


def test_json_with_usage_parses_and_returns_both(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "openai")
    monkeypatch.setenv("OPENAI_API_KEY", "sk-real")

    def fake_post_json(url, headers, body):
        return {
            "choices": [{"message": {"content": '{"is_relevant": true}'}}],
            "usage": {"prompt_tokens": 900, "completion_tokens": 150},
        }

    monkeypatch.setattr(llm, "_post_json", fake_post_json)
    parsed, usage = llm.llm_complete_json_with_usage("sys", "user")
    assert parsed == {"is_relevant": True}
    assert usage["input_tokens"] == 900
    assert usage["output_tokens"] == 150

    # Legacy llm_complete_json must still work identically.
    parsed2 = llm.llm_complete_json("sys", "user")
    assert parsed2 == {"is_relevant": True}


def test_mock_mode_returns_zero_usage(monkeypatch):
    monkeypatch.setattr(llm, "LLM_PROVIDER", "mock")
    text, usage = llm.llm_complete_with_usage("sys", "user")
    assert usage == {"input_tokens": 0, "output_tokens": 0, "provider": "mock", "model": ""}
