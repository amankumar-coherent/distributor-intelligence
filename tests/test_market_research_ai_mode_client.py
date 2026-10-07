"""ai_mode_client.py: Google AI Mode backend adapter.

No real browser is ever launched here — vendor_intel.scraping.google_ai_mode
.ask is mocked in every test, matching this repo's documented "a missed mock
must fail loudly, never silently drive a browser" rule (tests/conftest.py).
"""
from __future__ import annotations

import pytest
from unittest.mock import patch

from vendor_intel.pipeline.market_research import ai_mode_client


def _mock_ask(return_value):
    return patch("vendor_intel.scraping.google_ai_mode.ask", return_value=return_value)


def test_chat_returns_answer_text():
    with _mock_ask("Paris is the capital of France."):
        assert ai_mode_client.chat("sys", "user") == "Paris is the capital of France."


def test_chat_prepends_system_to_prompt():
    """AI Mode takes a single prompt string, so system+user must be merged
    into one — a dropped system prompt would silently lose the
    anti-fabrication and distributor-only instructions."""
    with _mock_ask("ok") as m:
        ai_mode_client.chat("SYSTEM_MARKER", "USER_MARKER")
    sent = m.call_args.args[0]
    assert "SYSTEM_MARKER" in sent
    assert "USER_MARKER" in sent


def test_chat_accepts_and_ignores_chatgpt_backend_kwargs():
    """Drop-in compatibility with chatgpt_oauth_client: model/port/max_tokens
    must be accepted (and ignored) so swapping backends is one import line."""
    with _mock_ask("ok"):
        assert ai_mode_client.chat(
            "sys", "user", max_tokens=6000, model="gpt-5.4-mini", port=10531
        ) == "ok"


def test_chat_raises_when_ai_mode_returns_empty():
    """google_ai_mode.ask() returns "" on CAPTCHA/timeout/empty. That must
    surface as AIModeUnavailable, not as a successful empty answer that
    downstream code would try to parse as companies."""
    with _mock_ask(""):
        with pytest.raises(ai_mode_client.AIModeUnavailable):
            ai_mode_client.chat("sys", "user")


def test_chat_json_parses_plain_array():
    with _mock_ask('[{"company": "Acme"}]'):
        assert ai_mode_client.chat_json("sys", "user") == [{"company": "Acme"}]


def test_chat_json_strips_markdown_fence_and_json_tag():
    """AI Mode's rendered code blocks come through as a bare "json" language
    tag (and sometimes ``` fences) before the payload — verified live."""
    raw = 'json\n[\n  {"company": "Ingram Micro", "founded": "1979"}\n]\n'
    with _mock_ask(raw):
        assert ai_mode_client.chat_json("sys", "user") == [
            {"company": "Ingram Micro", "founded": "1979"}
        ]


def test_chat_json_strips_google_use_code_with_caution_footer():
    """Google appends its own "Use code with caution." UI footer after code
    blocks — verified live in the real AI Mode answer."""
    raw = 'json\n[{"company": "Arrow Electronics"}]\n\nUse code with caution.\n'
    with _mock_ask(raw):
        assert ai_mode_client.chat_json("sys", "user") == [{"company": "Arrow Electronics"}]


def test_chat_json_slices_payload_out_of_surrounding_prose():
    """The answer routinely has prose before/after the JSON; slicing to the
    outermost brackets is what makes this parse at all."""
    raw = 'Here are 2 distributors:\n[{"company": "A"}, {"company": "B"}]\nHope that helps!'
    with _mock_ask(raw):
        assert ai_mode_client.chat_json("sys", "user") == [{"company": "A"}, {"company": "B"}]


def test_chat_json_falls_back_to_object_when_no_array():
    with _mock_ask('{"company": "Solo Co"}'):
        assert ai_mode_client.chat_json("sys", "user") == {"company": "Solo Co"}


def test_chat_json_raises_on_unparseable_answer():
    with _mock_ask("I could not find any companies matching that."):
        with pytest.raises(ai_mode_client.AIModeUnavailable):
            ai_mode_client.chat_json("sys", "user")


def test_ensure_enabled_defaults_ai_mode_on(monkeypatch):
    """The pipeline's discovery step depends entirely on AI Mode, so an
    unset GOOGLE_AI_MODE_ENABLED must default ON rather than silently
    returning nothing for every batch."""
    monkeypatch.delenv("GOOGLE_AI_MODE_ENABLED", raising=False)
    with _mock_ask("ok"):
        ai_mode_client.chat("sys", "user")
    import os

    assert os.environ["GOOGLE_AI_MODE_ENABLED"] == "true"


def test_ensure_enabled_respects_explicit_false(monkeypatch):
    """An explicit opt-out must not be silently overridden."""
    monkeypatch.setenv("GOOGLE_AI_MODE_ENABLED", "false")
    with _mock_ask("ok"):
        ai_mode_client.chat("sys", "user")
    import os

    assert os.environ["GOOGLE_AI_MODE_ENABLED"] == "false"
