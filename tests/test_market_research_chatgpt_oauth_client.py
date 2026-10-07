"""chatgpt_oauth_client.py: proxy health-check + ChatGPT-only retry behavior.

No real network call or subprocess is ever allowed here — httpx.get and
subprocess.Popen are mocked in every test, matching this repo's documented
"a missed mock must fail loudly, never silently hit the network" rule
(see tests/conftest.py).

Per explicit user instruction, this client no longer falls back to a
different LLM_PROVIDER (e.g. DeepSeek) on a ChatGPT failure — every company
must come exclusively from the user's own ChatGPT account. A failure is
retried against ChatGPT itself instead (see _chat_with_retry).
"""
from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from vendor_intel.pipeline.market_research import chatgpt_oauth_client as client


def test_is_healthy_true_on_200():
    resp = MagicMock(status_code=200)
    with patch("httpx.get", return_value=resp) as mock_get:
        assert client.is_healthy("http://127.0.0.1:10531/v1") is True
        mock_get.assert_called_once()


def test_is_healthy_false_on_exception():
    with patch("httpx.get", side_effect=ConnectionError("refused")):
        assert client.is_healthy("http://127.0.0.1:10531/v1") is False


def test_ensure_running_skips_spawn_when_already_healthy():
    with patch.object(client, "is_healthy", return_value=True), patch(
        "subprocess.Popen"
    ) as mock_popen:
        assert client.ensure_running(port=10531) is True
        mock_popen.assert_not_called()


def test_ensure_running_returns_false_without_npx():
    with patch.object(client, "is_healthy", return_value=False), patch.object(
        client, "_npx_path", return_value=None
    ), patch("subprocess.Popen") as mock_popen:
        assert client.ensure_running(port=10531, timeout=1) is False
        mock_popen.assert_not_called()


def test_ensure_running_spawns_and_polls_until_healthy():
    health_calls = {"n": 0}

    def fake_healthy(base_url):
        health_calls["n"] += 1
        return health_calls["n"] >= 2

    with patch.object(client, "is_healthy", side_effect=fake_healthy), patch.object(
        client, "_npx_path", return_value="/usr/bin/npx"
    ), patch("subprocess.Popen") as mock_popen, patch("time.sleep"):
        assert client.ensure_running(port=10531, timeout=5) is True
        mock_popen.assert_called_once()


def test_chat_uses_proxy_when_available_and_restores_env(monkeypatch):
    monkeypatch.setenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    import vendor_intel.placeholders.llm as llm_module

    prior_provider = llm_module.LLM_PROVIDER

    with patch.object(client, "ensure_running", return_value=True), patch.object(
        llm_module, "llm_complete", return_value="proxy answer"
    ) as mock_complete:
        result = client.chat("system", "user", port=10531)

    assert result == "proxy answer"
    mock_complete.assert_called_once()
    # Must pass an explicit model the proxy actually serves — the generic
    # placeholders.llm default ("gpt-4o-mini") gets a bare HTTP 500 from
    # openai-oauth since it isn't one of the logged-in account's models.
    assert mock_complete.call_args.kwargs.get("model") == client.DEFAULT_MODEL
    # Provider/env restored to prior state after the call.
    assert llm_module.LLM_PROVIDER == prior_provider
    import os

    assert os.environ.get("OPENAI_BASE_URL") == "https://api.openai.com/v1"


def test_chat_uses_extended_timeout_and_restores_it(monkeypatch):
    """A full-size draft response measured at ~267s against the real proxy
    — well past placeholders.llm's 120s default. The proxy call must
    temporarily raise _TIMEOUT to PROXY_TIMEOUT, and restore the prior
    value afterward so other pipelines' calls aren't affected."""
    import vendor_intel.placeholders.llm as llm_module

    prior_timeout = llm_module._TIMEOUT
    seen_timeout_during_call = {}

    def fake_complete(*a, **kw):
        seen_timeout_during_call["value"] = llm_module._TIMEOUT
        return "proxy answer"

    with patch.object(client, "ensure_running", return_value=True), patch.object(
        llm_module, "llm_complete", side_effect=fake_complete
    ):
        client.chat("system", "user", port=10531)

    assert seen_timeout_during_call["value"] == client.PROXY_TIMEOUT
    assert llm_module._TIMEOUT == prior_timeout


def test_chat_retries_chatgpt_itself_on_transient_failure_not_a_different_provider(monkeypatch):
    """A ChatGPT-side transient failure (not a usage limit — e.g. a one-off
    proxy hiccup) must be retried against ChatGPT again, never routed to a
    different LLM_PROVIDER."""
    monkeypatch.setattr(client, "_RETRY_PAUSE_SECONDS", 0)  # don't actually pause in tests
    with patch.object(client, "ensure_running", return_value=True), patch(
        "vendor_intel.placeholders.llm.llm_complete",
        side_effect=[RuntimeError("connection reset"), "real chatgpt answer"],
    ) as mock_complete:
        result = client.chat("system", "user")

    assert result == "real chatgpt answer"
    assert mock_complete.call_count == 2
    # Every call must still be through the ChatGPT-proxy model, never a
    # different provider's model.
    for call in mock_complete.call_args_list:
        assert call.kwargs.get("model") == client.DEFAULT_MODEL


def test_chat_raises_after_exhausting_transient_retries_never_substitutes_answer(monkeypatch):
    """If ChatGPT genuinely fails every retry with a non-usage-limit error,
    chat() must raise rather than silently return a different provider's
    answer as if it came from ChatGPT."""
    monkeypatch.setattr(client, "_RETRY_PAUSE_SECONDS", 0)
    with patch.object(client, "ensure_running", return_value=True), patch(
        "vendor_intel.placeholders.llm.llm_complete",
        side_effect=RuntimeError("connection reset"),
    ) as mock_complete:
        with pytest.raises(client.ChatGPTOAuthUnavailable):
            client.chat("system", "user")

    assert mock_complete.call_count == client._MAX_RETRIES


def test_chat_retries_when_proxy_returns_llm_failed_payload(monkeypatch):
    """placeholders.llm swallows HTTP errors into a {"status":"llm_failed"}
    JSON string instead of raising — chat() must still detect that and
    retry (against ChatGPT again) rather than treating the error payload
    as a real answer."""
    monkeypatch.setattr(client, "_RETRY_PAUSE_SECONDS", 0)
    error_payload = '{"error": "http_500: connection reset", "status": "llm_failed"}'
    with patch.object(client, "ensure_running", return_value=True), patch(
        "vendor_intel.placeholders.llm.llm_complete",
        side_effect=[error_payload, "real chatgpt answer"],
    ) as mock_complete:
        result = client.chat("system", "user")
        assert result == "real chatgpt answer"
        assert mock_complete.call_count == 2


def test_chat_json_also_retries_chatgpt_only(monkeypatch):
    monkeypatch.setattr(client, "_RETRY_PAUSE_SECONDS", 0)
    with patch.object(client, "ensure_running", return_value=True), patch(
        "vendor_intel.placeholders.llm.llm_complete_json",
        side_effect=[RuntimeError("connection reset"), {"ok": True}],
    ) as mock_complete:
        result = client.chat_json("system", "user")

    assert result == {"ok": True}
    assert mock_complete.call_count == 2


def test_chat_uses_long_pause_and_more_retries_specifically_for_usage_limit(monkeypatch):
    """A genuine account usage-limit error (the actual failure mode hit in
    a real run) does not clear in seconds — it needs an hours-scale wait,
    not the short _RETRY_PAUSE_SECONDS used for transient errors. This
    must go through _MAX_USAGE_LIMIT_RETRIES attempts at
    _USAGE_LIMIT_RETRY_PAUSE_SECONDS apart, not the short retry budget."""
    monkeypatch.setattr(client, "_USAGE_LIMIT_RETRY_PAUSE_SECONDS", 0)  # don't actually pause in tests
    with patch.object(client, "ensure_running", return_value=True), patch(
        "vendor_intel.placeholders.llm.llm_complete",
        side_effect=[RuntimeError("The usage limit has been reached")] * 3 + ["real chatgpt answer"],
    ) as mock_complete:
        result = client.chat("system", "user")

    assert result == "real chatgpt answer"
    assert mock_complete.call_count == 4


def test_chat_raises_after_exhausting_usage_limit_retries(monkeypatch):
    monkeypatch.setattr(client, "_USAGE_LIMIT_RETRY_PAUSE_SECONDS", 0)
    with patch.object(client, "ensure_running", return_value=True), patch(
        "vendor_intel.placeholders.llm.llm_complete",
        side_effect=RuntimeError("The usage limit has been reached"),
    ) as mock_complete:
        with pytest.raises(client.ChatGPTOAuthUnavailable):
            client.chat("system", "user")

    assert mock_complete.call_count == client._MAX_USAGE_LIMIT_RETRIES


def test_looks_like_usage_limit_detects_common_phrasings():
    assert client._looks_like_usage_limit("The usage limit has been reached") is True
    assert client._looks_like_usage_limit("rate limit exceeded") is True
    assert client._looks_like_usage_limit("quota exceeded") is True
    assert client._looks_like_usage_limit("connection reset") is False
    assert client._looks_like_usage_limit("") is False


def test_chat_raises_immediately_when_proxy_never_becomes_available(monkeypatch):
    """ensure_running() returning False every time (proxy can't start at
    all — npx missing, spawn failed) should still go through the retry loop
    and eventually raise, not silently substitute another provider."""
    monkeypatch.setattr(client, "_RETRY_PAUSE_SECONDS", 0)
    with patch.object(client, "ensure_running", return_value=False):
        with pytest.raises(client.ChatGPTOAuthUnavailable):
            client.chat("system", "user")
