"""Tests for the Google AI Mode (udm=50) Playwright client.

No browser is launched here — the page object is faked. What's worth
locking down is the completion-detection logic, because getting it wrong is
silent and expensive: an early version required >400 chars of answer text,
so every short reply spun until the 90s timeout (measured ~94s per query
instead of ~5s) while still returning the right answer.
"""
from __future__ import annotations

import pytest

from vendor_intel.scraping import google_ai_mode as aim


class FakePage:
    """Minimal stand-in for a Playwright page.

    `states` is a list of (body_text, thinking_count, main_text) applied one
    per poll, so a test can describe how the page evolves while generating.
    """

    def __init__(self, states):
        self.states = list(states)
        self.polls = 0
        self.url = "https://www.google.com/search?q=x&udm=50"

    def _state(self):
        idx = min(self.polls, len(self.states) - 1)
        return self.states[idx]

    def inner_text(self, _selector, timeout=None):
        return self._state()[0]

    def get_by_text(self, _pattern):
        count = self._state()[1]
        return type("L", (), {"count": staticmethod(lambda: count)})()

    def locator(self, _selector):
        main_text = self._state()[2]
        first = type("F", (), {"inner_text": staticmethod(lambda timeout=None: main_text)})()
        return type("L", (), {"first": first})()

    def wait_for_timeout(self, _ms):
        self.polls += 1


@pytest.fixture(autouse=True)
def _fast_polls(monkeypatch):
    """Keep the poll interval nominal — FakePage advances on wait_for_timeout,
    it doesn't really sleep."""
    monkeypatch.setattr(aim, "THINKING_POLL_MS", 1)


class TestWaitForResponse:
    def test_returns_immediately_on_the_ready_marker(self):
        page = FakePage([("AI Mode response is ready", 0, "short answer")])
        aim._wait_for_response(page)
        assert page.polls == 0  # no waiting at all

    def test_short_answer_does_not_spin_until_timeout(self):
        """Regression: a 280-char answer never satisfied the old >400-char
        check, so the loop ran the full 90s on every short reply."""
        short = "READY"
        page = FakePage([("body", 0, short), ("body", 0, short), ("body", 0, short)])
        aim._wait_for_response(page)
        assert page.polls <= 3

    def test_waits_while_thinking_then_returns_once_cleared(self):
        page = FakePage(
            [
                ("body", 1, ""),          # thinking
                ("body", 1, "partial"),   # still thinking
                ("body", 0, "the answer"),  # cleared
            ]
        )
        aim._wait_for_response(page)
        assert page.polls >= 2

    def test_waits_while_text_is_still_growing(self):
        page = FakePage(
            [
                ("body", 0, "a"),
                ("body", 0, "ab"),
                ("body", 0, "abc"),
                ("body", 0, "abc"),
                ("body", 0, "abc"),
            ]
        )
        aim._wait_for_response(page)
        assert page.polls >= 3  # didn't bail while the text was changing

    def test_gives_up_at_the_deadline_rather_than_hanging(self, monkeypatch):
        monkeypatch.setattr(aim, "RESPONSE_TIMEOUT_MS", 30)
        page = FakePage([("body", 1, "")])  # thinking forever
        aim._wait_for_response(page)  # must return, not hang


class TestCompletionSignals:
    @pytest.mark.parametrize(
        "body,expected",
        [
            ("AI Mode response is ready", True),
            ("some text ... Response is ready", True),
            ("still generating", False),
            ("", False),
        ],
    )
    def test_answer_is_complete(self, body, expected):
        assert aim._answer_is_complete(FakePage([(body, 0, "")])) is expected


class TestCaptchaDetection:
    def test_sorry_url_is_a_captcha(self):
        page = FakePage([("nothing unusual", 0, "")])
        page.url = "https://www.google.com/sorry/index?continue=..."
        assert aim._page_is_captcha(page) is True

    @pytest.mark.parametrize("marker", aim.CAPTCHA_MARKERS)
    def test_body_markers_are_captchas(self, marker):
        assert aim._page_is_captcha(FakePage([(f"...{marker}...", 0, "")])) is True

    def test_normal_page_is_not_a_captcha(self):
        assert aim._page_is_captcha(FakePage([("a normal answer", 0, "")])) is False


class TestUrlBuilding:
    def test_uses_udm_50_and_encodes_the_prompt(self):
        url = aim.build_ai_mode_url('is "Acme & Co" a distributor?')
        assert "udm=50" in url
        assert url.startswith("https://www.google.com/search?q=")
        assert " " not in url  # fully percent-encoded

    def test_empty_prompt_short_circuits(self):
        assert aim.ask("   ") == ""


class TestConfig:
    def test_disabled_by_default(self, monkeypatch):
        monkeypatch.delenv("GOOGLE_AI_MODE_ENABLED", raising=False)
        assert aim.ai_mode_enabled() is False

    @pytest.mark.parametrize("value", ["true", "1", "yes", "on", "TRUE"])
    def test_enabled_values(self, monkeypatch, value):
        monkeypatch.setenv("GOOGLE_AI_MODE_ENABLED", value)
        assert aim.ai_mode_enabled() is True

    def test_headless_off_by_default(self, monkeypatch):
        """A CAPTCHA can only be solved by hand, so the window must be
        visible until the profile has cleared one."""
        monkeypatch.delenv("GOOGLE_AI_MODE_HEADLESS", raising=False)
        assert aim.ai_mode_headless() is False
