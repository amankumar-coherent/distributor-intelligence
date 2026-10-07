"""google_ai_mode query pacing: anti-rate-limit spacing between queries.

No browser is launched here — only the pure timing/config helpers are
exercised, per this repo's rule that a missed mock must never silently
drive a real browser (tests/conftest.py).
"""
from __future__ import annotations

import time
from unittest.mock import patch

from vendor_intel.scraping import google_ai_mode as aim


def test_query_delay_defaults_when_unset(monkeypatch):
    monkeypatch.delenv("GOOGLE_AI_MODE_QUERY_DELAY", raising=False)
    assert aim.query_delay_seconds() == aim._DEFAULT_QUERY_DELAY_SECONDS


def test_query_delay_honours_override(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "3.5")
    assert aim.query_delay_seconds() == 3.5


def test_query_delay_zero_disables_pacing(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "0")
    assert aim.query_delay_seconds() == 0.0


def test_query_delay_falls_back_on_garbage(monkeypatch):
    """A typo in .env must not crash a long run."""
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "not-a-number")
    assert aim.query_delay_seconds() == aim._DEFAULT_QUERY_DELAY_SECONDS


def test_pace_query_sleeps_when_called_too_soon(monkeypatch):
    """Back-to-back queries are what triggered Google's IP throttle
    ("unusual traffic") ~200 companies into a real run, so a query issued
    immediately after the previous one must be made to wait."""
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "5")
    monkeypatch.setattr(aim, "_last_query_at", time.monotonic())
    slept: list[float] = []
    with patch("time.sleep", side_effect=slept.append):
        aim._pace_query()
    assert slept and slept[0] > 0


def test_pace_query_does_not_sleep_when_disabled(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "0")
    monkeypatch.setattr(aim, "_last_query_at", time.monotonic())
    slept: list[float] = []
    with patch("time.sleep", side_effect=slept.append):
        aim._pace_query()
    assert not slept


def test_pace_query_does_not_sleep_on_first_query(monkeypatch):
    """Nothing to space out from yet — the first query should go straight
    through rather than paying the delay for no reason."""
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "5")
    monkeypatch.setattr(aim, "_last_query_at", 0.0)
    slept: list[float] = []
    with patch("time.sleep", side_effect=slept.append):
        aim._pace_query()
    assert not slept


def test_pace_query_skips_wait_when_enough_time_already_passed(monkeypatch):
    """Generation time counts toward the gap: a slow answer shouldn't then
    also pay the full delay on top of it."""
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "2")
    monkeypatch.setattr(aim, "_last_query_at", time.monotonic() - 60)
    slept: list[float] = []
    with patch("time.sleep", side_effect=slept.append):
        aim._pace_query()
    assert not slept
