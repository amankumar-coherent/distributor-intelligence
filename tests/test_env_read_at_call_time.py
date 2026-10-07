"""Tunable settings must be read at CALL time, not import time.

Real bug this pins: MARKET_RESEARCH_BLOCK_COOLOFF=0 was set in .env and
silently ignored — the run kept waiting the full 300s. Cause: the constant
was initialised with a module-level os.getenv(), which runs BEFORE
apply_env_overrides() loads .env, so it always captured its default.

The symptom is nasty because nothing errors: the setting simply has no
effect, and it looks like the feature is broken rather than unconfigured.
"""
from __future__ import annotations

import pytest

from vendor_intel.pipeline.market_research import batch_runner as br
from vendor_intel.scraping import google_ai_mode as aim


# --- the exact reported case ---------------------------------------------


def test_block_cooloff_zero_is_honoured(monkeypatch):
    """"dont wait just retry" — 0 must mean no wait at all."""
    monkeypatch.setenv("MARKET_RESEARCH_BLOCK_COOLOFF", "0")
    assert br.block_cool_off_seconds() == 0.0


def test_block_cooloff_set_after_import_is_still_seen(monkeypatch):
    """Directly models the load order: the module is already imported, and
    .env is applied afterwards."""
    monkeypatch.setenv("MARKET_RESEARCH_BLOCK_COOLOFF", "45")
    assert br.block_cool_off_seconds() == 45.0


def test_block_cooloff_defaults_when_unset(monkeypatch):
    monkeypatch.delenv("MARKET_RESEARCH_BLOCK_COOLOFF", raising=False)
    assert br.block_cool_off_seconds() == br._DEFAULT_BLOCKED_COOL_OFF_SECONDS


def test_block_cooloff_ignores_garbage(monkeypatch):
    """A typo must fall back to the default, not crash the run."""
    monkeypatch.setenv("MARKET_RESEARCH_BLOCK_COOLOFF", "not-a-number")
    assert br.block_cool_off_seconds() == br._DEFAULT_BLOCKED_COOL_OFF_SECONDS


def test_block_cooloff_clamps_negative(monkeypatch):
    monkeypatch.setenv("MARKET_RESEARCH_BLOCK_COOLOFF", "-60")
    assert br.block_cool_off_seconds() == 0.0


# --- the same class of bug in the other tunables -------------------------


@pytest.mark.parametrize(
    "env_name,accessor",
    [
        ("MARKET_RESEARCH_MAX_BLOCKED_ROUNDS", lambda: br.max_blocked_rounds()),
        ("GOOGLE_AI_MODE_MAX_FAILURES", lambda: aim.max_consecutive_failures()),
        ("GOOGLE_AI_MODE_MAX_RESETS", lambda: aim.max_session_resets()),
    ],
)
def test_int_tunables_read_env_at_call_time(env_name, accessor, monkeypatch):
    monkeypatch.setenv(env_name, "7")
    assert accessor() == 7


@pytest.mark.parametrize(
    "env_name,accessor",
    [
        ("MARKET_RESEARCH_MAX_BLOCKED_ROUNDS", lambda: br.max_blocked_rounds()),
        ("GOOGLE_AI_MODE_MAX_FAILURES", lambda: aim.max_consecutive_failures()),
        ("GOOGLE_AI_MODE_MAX_RESETS", lambda: aim.max_session_resets()),
    ],
)
def test_int_tunables_survive_garbage(env_name, accessor, monkeypatch):
    monkeypatch.setenv(env_name, "")
    assert accessor() >= 1
    monkeypatch.setenv(env_name, "abc")
    assert accessor() >= 1


def test_query_delay_zero_disables_pacing(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_QUERY_DELAY", "0")
    assert aim.query_delay_seconds() == 0.0


def test_no_tunable_is_captured_at_import(monkeypatch):
    """Guard against the pattern coming back: a module-level
    `X = int(os.getenv(...))` cannot see .env and must not be reintroduced."""
    import re
    from pathlib import Path

    root = Path(__file__).resolve().parents[1] / "src" / "vendor_intel"
    pattern = re.compile(
        r"^_?[A-Z][A-Z0-9_]*\s*=\s*(?:int|float|bool)?\(?\s*os\.getenv\(", re.M
    )
    offenders = []
    for f in list((root / "pipeline" / "market_research").glob("*.py")) + [
        root / "scraping" / "google_ai_mode.py"
    ]:
        for m in pattern.finditer(f.read_text(encoding="utf-8")):
            offenders.append(f"{f.name}: {m.group(0).strip()}")
    assert not offenders, (
        "module-level os.getenv() runs before .env is loaded; use an "
        "accessor function instead:\n  " + "\n  ".join(offenders)
    )
