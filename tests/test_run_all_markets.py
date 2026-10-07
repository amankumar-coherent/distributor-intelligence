"""Sequential multi-market queue runner.

Requested as a single command that runs one market to completion, then the
next. The properties worth pinning are about what happens when something
goes wrong mid-queue, since these runs take hours and get interrupted:

  * one market failing must NOT abort the remaining markets (a plain
    shell `&&` chain would),
  * markets must run strictly one at a time (they share one browser
    profile and one IP; concurrency causes profile-lock collisions and
    rate-limiting), and
  * Ctrl+C must stop the queue rather than being swallowed as a
    single-market failure.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location(
    "run_all_markets", _ROOT / "run_all_markets.py"
)
ram = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(ram)


# --- the queue ------------------------------------------------------------


def test_all_six_requested_markets_are_queued():
    joined = " | ".join(ram.MARKETS).lower()
    for fragment in (
        "silicon carbide", "wearable glucometer", "dry eyes",
        "solar rooftop", "polyurethane", "cpaas",
    ):
        assert fragment in joined, fragment


def test_cpaas_is_last():
    """Its distributor definition is the inverted one (channel partners ARE
    the distributors), so it runs after the straightforward markets have
    already banked their results."""
    assert "cpaas" in ram.MARKETS[-1].lower()


def test_market_names_carry_the_distributors_only_suffix():
    """The market string is the pipeline's identity for a market - the
    state file, the Excel name and the prompt all derive from it - so a
    changed string silently starts a NEW market instead of resuming."""
    for m in ram.MARKETS:
        assert m.startswith("Global "), m
        assert m.endswith("Distributors Only"), m


def test_market_names_are_unique():
    assert len(set(ram.MARKETS)) == len(ram.MARKETS)


# --- failure isolation ----------------------------------------------------


def test_one_failing_market_does_not_stop_the_queue(monkeypatch):
    """The whole reason this is a script and not `cmd && cmd && cmd`."""
    calls = []

    class _Result:
        def __init__(self, code):
            self.returncode = code

    def _fake_run(cmd, **_kw):
        market = cmd[cmd.index("--market") + 1]
        calls.append(market)
        return _Result(1 if "Polyurethane" in market else 0)

    monkeypatch.setattr(ram.subprocess, "run", _fake_run)
    monkeypatch.setattr(ram, "_saved_count", lambda _m: 0)

    results = [
        ram.run_market(m, target=10, verify=False, batch_size=20)
        for m in ram.MARKETS
    ]

    assert len(calls) == len(ram.MARKETS), "every market must still be attempted"
    failed = [r for r in results if r["status"] != "ok"]
    assert len(failed) == 1 and "Polyurethane" in failed[0]["market"]


def test_a_crashing_subprocess_is_caught(monkeypatch):
    """An OSError launching the child must not take down the queue."""
    def _boom(*_a, **_kw):
        raise OSError("cannot spawn")

    monkeypatch.setattr(ram.subprocess, "run", _boom)
    monkeypatch.setattr(ram, "_saved_count", lambda _m: 0)

    out = ram.run_market("Test Market", target=10, verify=False, batch_size=20)
    assert out["status"] != "ok"


def test_keyboard_interrupt_propagates(monkeypatch):
    """Ctrl+C should stop the QUEUE, not read as one market failing."""
    def _interrupt(*_a, **_kw):
        raise KeyboardInterrupt

    monkeypatch.setattr(ram.subprocess, "run", _interrupt)
    monkeypatch.setattr(ram, "_saved_count", lambda _m: 0)

    with pytest.raises(KeyboardInterrupt):
        ram.run_market("Test Market", target=10, verify=False, batch_size=20)


# --- the command each market gets -----------------------------------------


def _captured_cmd(monkeypatch, **kwargs):
    seen = {}

    class _Result:
        returncode = 0

    def _fake_run(cmd, **_kw):
        seen["cmd"] = cmd
        return _Result()

    monkeypatch.setattr(ram.subprocess, "run", _fake_run)
    monkeypatch.setattr(ram, "_saved_count", lambda _m: 0)
    ram.run_market("Global Test Market - Distributors Only", **kwargs)
    return seen["cmd"]


def test_export_and_verify_are_requested(monkeypatch):
    cmd = _captured_cmd(monkeypatch, target=400, verify=True, batch_size=20)
    assert "--export-xlsx" in cmd
    assert "--verify" in cmd


def test_verify_can_be_switched_off(monkeypatch):
    cmd = _captured_cmd(monkeypatch, target=400, verify=False, batch_size=20)
    assert "--export-xlsx" in cmd
    assert "--verify" not in cmd


def test_market_string_is_passed_verbatim(monkeypatch):
    """Any mangling here would start a new market rather than resume the
    existing one, because the state filename derives from this string."""
    cmd = _captured_cmd(monkeypatch, target=400, verify=True, batch_size=20)
    assert cmd[cmd.index("--market") + 1] == "Global Test Market - Distributors Only"


def test_target_is_forwarded(monkeypatch):
    cmd = _captured_cmd(monkeypatch, target=250, verify=True, batch_size=20)
    assert cmd[cmd.index("--target") + 1] == "250"


def test_runs_the_real_pipeline_entrypoint(monkeypatch):
    cmd = _captured_cmd(monkeypatch, target=400, verify=True, batch_size=20)
    assert any("run_market_research.py" in str(part) for part in cmd)
    assert (_ROOT / "run_market_research.py").is_file()


# --- progress reporting ---------------------------------------------------


def test_result_reports_companies_gained(monkeypatch):
    """The summary is the only place a long unattended run shows what it
    actually achieved per market."""
    counts = iter([120, 400])   # before, after

    class _Result:
        returncode = 0

    monkeypatch.setattr(ram.subprocess, "run", lambda *_a, **_k: _Result())
    monkeypatch.setattr(ram, "_saved_count", lambda _m: next(counts))

    out = ram.run_market("M", target=400, verify=False, batch_size=20)
    assert out["before"] == 120
    assert out["after"] == 400


def test_saved_count_is_zero_for_an_unknown_market():
    """Must not raise on a market that has no state file yet."""
    assert ram._saved_count("A Market That Has Never Been Run") == 0
