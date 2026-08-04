"""run_batch.py CLI wiring — registry load -> confirm -> orchestrator.run_batch,
with the orchestrator itself mocked out (it has its own thorough test suite
in test_batch_orchestrator.py; this file is only about the CLI plumbing:
argument parsing, --input vs resume, the "nothing pending" short-circuit, and
that CLI flags reach orchestrator.run_batch with the right values).
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

from vendor_intel.batch import db

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


@pytest.fixture
def run_batch_mod(monkeypatch):
    """Import run_batch.py as a fresh module each test (it has no import-time
    side effects — all work happens inside main())."""
    monkeypatch.syspath_prepend(str(ROOT / "src"))
    spec = importlib.util.spec_from_file_location("run_batch", ROOT / "run_batch.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _write_markets_csv(tmp_path) -> Path:
    p = tmp_path / "markets.csv"
    p.write_text("query,country\nMarket A,Global\nMarket B,Global\n", encoding="utf-8")
    return p


def test_main_loads_registry_and_calls_orchestrator(tmp_path, monkeypatch, run_batch_mod):
    p = _write_markets_csv(tmp_path)
    calls = []

    def fake_run_batch(batch_id, settings, **kw):
        calls.append((batch_id, kw))
        return {"done": 2}

    monkeypatch.setattr("vendor_intel.batch.orchestrator.run_batch", fake_run_batch)
    monkeypatch.setattr(
        sys, "argv",
        ["run_batch.py", "--input", str(p), "--batch-id", "b1", "--yes", "--max-concurrent", "4"],
    )

    run_batch_mod.main()

    assert len(calls) == 1
    batch_id, kw = calls[0]
    assert batch_id == "b1"
    assert kw["max_concurrent_markets"] == 4
    assert kw["fmcg_xlsx"] is True  # --no-fmcg-xlsx not passed
    assert kw["limit"] is None

    # The registry itself was actually populated (not just the orchestrator call).
    assert len(db.list_markets("b1")) == 2


def test_main_no_fmcg_xlsx_flag_reaches_orchestrator(tmp_path, monkeypatch, run_batch_mod):
    p = _write_markets_csv(tmp_path)
    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda batch_id, settings, **kw: calls.append(kw) or {"done": 2},
    )
    monkeypatch.setattr(
        sys, "argv",
        ["run_batch.py", "--input", str(p), "--batch-id", "b1", "--yes", "--no-fmcg-xlsx"],
    )
    run_batch_mod.main()
    assert calls[0]["fmcg_xlsx"] is False


def test_main_resumes_without_input(tmp_path, monkeypatch, run_batch_mod):
    p = _write_markets_csv(tmp_path)
    from vendor_intel.batch.registry_loader import populate_registry

    populate_registry(p, "b1")  # pre-load, simulating a prior run

    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda batch_id, settings, **kw: calls.append(batch_id) or {"done": 2},
    )
    monkeypatch.setattr(sys, "argv", ["run_batch.py", "--batch-id", "b1", "--yes"])

    run_batch_mod.main()
    assert calls == ["b1"]


def test_main_errors_when_resuming_unknown_batch(monkeypatch, run_batch_mod, capsys):
    monkeypatch.setattr(sys, "argv", ["run_batch.py", "--batch-id", "does-not-exist", "--yes"])
    with pytest.raises(SystemExit):
        run_batch_mod.main()
    assert "no markets loaded" in capsys.readouterr().out


def test_main_nothing_pending_short_circuits(tmp_path, monkeypatch, run_batch_mod, capsys):
    p = _write_markets_csv(tmp_path)
    from vendor_intel.batch.registry_loader import populate_registry

    populate_registry(p, "b1")
    for m in db.list_markets("b1"):
        db.claim_next_market("b1", "w1")
        db.mark_market_done(m["id"], companies_exported=0)

    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda *a, **kw: calls.append(1) or {},
    )
    monkeypatch.setattr(sys, "argv", ["run_batch.py", "--batch-id", "b1", "--yes"])

    run_batch_mod.main()
    assert calls == []  # orchestrator never invoked — nothing pending
    assert "Nothing pending" in capsys.readouterr().out


def test_main_cancelled_without_yes(tmp_path, monkeypatch, run_batch_mod, capsys):
    p = _write_markets_csv(tmp_path)
    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda *a, **kw: calls.append(1) or {},
    )
    monkeypatch.setattr(sys, "argv", ["run_batch.py", "--input", str(p), "--batch-id", "b1"])
    monkeypatch.setattr("builtins.input", lambda *_: "n")  # anything non-empty cancels

    run_batch_mod.main()
    assert calls == []
    assert "Cancelled" in capsys.readouterr().out


def test_main_full_enrich_flag_reaches_orchestrator(tmp_path, monkeypatch, run_batch_mod):
    p = _write_markets_csv(tmp_path)
    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda batch_id, settings, **kw: calls.append(kw) or {"done": 2},
    )
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "run_batch.py",
            "--input",
            str(p),
            "--batch-id",
            "b1",
            "--yes",
            "--full-enrich",
            "--enrich-concurrent",
            "2",
            "--search-backfill",
        ],
    )
    run_batch_mod.main()
    assert calls[0]["full_enrich"] is True
    assert calls[0]["enrich_concurrent"] == 2
    assert calls[0]["search_backfill"] is True


def test_main_with_audit_flag_reaches_orchestrator(tmp_path, monkeypatch, run_batch_mod):
    p = _write_markets_csv(tmp_path)
    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda batch_id, settings, **kw: calls.append(kw) or {"done": 2},
    )
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "run_batch.py",
            "--input",
            str(p),
            "--batch-id",
            "b1",
            "--yes",
            "--with-audit",
            "--audit-concurrent",
            "4",
        ],
    )
    run_batch_mod.main()
    assert calls[0]["with_audit"] is True
    assert calls[0]["audit_concurrent"] == 4
