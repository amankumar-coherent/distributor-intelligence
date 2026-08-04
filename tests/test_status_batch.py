"""status_batch.py — reads the SQLite registry directly and prints progress."""
from __future__ import annotations

import importlib.util
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
def status_batch_mod(monkeypatch):
    monkeypatch.syspath_prepend(str(ROOT / "src"))
    spec = importlib.util.spec_from_file_location("status_batch", ROOT / "status_batch.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_print_status_empty_batch(status_batch_mod, capsys):
    counts = status_batch_mod.print_status("nope")
    assert counts == {}
    out = capsys.readouterr().out
    assert "no markets loaded" in out


def test_print_status_shows_counts_and_no_eta_before_any_done(status_batch_mod, capsys):
    db.create_batch("b1")
    db.upsert_market("b1", query="Widget Market", country="Germany", market_slug="widget")
    db.upsert_market("b1", query="Gadget Market", country="Global", market_slug="gadget")
    db.claim_next_market("b1", "worker-1")  # one now 'running'

    counts = status_batch_mod.print_status("b1")
    assert counts == {"pending": 1, "running": 1}
    out = capsys.readouterr().out
    assert "Total markets    : 2" in out
    assert "Pending          : 1" in out
    assert "Running          : 1" in out
    assert "no completed markets yet" in out


def test_print_status_shows_eta_after_completion(status_batch_mod, capsys):
    db.create_batch("b1")
    db.set_batch_concurrency("b1", 2)
    m1, _ = db.upsert_market("b1", query="Widget Market", country="Germany", market_slug="widget")
    m2, _ = db.upsert_market("b1", query="Gadget Market", country="Global", market_slug="gadget")
    db.claim_next_market("b1", "worker-1")
    db.mark_market_done(m1, companies_exported=15, csv_path="widget.csv")

    counts = status_batch_mod.print_status("b1")
    assert counts == {"done": 1, "pending": 1}
    out = capsys.readouterr().out
    assert "Companies found  : 15" in out
    assert "ETA (remaining)" in out
    assert "concurrency=2" in out


def test_print_status_lists_recent_failures(status_batch_mod, capsys):
    db.create_batch("b1")
    m1, _ = db.upsert_market("b1", query="Bad Market", country="Global", market_slug="bad")
    db.claim_next_market("b1", "worker-1")
    db.mark_market_failed(m1, error="RuntimeError: boom", kind="permanent")

    status_batch_mod.print_status("b1")
    out = capsys.readouterr().out
    assert "Recent failures" in out
    assert "Bad Market" in out
    assert "boom" in out


def test_fmt_duration(status_batch_mod):
    assert status_batch_mod._fmt_duration(None) == "—"
    assert status_batch_mod._fmt_duration(30) == "30s"
    assert status_batch_mod._fmt_duration(150) == "2.5m"
    assert status_batch_mod._fmt_duration(7200) == "2.0h"


def test_main_watch_exits_when_nothing_pending(status_batch_mod, monkeypatch, capsys):
    import sys

    db.create_batch("b1")
    m1, _ = db.upsert_market("b1", query="Done Market", country="Global", market_slug="done")
    db.claim_next_market("b1", "worker-1")
    db.mark_market_done(m1, companies_exported=1)

    monkeypatch.setattr(sys, "argv", ["status_batch.py", "--batch-id", "b1", "--watch", "--interval", "0"])
    status_batch_mod.main()  # must return promptly, not loop forever
    out = capsys.readouterr().out
    assert "Batch finished" in out
