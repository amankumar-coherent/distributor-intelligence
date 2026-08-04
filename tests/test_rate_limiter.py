"""Shared search-engine rate limiter + circuit breaker
(src/vendor_intel/clients/rate_limiter.py). All timing uses small, fast
capacities/rates so the suite stays quick — no test should take more than a
fraction of a second beyond explicit small sleeps.
"""
from __future__ import annotations

import time

import pytest

from vendor_intel.batch import db
from vendor_intel.clients import rate_limiter as rl


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    rl.reset_for_tests()
    yield
    rl.reset_for_tests()
    db.reset_for_tests()


def test_acquire_consumes_burst_capacity_instantly():
    rl.configure("test_engine", capacity=3, refill_per_sec=1000)  # fast refill, don't care about timing here
    t0 = time.time()
    for _ in range(3):
        assert rl.acquire("test_engine", max_wait=1.0) is True
    assert time.time() - t0 < 0.5  # burst capacity should not need to wait


def test_acquire_waits_for_refill_then_succeeds():
    rl.configure("test_engine", capacity=1, refill_per_sec=20)  # 1 token every 50ms after burst
    assert rl.acquire("test_engine", max_wait=1.0) is True  # burst token
    t0 = time.time()
    assert rl.acquire("test_engine", max_wait=1.0) is True  # must wait for refill
    assert time.time() - t0 >= 0.03  # actually waited, wasn't instant


def test_acquire_times_out_returns_false():
    rl.configure("test_engine", capacity=1, refill_per_sec=0.01)  # effectively no refill within test window
    assert rl.acquire("test_engine", max_wait=1.0) is True  # burst token
    assert rl.acquire("test_engine", max_wait=0.1) is False  # times out waiting for the next one


def test_note_outcome_success_does_not_trip_breaker():
    rl.note_outcome("test_engine", blocked=False)
    assert rl.is_dead("test_engine") is False


def test_note_outcome_blocked_trips_breaker():
    rl.note_outcome("test_engine", blocked=True)
    assert rl.is_dead("test_engine") is True
    assert rl.dead_until("test_engine") is not None
    assert rl.dead_until("test_engine") > time.time()


def test_acquire_returns_false_immediately_when_dead():
    rl.note_outcome("test_engine", blocked=True)
    t0 = time.time()
    assert rl.acquire("test_engine", max_wait=5.0) is False
    assert time.time() - t0 < 0.5  # doesn't wait out max_wait, fails fast


def test_repeated_blocks_escalate_cooldown():
    rl.note_outcome("test_engine", blocked=True)
    first_until = rl.dead_until("test_engine")
    rl.note_outcome("test_engine", blocked=True)
    second_until = rl.dead_until("test_engine")
    assert second_until > first_until  # escalating, not flat


def test_success_after_block_clears_fail_streak_but_not_active_cooldown():
    rl.note_outcome("test_engine", blocked=True)
    assert rl.is_dead("test_engine") is True
    rl.note_outcome("test_engine", blocked=False)
    # A success clears the streak for future escalation, but doesn't
    # retroactively un-trip an already-active cooldown.
    assert rl.is_dead("test_engine") is True


def test_note_outcome_blocked_checkpoints_to_db():
    rl.note_outcome("ddgs", blocked=True)
    state = db.load_search_rate_state()
    assert state["ddgs"] is not None
    assert state["ddgs"] > time.time()


def test_hydrate_from_db_restores_active_cooldown():
    future = time.time() + 3600
    db.checkpoint_search_rate_state("bing_html", tokens=0.0, dead_until=future)
    assert rl.is_dead("bing_html") is False  # not hydrated yet

    rl.hydrate_from_db()
    assert rl.is_dead("bing_html") is True
    assert abs(rl.dead_until("bing_html") - future) < 1.0


def test_hydrate_from_db_ignores_expired_cooldowns():
    past = time.time() - 10
    db.checkpoint_search_rate_state("ddgs", tokens=0.0, dead_until=past)
    rl.hydrate_from_db()
    assert rl.is_dead("ddgs") is False


def test_checkpoint_happens_periodically_not_every_acquire():
    rl.configure("test_engine", capacity=100, refill_per_sec=1000)
    for _ in range(rl._CHECKPOINT_EVERY - 1):
        rl.acquire("test_engine", max_wait=1.0)
    # Not yet a multiple of _CHECKPOINT_EVERY -> no checkpoint written.
    assert db.load_search_rate_state() == {}

    rl.acquire("test_engine", max_wait=1.0)  # the Nth acquire triggers a checkpoint
    assert "test_engine" in db.load_search_rate_state()
