"""A hand-solved CAPTCHA must resume the query, not waste the solve.

Reported live, with a screenshot of a solved "I'm not a robot" checkbox on
google.com/sorry/index: "after solving captcha only buffering, not move to
anything, after 60 second again come captcha". Two independent bugs caused
that, and this file pins both fixes:

  1. ask() slept a flat 60s then `continue`d WITHOUT re-issuing page.goto,
     so the next attempt re-read the same /sorry/ page. With retries=1 the
     budget was already spent, so a successful solve still returned "".
  2. run_until_target() counted the resulting empty batch as evidence the
     market was exhausted, and stopped. A real run ended at 318/500
     companies that way, reporting exhaustion when it was a rate limit.
"""
from __future__ import annotations

from unittest.mock import patch

from vendor_intel.pipeline.market_research import batch_runner
from vendor_intel.scraping import google_ai_mode as aim


class _FakePage:
    """Page that shows a CAPTCHA until `captcha_polls` polls have elapsed."""

    def __init__(self, captcha_polls: int, answer: str = "REAL ANSWER"):
        self.url = "https://www.google.com/sorry/index"
        self._polls_left = captcha_polls
        self._answer = answer
        self.goto_calls: list[str] = []

    def goto(self, url, **_kw):
        self.goto_calls.append(url)
        # Navigating while blocked lands back on /sorry/; once the CAPTCHA
        # has been solved, the query URL loads normally.
        self.url = (
            "https://www.google.com/sorry/index"
            if self._polls_left > 0
            else "https://www.google.com/search?udm=50"
        )

    def wait_for_timeout(self, _ms):
        if self._polls_left > 0:
            self._polls_left -= 1
        if self._polls_left == 0:
            self.url = "https://www.google.com/search?udm=50"


def _patched_ask(page, monkeypatch):
    """Run ask() against a fake page with all real waiting stubbed out."""
    monkeypatch.setattr(aim, "_get_page", lambda: page)
    monkeypatch.setattr(aim, "_pace_query", lambda: None)
    monkeypatch.setattr(aim, "_wait_for_response", lambda _p: None)
    monkeypatch.setattr(aim, "_extract_answer_text", lambda _p: "REAL ANSWER")
    monkeypatch.setattr(aim.time, "sleep", lambda _s: None)
    monkeypatch.setattr(aim, "_consecutive_captchas", 0)
    return aim.ask("test prompt")


def test_solved_captcha_resends_the_query_and_returns_the_answer(monkeypatch):
    """The core regression: a solve must produce an answer, not ""."""
    page = _FakePage(captcha_polls=2)

    answer = _patched_ask(page, monkeypatch)

    assert answer == "REAL ANSWER", "a solved CAPTCHA must yield the answer"
    assert len(page.goto_calls) >= 2, (
        "the query must be RE-ISSUED after the solve; without a second goto "
        "the next pass just re-reads the /sorry/ page"
    )


def test_solve_does_not_consume_the_retry_budget(monkeypatch):
    """With retries=1 the old loop spent both attempts on the CAPTCHA page.
    A solve returns the attempt, so the query still gets a real try."""
    page = _FakePage(captcha_polls=1)

    monkeypatch.setattr(aim, "_get_page", lambda: page)
    monkeypatch.setattr(aim, "_pace_query", lambda: None)
    monkeypatch.setattr(aim, "_wait_for_response", lambda _p: None)
    monkeypatch.setattr(aim, "_extract_answer_text", lambda _p: "OK")
    monkeypatch.setattr(aim.time, "sleep", lambda _s: None)
    monkeypatch.setattr(aim, "_consecutive_captchas", 0)

    assert aim.ask("prompt", retries=1) == "OK"


class _StuckPage:
    """A CAPTCHA that is never solved — models walking away from the run."""

    url = "https://www.google.com/sorry/index"

    def __init__(self):
        self.goto_calls: list[str] = []

    def goto(self, url, **_kw):
        self.goto_calls.append(url)

    def wait_for_timeout(self, _ms):
        pass  # stays on /sorry/ forever


def test_unsolved_captcha_still_gives_up_and_returns_empty(monkeypatch):
    """The breaker must still work: an unattended CAPTCHA cannot loop
    forever, or a run spins for hours with nobody watching."""
    page = _StuckPage()

    # timeout_seconds=0 via a patched waiter keeps this instant while still
    # exercising the real give-up path.
    monkeypatch.setattr(aim, "_wait_for_captcha_clear", lambda *_a, **_k: False)
    answer = _patched_ask(page, monkeypatch)

    assert answer == ""


def test_unsolved_captcha_terminates_and_does_not_hang(monkeypatch):
    """Bounded work, not just a bounded answer: a permanently blocked query
    must stop issuing requests rather than looping on the CAPTCHA."""
    page = _StuckPage()
    monkeypatch.setattr(aim, "_wait_for_captcha_clear", lambda *_a, **_k: False)

    _patched_ask(page, monkeypatch)

    assert len(page.goto_calls) <= aim._MAX_CAPTCHA_ROUNDS_PER_QUERY + 2, (
        f"too many requests while blocked: {len(page.goto_calls)}"
    )


def test_captcha_rounds_are_bounded(monkeypatch):
    """Even with repeated solves, one query can't retry indefinitely."""
    assert aim._MAX_CAPTCHA_ROUNDS_PER_QUERY >= 1
    assert aim._MAX_CAPTCHA_ROUNDS_PER_QUERY <= 10


def test_wait_for_captcha_clear_returns_early_on_solve(monkeypatch):
    """Polling, not a blind sleep: a 5s solve must not cost a full 60s."""
    page = _FakePage(captcha_polls=1)
    monkeypatch.setattr(aim, "_consecutive_captchas", 1)

    assert aim._wait_for_captcha_clear(page, timeout_seconds=60) is True


def test_wait_for_captcha_clear_times_out_when_unsolved(monkeypatch):
    page = _FakePage(captcha_polls=10_000)
    monkeypatch.setattr(aim, "_consecutive_captchas", 1)

    # timeout_seconds=0 makes the deadline immediate, no real waiting.
    assert aim._wait_for_captcha_clear(page, timeout_seconds=0) is False


def test_wait_for_captcha_clear_survives_closed_window(monkeypatch):
    """If the user closes the browser mid-CAPTCHA, this must return False
    rather than raising into the pipeline."""

    class _Dead:
        url = "https://www.google.com/sorry/index"

        def wait_for_timeout(self, _ms):
            raise RuntimeError("Target page, context or browser has been closed")

    monkeypatch.setattr(aim, "_consecutive_captchas", 1)
    assert aim._wait_for_captcha_clear(_Dead(), timeout_seconds=60) is False


# --- batch_runner: blocking must not be reported as exhaustion -----------


def test_blocked_batch_cools_off_instead_of_declaring_exhaustion(monkeypatch):
    """The bug that produced "318/500, exhausted": a CAPTCHA-blocked batch
    was counted toward the empty-batch limit. It must cool off and retry
    instead, since the market is not actually out of companies."""
    calls = {"n": 0}

    def _fake_batch(_state, _market, **_kw):
        calls["n"] += 1
        # Blocked twice, then Google lets us through.
        return [] if calls["n"] <= 2 else [{"Company": "Found Co"}]

    blocked = {"v": True}
    monkeypatch.setattr(batch_runner, "_backend_is_blocked", lambda: blocked["v"])
    monkeypatch.setattr(batch_runner, "_reset_backend_block", lambda: None)
    monkeypatch.setattr(batch_runner, "_run_one_batch", _fake_batch)
    monkeypatch.setattr(batch_runner.time, "sleep", lambda _s: None)

    class _State:
        companies: list = []

        @staticmethod
        def load(_m):
            return _State()

        def add_batch(self, rows):
            type(self).companies = list(self.companies) + list(rows)
            blocked["v"] = False

        def save(self):
            pass

    _State.companies = []
    monkeypatch.setattr(batch_runner, "MarketResearchState", _State)

    summary = batch_runner.run_until_target("M", target=1, batch_size=1)

    assert calls["n"] >= 3, "must retry past the block rather than stopping"
    assert "1/1" in summary


def test_blocked_rounds_are_bounded(monkeypatch):
    """A permanent block must still terminate, saving progress."""
    monkeypatch.setattr(batch_runner, "_backend_is_blocked", lambda: True)
    monkeypatch.setattr(batch_runner, "_reset_backend_block", lambda: None)
    monkeypatch.setattr(batch_runner, "_run_one_batch", lambda *_a, **_k: [])
    monkeypatch.setattr(batch_runner.time, "sleep", lambda _s: None)

    class _State:
        companies: list = []

        @staticmethod
        def load(_m):
            return _State()

        def add_batch(self, rows):
            pass

        def save(self):
            pass

    monkeypatch.setattr(batch_runner, "MarketResearchState", _State)

    summary = batch_runner.run_until_target("M", target=500, batch_size=1)
    assert "0/500" in summary


def test_genuine_exhaustion_still_stops_quickly(monkeypatch):
    """When nothing is blocking, empty batches must still end the run —
    the cool-off path must not make real exhaustion take 15 extra minutes."""
    monkeypatch.setattr(batch_runner, "_backend_is_blocked", lambda: False)
    monkeypatch.setattr(batch_runner, "_run_one_batch", lambda *_a, **_k: [])
    slept: list[float] = []
    monkeypatch.setattr(batch_runner.time, "sleep", lambda s: slept.append(s))

    class _State:
        companies: list = []

        @staticmethod
        def load(_m):
            return _State()

        def add_batch(self, rows):
            pass

        def save(self):
            pass

    monkeypatch.setattr(batch_runner, "MarketResearchState", _State)

    batch_runner.run_until_target("M", target=500, batch_size=1)

    assert all(s <= batch_runner._EMPTY_BATCH_RETRY_PAUSE_SECONDS for s in slept), (
        "genuine exhaustion must use the short pause, not the long cool-off"
    )


def test_solved_captcha_still_counts_as_blocking(monkeypatch):
    """The bug that stopped a run at 467/500.

    Observed log, same batch:
        [ai-mode] CAPTCHA cleared (53s early) — re-sending the query
        [market-research] ... no sign of blocking — treating as genuinely
                          exhausted, stopping

    Cause: blocking was detected via _consecutive_captchas, which ask()
    resets to 0 on a successful solve. So a solved CAPTCHA left no
    evidence and looked exactly like no CAPTCHA at all. A solve gets one
    query through; it does NOT lift Google's throttle.
    """
    import time as _t

    # Exactly the post-solve state: breaker cleared, but a CAPTCHA did happen.
    monkeypatch.setattr(aim, "_consecutive_captchas", 0)
    monkeypatch.setattr(aim, "_captcha_total", 1)
    monkeypatch.setattr(aim, "_last_captcha_at", _t.monotonic())

    assert batch_runner._backend_is_blocked() is True, (
        "a solved CAPTCHA must still be treated as Google blocking us"
    )


def test_no_captcha_at_all_is_not_blocking(monkeypatch):
    """Genuine exhaustion must still be reachable, or a finished market
    would cool off forever instead of stopping."""
    monkeypatch.setattr(aim, "_consecutive_captchas", 0)
    monkeypatch.setattr(aim, "_captcha_total", 0)
    monkeypatch.setattr(aim, "_last_captcha_at", 0.0)

    assert batch_runner._backend_is_blocked() is False


def test_old_captcha_ages_out_of_the_blocking_window(monkeypatch):
    """A CAPTCHA from an hour ago must not keep a finished run cooling off
    forever — otherwise the run can never conclude 'exhausted'."""
    import time as _t

    monkeypatch.setattr(aim, "_captcha_total", 1)
    monkeypatch.setattr(
        aim, "_last_captcha_at", _t.monotonic() - (batch_runner._BLOCK_LOOKBACK_SECONDS + 60)
    )

    assert batch_runner._backend_is_blocked() is False


def test_saw_captcha_recently_survives_a_solve(monkeypatch):
    """Unit-level version of the same guarantee, at the backend."""
    import time as _t

    monkeypatch.setattr(aim, "_last_captcha_at", _t.monotonic())
    monkeypatch.setattr(aim, "_consecutive_captchas", 0)  # solved => breaker cleared

    assert aim.saw_captcha_recently(600.0) is True


def test_captcha_stats_reports_totals(monkeypatch):
    import time as _t

    monkeypatch.setattr(aim, "_captcha_total", 4)
    monkeypatch.setattr(aim, "_consecutive_captchas", 1)
    monkeypatch.setattr(aim, "_last_captcha_at", _t.monotonic())

    s = aim.captcha_stats()
    assert s["total"] == 4
    assert s["consecutive"] == 1
    assert s["seconds_since_last"] is not None


def test_refusal_page_is_detected():
    """Observed live (screenshot): AI Mode served the page and returned
    "It looks like there's no response available for this search. Try
    asking something else." That is NOT a block — Google answered fine and
    the model declined this prompt — so it must not trigger a cool-off."""
    assert aim.is_ai_mode_refusal(
        "It looks like there's no response available for this search. "
        "Try asking something else.\nAI responses may include mistakes."
    )


def test_real_answer_is_not_mistaken_for_a_refusal():
    """A long real answer must never be discarded, even if it happens to
    contain similar wording."""
    answer = '[{"company": "Ingram Micro", "website": "ingrammicro.com"}]' + " x" * 500
    assert aim.is_ai_mode_refusal(answer) is False


def test_empty_text_is_not_a_refusal():
    assert aim.is_ai_mode_refusal("") is False
    assert aim.is_ai_mode_refusal(None) is False


def test_refusal_is_not_reported_as_blocking(monkeypatch):
    """The overcorrection this fixes: a refusal was counted as a CAPTCHA
    block, so the run cooled off 300s and the wait changed nothing."""
    import time as _t

    now = _t.monotonic()
    monkeypatch.setattr(aim, "_last_captcha_at", now)  # a CAPTCHA did happen earlier
    monkeypatch.setattr(aim, "_captcha_total", 1)
    monkeypatch.setattr(aim, "_last_refusal_at", now)  # but AI Mode also refused

    assert batch_runner._backend_is_blocked() is False, (
        "a refusal must not be treated as blocking — waiting cannot fix it"
    )
    assert batch_runner._backend_refused() is True


def test_captcha_without_refusal_is_still_blocking(monkeypatch):
    """The earlier fix must keep working: a real block still cools off."""
    import time as _t

    monkeypatch.setattr(aim, "_last_captcha_at", _t.monotonic())
    monkeypatch.setattr(aim, "_captcha_total", 1)
    monkeypatch.setattr(aim, "_last_refusal_at", 0.0)

    assert batch_runner._backend_is_blocked() is True


def test_ask_returns_empty_on_refusal_not_the_refusal_text(monkeypatch):
    """Returning the refusal text would be parsed downstream as a broken
    JSON answer, hiding the real cause."""
    class _RefusingPage:
        url = "https://www.google.com/search?udm=50"

        def __init__(self):
            self.goto_calls = []

        def goto(self, url, **_kw):
            self.goto_calls.append(url)

        def wait_for_timeout(self, _ms):
            pass

    page = _RefusingPage()
    monkeypatch.setattr(aim, "_get_page", lambda: page)
    monkeypatch.setattr(aim, "_pace_query", lambda: None)
    monkeypatch.setattr(aim, "_wait_for_response", lambda _p: None)
    monkeypatch.setattr(
        aim,
        "_extract_answer_text",
        lambda _p: "It looks like there's no response available for this search.",
    )
    monkeypatch.setattr(aim.time, "sleep", lambda _s: None)
    monkeypatch.setattr(aim, "_last_refusal_at", 0.0)
    monkeypatch.setattr(aim, "_refusal_total", 0)

    assert aim.ask("prompt") == ""
    assert aim._refusal_total >= 1, "the refusal must be recorded"
    assert aim.saw_refusal_recently(600.0) is True


def test_refused_batch_retries_quickly_not_after_a_long_cooloff(monkeypatch):
    """A refusal must use the short retry pause, never the 300s cool-off.

    The cool-off is pinned here rather than read from the environment: the
    real .env may legitimately set MARKET_RESEARCH_BLOCK_COOLOFF=0 ("don't
    wait, just retry"), which would make an env-derived assertion
    vacuously true (or impossible) and stop testing anything.
    """
    monkeypatch.setenv("MARKET_RESEARCH_BLOCK_COOLOFF", "300")
    monkeypatch.setattr(batch_runner, "_backend_is_blocked", lambda: False)
    monkeypatch.setattr(batch_runner, "_backend_refused", lambda: True)
    monkeypatch.setattr(batch_runner, "_reset_backend_refusal", lambda: None)
    monkeypatch.setattr(batch_runner, "_run_one_batch", lambda *_a, **_k: [])
    slept: list[float] = []
    monkeypatch.setattr(batch_runner.time, "sleep", lambda s: slept.append(s))

    class _State:
        companies: list = []

        @staticmethod
        def load(_m):
            return _State()

        def add_batch(self, rows):
            pass

        def save(self):
            pass

    monkeypatch.setattr(batch_runner, "MarketResearchState", _State)

    batch_runner.run_until_target("M", target=500, batch_size=1)

    assert slept, "a refusal should retry, not fall through silently"
    assert all(s < 300.0 for s in slept), (
        f"refusal must not use the block cool-off; slept {slept}"
    )


def test_reset_clears_the_recency_marker(monkeypatch):
    """Without clearing this, every later empty batch still looks blocked
    and the run cools off repeatedly instead of ever finishing."""
    import time as _t

    monkeypatch.setattr(aim, "_last_captcha_at", _t.monotonic())
    monkeypatch.setattr(aim, "_consecutive_captchas", 3)

    batch_runner._reset_backend_block()

    assert aim._consecutive_captchas == 0
    assert aim.saw_captcha_recently(600.0) is False


def test_ask_records_captcha_in_a_solve_proof_way(monkeypatch):
    """End-to-end: after ask() solves a CAPTCHA and returns an answer, the
    blocking evidence must still be visible to the batch runner."""
    monkeypatch.setattr(aim, "_captcha_total", 0)
    monkeypatch.setattr(aim, "_last_captcha_at", 0.0)
    page = _FakePage(captcha_polls=2)

    assert _patched_ask(page, monkeypatch) == "REAL ANSWER"
    assert aim._captcha_total >= 1, "the CAPTCHA must be recorded"
    assert aim.saw_captcha_recently(600.0) is True, (
        "evidence must survive the solve, or the runner mislabels the next "
        "empty batch as exhaustion"
    )


def test_reset_backend_block_clears_the_breaker(monkeypatch):
    """Without this the cool-off is useless: the tripped breaker fails
    every query fast, so the retry never reaches Google."""
    monkeypatch.setattr(aim, "_consecutive_captchas", 5)
    batch_runner._reset_backend_block()
    assert aim._consecutive_captchas == 0
