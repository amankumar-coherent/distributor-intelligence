"""Wait-and-retry when Google serves a CAPTCHA instead of results.

A block is temporary — it lifts once the request burst stops, or when the
user solves the challenge in the visible tab. The old behaviour threw the
company away (and, before the detection guard, parsed the block page as a
real answer). These tests pin: detect, wait, retry, and make the ~10
concurrent tabs share ONE wait instead of each hammering Google.
"""
from __future__ import annotations

import asyncio

import pytest

from vendor_intel.clients import google_ai_scraper as gs

BLOCK_PAGE = (
    "About this page Our systems have detected unusual traffic from your computer "
    "network. This page checks to see if it's really you sending the requests, and "
    "not a robot. Why did this happen? ... /search?q=Is%20Acme%20a%20distributor"
)
REAL_ANSWER = "Acme S.p.A. is a distributor of wearable medical devices in Italy."


@pytest.fixture(autouse=True)
def fast_and_isolated(monkeypatch):
    """Real waits would make the suite sleep for minutes."""
    monkeypatch.setenv("GOOGLE_AI_CAPTCHA_WAIT_SEC", "0.05")
    monkeypatch.setenv("GOOGLE_AI_CAPTCHA_RETRIES", "3")
    monkeypatch.setenv("GOOGLE_AI_SCRAPER_FORCE_OVERVIEW", "0")
    # Probe fast, and cap the hold, so a test can never sit on the real
    # 15s/30min production defaults.
    monkeypatch.setenv("GOOGLE_AI_CAPTCHA_PROBE_SEC", "0.01")
    monkeypatch.setattr(gs, "_captcha_probe_timeout", lambda: 0.3)
    # A fresh gate per test, or one test's block leaks into the next.
    monkeypatch.setattr(gs, "_captcha_gate", gs._CaptchaGate())


class TestDetection:
    def test_block_page_detected(self):
        assert gs.is_captcha_page({"markdown_raw": BLOCK_PAGE}) is True

    def test_detected_on_raw_text_even_when_cleaned_looks_fine(self):
        """The cleaners reformat the block page into readable prose, so
        detection must look at the RAW scrape."""
        assert gs.is_captcha_page(
            {"markdown_raw": BLOCK_PAGE, "markdown": "Acme is a distributor."}
        ) is True

    def test_real_answer_not_flagged(self):
        assert gs.is_captcha_page({"markdown_raw": REAL_ANSWER}) is False

    def test_empty_scrape_not_flagged(self):
        assert gs.is_captcha_page({}) is False


class TestRetry:
    def test_retries_after_waiting_and_succeeds(self, monkeypatch):
        """The company must be recovered, not discarded."""
        calls = {"n": 0}

        async def fake_ask_once(base, q, *, thread_id=None, mode=None):
            calls["n"] += 1
            if calls["n"] == 1:
                return {"markdown_raw": BLOCK_PAGE, "markdown": BLOCK_PAGE}
            return {"markdown_raw": REAL_ANSWER, "markdown": REAL_ANSWER}

        monkeypatch.setattr(gs, "_ask_once", fake_ask_once)
        monkeypatch.setattr(gs, "_apply_deepseek_sentence_clean",
                            lambda q, d, llm_clean=None: _identity(d))

        out = asyncio.run(gs.google_ai_ask("Is Acme a distributor?"))
        # 1 blocked attempt, 1 clearance probe, then the real retry — the
        # probe is what confirms the block lifted before more tabs open.
        assert calls["n"] == 3
        assert "distributor" in out["markdown"]
        assert not out.get("captcha_blocked")

    def test_gives_up_after_the_retry_budget(self, monkeypatch):
        """Never return the block page as an answer — it echoes the query,
        so it parses as a confident 'distributor' verdict."""
        async def always_blocked(base, q, *, thread_id=None, mode=None):
            return {"markdown_raw": BLOCK_PAGE, "markdown": BLOCK_PAGE}

        monkeypatch.setattr(gs, "_ask_once", always_blocked)
        out = asyncio.run(gs.google_ai_ask("Is Acme a distributor?"))
        assert out["captcha_blocked"] is True
        assert out["error"] == "captcha_blocked"
        assert out["markdown"] == ""

    def test_clean_query_never_waits(self, monkeypatch):
        async def fine(base, q, *, thread_id=None, mode=None):
            return {"markdown_raw": REAL_ANSWER, "markdown": REAL_ANSWER}

        monkeypatch.setattr(gs, "_ask_once", fine)
        monkeypatch.setattr(gs, "_apply_deepseek_sentence_clean",
                            lambda q, d, llm_clean=None: _identity(d))

        async def timed():
            loop = asyncio.get_event_loop()
            t0 = loop.time()
            await gs.google_ai_ask("Is Acme a distributor?")
            return loop.time() - t0

        assert asyncio.run(timed()) < 0.5


class TestSharedGate:
    def test_concurrent_callers_share_one_wait(self):
        """Ten tabs hitting the same block must not serialise into ten
        consecutive waits."""
        gate = gs._CaptchaGate()

        async def scenario():
            loop = asyncio.get_event_loop()
            t0 = loop.time()
            await asyncio.gather(*(gate.block_for(0.2) for _ in range(10)))
            return loop.time() - t0

        elapsed = asyncio.run(scenario())
        assert elapsed < 0.6, f"waits serialised: {elapsed:.2f}s"

    def test_later_query_waits_out_an_active_block(self):
        gate = gs._CaptchaGate()

        async def scenario():
            loop = asyncio.get_event_loop()
            await gate.block_for(0.0)
            asyncio.ensure_future(gate.block_for(0.2))
            await asyncio.sleep(0.01)
            t0 = loop.time()
            await gate.wait_if_blocked()
            return loop.time() - t0

        assert asyncio.run(scenario()) > 0.1

    def test_no_block_means_no_wait(self):
        gate = gs._CaptchaGate()

        async def scenario():
            loop = asyncio.get_event_loop()
            t0 = loop.time()
            await gate.wait_if_blocked()
            return loop.time() - t0

        assert asyncio.run(scenario()) < 0.05


async def _identity(d):
    return d


class TestWaitUntilSolved:
    """When a CAPTCHA appears, stop opening tabs and wait for it to be
    solved — rather than sleeping a fixed interval and resuming blind.

    Firing ten fresh tabs into an active block is what prolongs the block,
    so the queue must stay closed until Google actually answers again.
    """

    @pytest.fixture(autouse=True)
    def _fast(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_CAPTCHA_WAIT_FOR_SOLVE", "true")
        monkeypatch.setenv("GOOGLE_AI_CAPTCHA_PROBE_SEC", "0.05")
        monkeypatch.setenv("GOOGLE_AI_CAPTCHA_MAX_WAIT_SEC", "60")
        monkeypatch.setattr(gs, "_captcha_gate", gs._CaptchaGate())

    def test_resumes_only_once_the_block_clears(self, monkeypatch):
        """The probe returns blocked twice, then clean — as happens when
        the user solves the challenge in the visible tab."""
        seq = [BLOCK_PAGE, BLOCK_PAGE, REAL_ANSWER]
        calls = {"n": 0}

        async def probe(base, q, *, thread_id=None, mode=None):
            i = min(calls["n"], len(seq) - 1)
            calls["n"] += 1
            return {"markdown_raw": seq[i], "markdown": seq[i]}

        gate = gs._CaptchaGate()
        asyncio.run(gate.hold_until_clear("http://x", "pro", probe))
        assert calls["n"] == 3, "must keep probing until the answer is clean"

    def test_gives_up_after_the_wait_budget(self, monkeypatch):
        """An unattended overnight run must not stall forever on a
        challenge nobody is there to solve."""
        # 60 is the floor the getter enforces; with a 0.05s probe interval
        # that is ~1200 probes, so cap the probe count instead by making the
        # budget expire quickly in loop time.
        monkeypatch.setenv("GOOGLE_AI_CAPTCHA_MAX_WAIT_SEC", "60")
        monkeypatch.setattr(gs, "_captcha_probe_timeout", lambda: 0.3)

        probes = {"n": 0}

        async def always_blocked(base, q, *, thread_id=None, mode=None):
            probes["n"] += 1
            return {"markdown_raw": BLOCK_PAGE}

        gate = gs._CaptchaGate()
        asyncio.run(gate.hold_until_clear("http://x", "pro", always_blocked))
        # It gave up rather than looping forever.
        assert probes["n"] >= 1

    def test_only_one_caller_probes_the_rest_wait(self):
        """Ten blocked tabs must not each start their own probe — that is
        exactly the tab storm this change exists to prevent."""
        probes = {"n": 0}

        async def probe(base, q, *, thread_id=None, mode=None):
            probes["n"] += 1
            return {"markdown_raw": REAL_ANSWER}

        async def scenario():
            gate = gs._CaptchaGate()
            await asyncio.gather(
                *(gate.hold_until_clear("http://x", "pro", probe) for _ in range(10))
            )

        asyncio.run(scenario())
        assert probes["n"] <= 2, f"expected a single prober, got {probes['n']}"

    def test_fixed_sleep_mode_still_available(self, monkeypatch):
        """Setting the flag false restores the old fixed-wait behaviour."""
        monkeypatch.setenv("GOOGLE_AI_CAPTCHA_WAIT_FOR_SOLVE", "false")
        assert gs._captcha_wait_for_solve() is False

    def test_enabled_by_default(self, monkeypatch):
        monkeypatch.delenv("GOOGLE_AI_CAPTCHA_WAIT_FOR_SOLVE", raising=False)
        assert gs._captcha_wait_for_solve() is True


class TestWaitersResumePromptly:
    """Non-probing tabs must resume when the block CLEARS, not when the
    timeout budget expires.

    The first version had them sleep `_until - now`, which was the full
    30-minute probe budget. So after the prober logged "CAPTCHA cleared"
    the other nine tabs stayed parked and the run went silent for half an
    hour — observed live as a screening run that stopped sending queries
    with the server sitting at queued=0.
    """

    @pytest.fixture(autouse=True)
    def _fast(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_CAPTCHA_PROBE_SEC", "0.01")
        # A deliberately LONG budget: the point is that waiters must not
        # sleep for it once the block has lifted.
        monkeypatch.setattr(gs, "_captcha_probe_timeout", lambda: 30.0)

    def test_waiters_return_soon_after_the_block_lifts(self):
        async def probe(base, q, *, thread_id=None, mode=None):
            return {"markdown_raw": REAL_ANSWER}

        async def scenario():
            gate = gs._CaptchaGate()
            loop = asyncio.get_event_loop()
            t0 = loop.time()
            await asyncio.gather(
                *(gate.hold_until_clear("http://x", "pro", probe) for _ in range(5))
            )
            return loop.time() - t0

        elapsed = asyncio.run(scenario())
        # Must be seconds, not the 30s budget.
        assert elapsed < 5.0, f"waiters slept the whole budget: {elapsed:.1f}s"

    def test_gate_is_reusable_after_a_hold(self):
        """A second CAPTCHA later in the run must hold again, not be
        permanently 'cleared'."""
        async def probe(base, q, *, thread_id=None, mode=None):
            return {"markdown_raw": REAL_ANSWER}

        async def scenario():
            gate = gs._CaptchaGate()
            await gate.hold_until_clear("http://x", "pro", probe)
            await gate.hold_until_clear("http://x", "pro", probe)
            # After both holds the gate must be open.
            loop = asyncio.get_event_loop()
            t0 = loop.time()
            await gate.wait_if_blocked()
            return loop.time() - t0

        assert asyncio.run(scenario()) < 0.5
