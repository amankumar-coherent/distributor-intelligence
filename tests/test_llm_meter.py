"""LLM call meter (src/vendor_intel/pipeline/llm_meter.py) — the meter is
correctly isolated per concurrent asyncio Task instead of being one shared
global that concurrently-running markets would stomp on each other.
"""
from __future__ import annotations

import asyncio

from vendor_intel.pipeline import llm_meter


def test_reset_then_get_returns_fresh_meter():
    m1 = llm_meter.reset_meter()
    assert m1.calls == 0
    m1.add_classify()
    m2 = llm_meter.get_meter()
    assert m2 is m1  # same meter until reset again
    m3 = llm_meter.reset_meter()
    assert m3.calls == 0
    assert m3 is not m1


def test_summary_prices_estimated_tokens_at_real_rate():
    m = llm_meter.reset_meter()
    m.add_phase1()  # estimated-token path: 4000 in / 2000 out

    summary = m.summary(provider="openai", model="gpt-4o-mini")
    expected = (4000 * 0.15 + 2000 * 0.60) / 1_000_000
    assert abs(summary["estimated_cost_usd"] - round(expected, 4)) < 1e-6
    assert summary["llm_calls_total"] == 1


def test_summary_zero_cost_for_free_provider():
    m = llm_meter.reset_meter()
    m.add_classify()
    m.add_phase1()
    summary = m.summary(provider="opencode", model="deepseek-v4-flash-free")
    assert summary["estimated_cost_usd"] == 0.0


def test_concurrent_markets_do_not_stomp_each_others_meter():
    """The core correctness fix: two 'markets' running concurrently as
    separate asyncio Tasks must each see only their own call counts, even
    though both call reset_meter()/get_meter() with no explicit meter
    object passed around (mirrors how classifier.py/orchestrator.py call it)."""

    async def run_market(n_calls: int, delay: float) -> int:
        llm_meter.reset_meter()
        for _ in range(n_calls):
            get = llm_meter.get_meter()
            get.add_classify()
            await asyncio.sleep(delay)  # yield control so the other task interleaves
        return llm_meter.get_meter().calls

    async def main():
        # Market A does 5 calls slowly; Market B does 3 calls, also slowly,
        # heavily interleaved with A via the sleeps above.
        results = await asyncio.gather(run_market(5, 0.01), run_market(3, 0.01))
        return results

    results = asyncio.run(main())
    assert results == [5, 3], (
        f"expected each concurrent market to see only its own calls, got {results}"
    )
