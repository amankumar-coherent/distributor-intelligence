"""Per-provider/model pricing table (src/vendor_intel/pipeline/llm_pricing.py)."""
from __future__ import annotations

from vendor_intel.pipeline import llm_pricing
from vendor_intel.pipeline.llm_pricing import estimate_cost_usd, price_per_million


def test_known_model_price():
    assert price_per_million("openai", "gpt-4o-mini") == (0.15, 0.60)


def test_free_providers_are_zero():
    assert price_per_million("opencode", "deepseek-v4-flash-free") == (0.0, 0.0)
    assert price_per_million("mock", "") == (0.0, 0.0)


def test_unknown_model_on_known_provider_falls_back_to_cheapest():
    prices = price_per_million("openai", "some-future-model")
    assert prices is not None
    known_openai_prices = [p for (pr, _m), p in llm_pricing._PRICING.items() if pr == "openai"]
    assert prices == min(known_openai_prices, key=lambda p: p[0] + p[1])


def test_unknown_provider_returns_none():
    assert price_per_million("some-new-provider", "some-model") is None


def test_estimate_cost_usd_matches_manual_math():
    # gpt-4o-mini: $0.15/$0.60 per 1M -> 1000 in + 500 out
    cost = estimate_cost_usd("openai", "gpt-4o-mini", 1000, 500)
    expected = (1000 * 0.15 + 500 * 0.60) / 1_000_000
    assert abs(cost - expected) < 1e-9


def test_estimate_cost_usd_zero_for_free_provider():
    assert estimate_cost_usd("opencode", "deepseek-v4-flash-free", 100_000, 50_000) == 0.0


def test_estimate_cost_usd_zero_for_unknown_provider():
    assert estimate_cost_usd("brand-new-provider", "x", 1000, 1000) == 0.0


def test_anthropic_sonnet_priced_differently_than_haiku():
    haiku = estimate_cost_usd("anthropic", "claude-haiku-4-5-20251001", 1000, 1000)
    sonnet = estimate_cost_usd("anthropic", "claude-sonnet-4-6", 1000, 1000)
    assert sonnet > haiku  # the exact bug this replaces: a flat rate treated these as equal
