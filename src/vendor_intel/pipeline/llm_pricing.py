"""Static per-provider/per-model USD pricing (per 1M tokens).

Replaces llm_meter.py's old flat rate (one $/token number applied to every
call regardless of provider or model — wildly wrong once you use more than
one provider, e.g. pricing an Anthropic Sonnet call as if it were OpenAI's
gpt-4o-mini). This is a lookup table only, no network calls — update the
numbers here if a provider changes pricing or a new model is added.

Prices as of the models this project actually references in config.py /
placeholders/llm.py (DEFAULT_COMPILER_MODEL, DEFAULT_GEMINI_MODEL,
DEFAULT_GROQ_MODEL, and the live-configured OPENAI_MODEL).
"""
from __future__ import annotations

# (provider, model) -> (input $ / 1M tokens, output $ / 1M tokens)
_PRICING: dict[tuple[str, str], tuple[float, float]] = {
    # OpenAI
    ("openai", "gpt-4o-mini"): (0.15, 0.60),
    ("openai", "gpt-4o"): (2.50, 10.00),
    ("openai", "gpt-4.1"): (2.00, 8.00),
    ("openai", "gpt-4.1-mini"): (0.40, 1.60),
    ("openai", "gpt-4.1-nano"): (0.10, 0.40),
    # Anthropic
    ("anthropic", "claude-haiku-4-5-20251001"): (1.00, 5.00),
    ("anthropic", "claude-sonnet-4-6"): (3.00, 15.00),
    ("anthropic", "claude-sonnet-4-20250514"): (3.00, 15.00),
    ("anthropic", "claude-opus-4-8"): (15.00, 75.00),
    # Gemini
    ("gemini", "gemini-2.0-flash"): (0.10, 0.40),
    ("gemini", "gemini-2.5-flash"): (0.30, 2.50),
    # Groq (Llama hosting — priced per Groq's published rates)
    ("groq", "llama-3.3-70b-versatile"): (0.59, 0.79),
}

# Providers that are genuinely free for this project's purposes — no cost to
# estimate, and no benefit in guessing a nonzero number for them.
_FREE_PROVIDERS = {"opencode", "mock", "none"}


def price_per_million(provider: str, model: str) -> tuple[float, float] | None:
    """Return (input_$/1M, output_$/1M) for (provider, model).

    Returns (0.0, 0.0) for known-free providers, or None if the provider is
    paid but the specific model isn't in the table AND no same-provider
    fallback exists either (caller should treat None as "can't estimate").
    """
    provider = (provider or "").strip().lower()
    if provider in _FREE_PROVIDERS:
        return (0.0, 0.0)

    model = (model or "").strip()
    key = (provider, model)
    if key in _PRICING:
        return _PRICING[key]

    # Unknown model on a known paid provider: fall back to the cheapest
    # model we know for that provider rather than silently reporting $0,
    # which would be a worse error (under-counting real spend) than a
    # same-provider approximation.
    same_provider = [prices for (pr, _m), prices in _PRICING.items() if pr == provider]
    if same_provider:
        return min(same_provider, key=lambda p: p[0] + p[1])

    return None


def estimate_cost_usd(provider: str, model: str, input_tokens: int, output_tokens: int) -> float:
    """USD cost for one call. Returns 0.0 for free/unknown providers."""
    prices = price_per_million(provider, model)
    if prices is None:
        return 0.0
    input_price, output_price = prices
    return (max(0, input_tokens) * input_price + max(0, output_tokens) * output_price) / 1_000_000
