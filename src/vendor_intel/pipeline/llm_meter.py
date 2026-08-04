"""LLM call counter for one pipeline run.

The meter is stored in a contextvars.ContextVar, not a plain module global —
important since the concurrent batch orchestrator runs several markets "at
once" as separate asyncio Tasks on one event loop. A plain global would let
one market's reset_meter() wipe out another concurrently-running market's
in-progress counts (each asyncio Task gets its own isolated copy of the
context, so ContextVar.set() in one Task never affects a sibling Task).

Call sites use the estimated-token helpers (add_classify()/add_profile()/
add_phase1()/add_market_understanding()) to bump call counts; summary()
prices those estimated tokens via llm_pricing.py's per-provider/model rate
table for an approximate end-of-run cost line.
"""
from __future__ import annotations

import contextvars
from dataclasses import dataclass

from vendor_intel.pipeline.llm_pricing import estimate_cost_usd


@dataclass
class LlmMeter:
    calls: int = 0
    phase1_compile: int = 0
    market_understanding: int = 0
    classify: int = 0
    export_profile: int = 0
    smart_crawl_chunks: int = 0
    estimated_input_tokens: int = 0
    estimated_output_tokens: int = 0

    def add_classify(self, tokens_in: int = 800, tokens_out: int = 150) -> None:
        self.calls += 1
        self.classify += 1
        self.estimated_input_tokens += tokens_in
        self.estimated_output_tokens += tokens_out

    def add_profile(self, tokens_in: int = 1200, tokens_out: int = 400) -> None:
        self.calls += 1
        self.export_profile += 1
        self.estimated_input_tokens += tokens_in
        self.estimated_output_tokens += tokens_out

    def add_phase1(self, tokens_in: int = 4000, tokens_out: int = 2000) -> None:
        self.calls += 1
        self.phase1_compile += 1
        self.estimated_input_tokens += tokens_in
        self.estimated_output_tokens += tokens_out

    def add_market_understanding(self, tokens_in: int = 3500, tokens_out: int = 2500) -> None:
        self.calls += 1
        self.market_understanding += 1
        self.estimated_input_tokens += tokens_in
        self.estimated_output_tokens += tokens_out

    def summary(self, *, provider: str = "opencode", model: str = "") -> dict:
        cost_usd = estimate_cost_usd(
            provider, model, self.estimated_input_tokens, self.estimated_output_tokens
        )
        return {
            "llm_calls_total": self.calls,
            "phase1_compile_calls": self.phase1_compile,
            "market_understanding_calls": self.market_understanding,
            "classify_calls": self.classify,
            "export_profile_calls": self.export_profile,
            "smart_crawl_chunk_calls": self.smart_crawl_chunks,
            "estimated_input_tokens": self.estimated_input_tokens,
            "estimated_output_tokens": self.estimated_output_tokens,
            "estimated_cost_usd": round(cost_usd, 4),
            "provider": provider,
            "model": model,
        }


_meter_var: contextvars.ContextVar[LlmMeter | None] = contextvars.ContextVar(
    "llm_meter", default=None
)


def get_meter() -> LlmMeter:
    meter = _meter_var.get()
    if meter is None:
        meter = LlmMeter()
        _meter_var.set(meter)
    return meter


def reset_meter() -> LlmMeter:
    meter = LlmMeter()
    _meter_var.set(meter)
    return meter
