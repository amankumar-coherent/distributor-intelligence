"""Concurrent multi-market worker pool.

N persistent "slots" each loop claim -> run -> claim against the shared
SQLite queue (batch/db.py), all on ONE shared asyncio event loop — so the
existing browser-concurrency safety net (crawler/smart_crawl.py's
module-level `_BROWSER_SEM`, sized by CRAWL_BROWSER_CONCURRENCY) and any
future search rate limiter stay correctly host-wide instead of being
silently duplicated per market. See the "Worker pool: how markets get
divided across workers" section of the batch-scaling plan — this is the
"vertical" (in-process) half of that design. The "horizontal" half (multiple
OS processes sharing one registry) needs no new code here: db.claim_next_market()'s
atomic UPDATE...RETURNING is already safe for concurrent callers across
processes, not just within one — running two `run_batch.py` invocations
against the same --batch-id today already divides the work correctly, we
just don't have a launcher that spawns N of them for you yet.
"""
from __future__ import annotations

import asyncio
import os
from typing import Any

from vendor_intel.batch import db
from vendor_intel.batch.worker import process_market_async
from vendor_intel.config import Settings


def scale_settings_for_concurrency(settings: Settings, max_concurrent_markets: int) -> Settings:
    """Divide per-market enrich/classify concurrency across the number of
    markets running at once, so N concurrent markets x per-market concurrency
    doesn't multiply past what CRAWL_BROWSER_CONCURRENCY or the LLM provider
    can actually sustain. A floor of 2 keeps each market's own enrichment from
    stalling entirely at high market concurrency.
    """
    if max_concurrent_markets <= 1:
        return settings
    scaled_enrich = max(2, settings.pipeline_enrich_concurrent // max_concurrent_markets)
    scaled_classify = max(2, settings.pipeline_classify_concurrent // max_concurrent_markets)
    return settings.model_copy(
        update={
            "pipeline_enrich_concurrent": scaled_enrich,
            "pipeline_classify_concurrent": scaled_classify,
        }
    )


async def _worker_slot(
    batch_id: str,
    worker_id: str,
    settings: Settings,
    *,
    fmcg_xlsx: bool,
    full_enrich: bool,
    enrich_concurrent: int,
    search_backfill: bool,
    with_audit: bool,
    audit_concurrent: int,
    gap_fill: bool,
    gap_fill_concurrent: int,
    gap_fill_sources: str | None,
    push_neon: bool,
    post_market: dict[str, Any] | None,
    resume_checkpoint: bool,
    limit: int | None,
    counter: dict[str, int],
) -> None:
    while True:
        if limit is not None:
            # Check-and-reserve with no `await` in between: since asyncio is
            # cooperative, this can't interleave with another slot's check,
            # so concurrent slots can never together reserve past `limit`
            # (unlike checking before claim and incrementing after, which
            # races: two slots can both pass the check before either claims).
            if counter["ran"] >= limit:
                return
            counter["ran"] += 1

        market = await asyncio.to_thread(db.claim_next_market, batch_id, worker_id)
        if market is None:
            if limit is not None:
                counter["ran"] -= 1  # give back the reservation, queue was empty
            return
        await process_market_async(
            market,
            settings,
            fmcg_xlsx=fmcg_xlsx,
            full_enrich=full_enrich,
            enrich_concurrent=enrich_concurrent,
            search_backfill=search_backfill,
            with_audit=with_audit,
            audit_concurrent=audit_concurrent,
            gap_fill=gap_fill,
            gap_fill_concurrent=gap_fill_concurrent,
            gap_fill_sources=gap_fill_sources,
            push_neon=push_neon,
            post_market=post_market,
            resume_checkpoint=resume_checkpoint,
        )


async def run_batch_async(
    batch_id: str,
    settings: Settings,
    *,
    max_concurrent_markets: int = 1,
    fmcg_xlsx: bool = True,
    full_enrich: bool = False,
    enrich_concurrent: int = 4,
    search_backfill: bool = False,
    with_audit: bool = False,
    audit_concurrent: int = 8,
    gap_fill: bool = False,
    gap_fill_concurrent: int = 4,
    gap_fill_sources: str | None = None,
    push_neon: bool = False,
    post_market: dict[str, Any] | None = None,
    resume_checkpoint: bool = True,
    limit: int | None = None,
) -> dict[str, int]:
    """Run every pending market in `batch_id` using `max_concurrent_markets`
    concurrent worker slots on one shared event loop. Returns final
    db.batch_status_counts() for the batch.
    """
    reaped = await asyncio.to_thread(db.reap_stale_running, batch_id)
    if reaped:
        print(
            f"  [batch] Reaped {reaped} market(s) stuck 'running' from a prior crashed process.",
            flush=True,
        )
    if resume_checkpoint:
        print(
            "  [batch] Checkpoint resume ON — reuse Phase 1/2 artifacts when present "
            "(continues at enrich/classify instead of re-discovering).",
            flush=True,
        )

    n_slots = max(1, int(max_concurrent_markets))
    effective_enrich_concurrent = enrich_concurrent
    if full_enrich and n_slots >= 5 and enrich_concurrent > 2:
        effective_enrich_concurrent = 2
        print(
            f"  [batch] full-enrich: scaled enrich-concurrent down to {effective_enrich_concurrent} "
            f"(max-concurrent markets={n_slots})",
            flush=True,
        )
    if full_enrich and n_slots > 3:
        print(
            f"  [batch] WARN: --full-enrich with --max-concurrent {n_slots} is heavy on memory/rate limits; "
            f"consider --max-concurrent 3 for large batches.",
            flush=True,
        )
    if with_audit and n_slots > 3:
        print(
            f"  [batch] WARN: --with-audit with --max-concurrent {n_slots} is heavy on search/LLM rate limits; "
            f"consider --max-concurrent 3 for large batches.",
            flush=True,
        )
    if gap_fill and n_slots > 3:
        print(
            f"  [batch] WARN: --gap-fill with --max-concurrent {n_slots} is heavy on search/LLM rate limits; "
            f"consider --max-concurrent 3 for large batches.",
            flush=True,
        )
    await asyncio.to_thread(db.set_batch_concurrency, batch_id, n_slots)
    scaled = scale_settings_for_concurrency(settings, n_slots)
    if n_slots > 1:
        print(
            f"  [batch] Concurrency: {n_slots} markets at once "
            f"(per-market enrich={scaled.pipeline_enrich_concurrent}, "
            f"classify={scaled.pipeline_classify_concurrent}, scaled down from "
            f"{settings.pipeline_enrich_concurrent}/{settings.pipeline_classify_concurrent})",
            flush=True,
        )

    pid = os.getpid()
    counter: dict[str, int] = {"ran": 0}
    slots = [
        _worker_slot(
            batch_id, f"main-{pid}-slot{i}", scaled,
            fmcg_xlsx=fmcg_xlsx,
            full_enrich=full_enrich,
            enrich_concurrent=effective_enrich_concurrent,
            search_backfill=search_backfill,
            with_audit=with_audit,
            audit_concurrent=audit_concurrent,
            gap_fill=gap_fill,
            gap_fill_concurrent=gap_fill_concurrent,
            gap_fill_sources=gap_fill_sources,
            push_neon=push_neon,
            post_market=post_market,
            resume_checkpoint=resume_checkpoint,
            limit=limit,
            counter=counter,
        )
        for i in range(n_slots)
    ]
    await asyncio.gather(*slots)

    return await asyncio.to_thread(db.batch_status_counts, batch_id)


def run_batch(
    batch_id: str,
    settings: Settings,
    *,
    max_concurrent_markets: int = 1,
    fmcg_xlsx: bool = True,
    full_enrich: bool = False,
    enrich_concurrent: int = 4,
    search_backfill: bool = False,
    with_audit: bool = False,
    audit_concurrent: int = 8,
    gap_fill: bool = False,
    gap_fill_concurrent: int = 4,
    gap_fill_sources: str | None = None,
    push_neon: bool = False,
    post_market: dict[str, Any] | None = None,
    resume_checkpoint: bool = True,
    limit: int | None = None,
) -> dict[str, int]:
    """Sync entry point for run_batch.py — creates the one shared event loop
    the whole batch runs on."""
    return asyncio.run(
        run_batch_async(
            batch_id,
            settings,
            max_concurrent_markets=max_concurrent_markets,
            fmcg_xlsx=fmcg_xlsx,
            full_enrich=full_enrich,
            enrich_concurrent=enrich_concurrent,
            search_backfill=search_backfill,
            with_audit=with_audit,
            audit_concurrent=audit_concurrent,
            gap_fill=gap_fill,
            gap_fill_concurrent=gap_fill_concurrent,
            gap_fill_sources=gap_fill_sources,
            push_neon=push_neon,
            post_market=post_market,
            resume_checkpoint=resume_checkpoint,
            limit=limit,
        )
    )
