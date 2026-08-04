"""Multi-market ChatGPT expand worker pool (steps 0–7 per market).

Runs the same ``run_chatgpt_expand`` pipeline for every pending market in a
batch registry — designed for 1…10k markets, not one hardcoded demo.
"""
from __future__ import annotations

import asyncio
import os
import time
from typing import Any

from vendor_intel.batch import db
from vendor_intel.pipeline.web_expand import default_output_dir, slugify


async def process_expand_market_async(
    market: Any,
    *,
    worker_id: str,
    target: int = 320,
    max_fill: int = 320,
    max_queries: int = 120,
    fill_backend: str = "search",
    ddgs_harvest: bool = False,
    seeds_only: bool = False,
    skip_recall: bool = False,
    use_apollo: bool = True,
    use_linkedin_enrich: bool = True,
    resume: bool = True,
    force: bool = False,
    openai_discover: bool = True,
    web_fill: bool = True,
) -> dict[str, Any]:
    from vendor_intel.pipeline.chatgpt_expand import run_chatgpt_expand

    market_id = int(market["id"])
    query = str(market["query"])
    country = str(market["country"] or "global")
    slug = str(market["market_slug"] or slugify(query, country))
    seeds_file = market["seeds_file"] if "seeds_file" in market.keys() else None
    market_type = str(market["market_type"] or "general")

    print(
        f"  [expand:{worker_id}] START {slug} | {query[:60]} | family≈{market_type}",
        flush=True,
    )
    t0 = time.perf_counter()
    try:
        result = await run_chatgpt_expand(
            query,
            country=country,
            target=target,
            batch="all",
            max_queries=max_queries,
            max_fill=max_fill,
            fill_backend=fill_backend,
            ddgs_harvest=ddgs_harvest,
            seeds_only=seeds_only,
            seeds_path=seeds_file,
            skip_recall=skip_recall,
            use_apollo=use_apollo,
            use_linkedin_enrich=use_linkedin_enrich,
            resume=resume,
            force=force,
            openai_discover=openai_discover,
            web_fill=web_fill,
            family=None,  # derived from query; never hardcode a demo market
        )
    except Exception as exc:  # noqa: BLE001
        err = f"{type(exc).__name__}: {exc}"
        print(f"  [expand:{worker_id}] FAIL {slug}: {err}", flush=True)
        await asyncio.to_thread(db.mark_market_failed, market_id, error=err, kind="transient")
        return {"slug": slug, "error": err}

    elapsed = time.perf_counter() - t0
    companies = int(
        result.get("total_after")
        or result.get("companies")
        or 0
    )
    xlsx = str(result.get("xlsx") or result.get("final_path") or "")
    audit_path = ""
    if result.get("checkpoint"):
        # audit json sits next to checkpoint under web_expand/<slug>/
        from pathlib import Path

        ckpt = Path(str(result["checkpoint"]))
        candidate = ckpt.parent / "chatgpt_expand_batch_all_audit.json"
        if candidate.is_file():
            audit_path = str(candidate)

    if not xlsx:
        xlsx = str(default_output_dir(query, country) / f"{slugify(query, country)}_FINAL.xlsx")

    if result.get("skipped"):
        print(
            f"  [expand:{worker_id}] SKIP done {slug} ({companies} companies) in {elapsed:.0f}s",
            flush=True,
        )
    else:
        print(
            f"  [expand:{worker_id}] DONE {slug} → {companies} companies in {elapsed / 60:.1f}m | {xlsx}",
            flush=True,
        )

    await asyncio.to_thread(
        db.mark_market_done,
        market_id,
        companies_exported=companies,
        xlsx_path=xlsx,
        json_path=audit_path,
    )
    return {
        "slug": slug,
        "skipped": bool(result.get("skipped")),
        "companies": companies,
        "xlsx": xlsx,
    }


async def _expand_slot(
    batch_id: str,
    worker_id: str,
    *,
    expand_opts: dict[str, Any],
    limit: int | None,
    counter: dict[str, int],
) -> None:
    while True:
        if limit is not None:
            if counter["ran"] >= limit:
                return
            counter["ran"] += 1

        market = await asyncio.to_thread(db.claim_next_market, batch_id, worker_id)
        if market is None:
            if limit is not None:
                counter["ran"] -= 1
            return
        await process_expand_market_async(market, worker_id=worker_id, **expand_opts)


async def run_expand_batch_async(
    batch_id: str,
    *,
    max_concurrent_markets: int = 1,
    limit: int | None = None,
    target: int = 320,
    max_fill: int = 320,
    max_queries: int = 120,
    fill_backend: str = "search",
    ddgs_harvest: bool = False,
    seeds_only: bool = False,
    skip_recall: bool = False,
    use_apollo: bool = True,
    use_linkedin_enrich: bool = True,
    resume: bool = True,
    force: bool = False,
    openai_discover: bool = True,
    web_fill: bool = True,
) -> dict[str, int]:
    """Run ChatGPT expand steps 0–7 for every pending market in ``batch_id``."""
    reaped = await asyncio.to_thread(db.reap_stale_running, batch_id, older_than_seconds=0)
    if reaped:
        print(f"  [expand-batch] Reaped {reaped} stuck running market(s) → pending", flush=True)

    n_slots = max(1, int(max_concurrent_markets))
    if n_slots > 2:
        print(
            f"  [expand-batch] WARN: {n_slots} concurrent markets is heavy on OpenAI/search; "
            "prefer 1–2 for 10k-scale runs.",
            flush=True,
        )
    await asyncio.to_thread(db.set_batch_concurrency, batch_id, n_slots)

    expand_opts = {
        "target": target,
        "max_fill": max_fill,
        "max_queries": max_queries,
        "fill_backend": fill_backend,
        "ddgs_harvest": ddgs_harvest,
        "seeds_only": seeds_only,
        "skip_recall": skip_recall,
        "use_apollo": use_apollo,
        "use_linkedin_enrich": use_linkedin_enrich,
        "resume": resume,
        "force": force,
        "openai_discover": openai_discover,
        "web_fill": web_fill,
    }
    print(
        f"  [expand-batch] Steps 0–7 per market | concurrent_markets={n_slots} | "
        f"fill_backend={fill_backend} | target={target} | "
        f"fill_conc={os.getenv('EXPAND_FILL_CONCURRENT') or os.getenv('GAP_FILL_CONCURRENT') or '4'} | "
        f"residual_conc={os.getenv('EXPAND_RESIDUAL_CONCURRENT') or '2'}",
        flush=True,
    )

    pid = os.getpid()
    counter: dict[str, int] = {"ran": 0}
    slots = [
        _expand_slot(
            batch_id,
            f"expand-{pid}-slot{i}",
            expand_opts=expand_opts,
            limit=limit,
            counter=counter,
        )
        for i in range(n_slots)
    ]
    await asyncio.gather(*slots)
    return await asyncio.to_thread(db.batch_status_counts, batch_id)


def run_expand_batch(
    batch_id: str,
    *,
    max_concurrent_markets: int = 1,
    limit: int | None = None,
    **kwargs: Any,
) -> dict[str, int]:
    return asyncio.run(
        run_expand_batch_async(
            batch_id,
            max_concurrent_markets=max_concurrent_markets,
            limit=limit,
            **kwargs,
        )
    )
