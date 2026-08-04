"""Process one already-claimed market row end-to-end: run the pipeline, write
its exports, record done/failed. Shared by run_batch.py's sequential path
(one asyncio.run() per market) and batch/orchestrator.py's concurrent worker
pool (awaited directly on one shared event loop) — same code, same behavior,
the only difference is how many of these run at once.
"""
from __future__ import annotations

import asyncio
import json as _json
import time
from pathlib import Path
from typing import Any

from vendor_intel.batch import db
from vendor_intel.batch.failure_classifier import classify_failure
from vendor_intel.config import Settings, _project_root


async def process_market_async(
    market: Any,
    settings: Settings,
    *,
    fmcg_xlsx: bool,
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
) -> None:
    """Run `market` (a db.claim_next_market() row) and mark it done/failed."""
    from vendor_intel.pipeline.geo_limits import pipeline_limits
    from vendor_intel.pipeline.run_and_export import execute_and_export_async

    market_id = market["id"]
    query = market["query"]
    country = market["country"]
    slug = market["market_slug"]
    worker_id = market["worker_id"] or "?"
    market_type = market["market_type"] or "general"
    sections = _json.loads(market["sections_json"] or "[]")
    # Honor CSV_OUTPUT_DIR so batches can write to a fresh root (e.g. output/run5_demo10)
    out_base = Path(getattr(settings, "csv_output_dir", None) or "output")
    if not out_base.is_absolute():
        out_base = _project_root() / out_base
    output_dir = out_base / slug

    lim = pipeline_limits(settings, recall=(settings.pipeline_profile == "recall"), country=country)
    enrich_cap = int(lim["enrich"])
    classify_cap = int(lim["discover"])

    print(
        f"\n{'=' * 60}\n  [batch:{worker_id}] {query} ({country}) -> {output_dir.as_posix()}/\n{'=' * 60}",
        flush=True,
    )
    if resume_checkpoint:
        print(
            f"  [batch:{worker_id}] checkpoint resume ON "
            f"(skip Phase 1/2 when output/phase1 + output/phase2 artifacts exist)",
            flush=True,
        )
    t0 = time.perf_counter()
    try:
        result = await execute_and_export_async(
            query,
            country,
            settings,
            output_dir,
            enrich_limit=enrich_cap,
            classify_limit=classify_cap,
            sections=sections,
            fmcg_xlsx=fmcg_xlsx,
            market_type=market_type,
            resume_checkpoint=resume_checkpoint,
            market_slug=slug,
        )
        if full_enrich:
            from pathlib import Path

            from vendor_intel.enrichment.post_export_enrich import post_export_enrich_market

            json_path = Path(result.get("_json_path") or "")
            if not json_path.is_file():
                slug_file = output_dir / f"{slug}.json"
                if slug_file.is_file():
                    json_path = slug_file
            if json_path.is_file():
                print(
                    f"  [batch:{worker_id}] Post-export enrich (LinkedIn + Owler): {json_path.name}",
                    flush=True,
                )
                result = await post_export_enrich_market(
                    result,
                    json_path=json_path,
                    market_type=market_type,
                    concurrent=enrich_concurrent,
                    supplement=True,
                    linkedin=True,
                    owler=True,
                    search_backfill=search_backfill,
                    fmcg_xlsx=fmcg_xlsx,
                )
            else:
                print(
                    f"  [batch:{worker_id}] WARN: no JSON path for post-export enrich ({slug})",
                    flush=True,
                )

        if with_audit:
            from pathlib import Path

            from vendor_intel.enrichment.market_audit import audit_market_exports

            json_path = Path(result.get("_json_path") or output_dir / f"{slug}.json")
            csv_path = Path(result.get("_csv_path") or json_path.with_suffix(".csv"))
            xlsx_path = Path(result.get("_xlsx_path") or json_path.with_suffix(".xlsx"))
            print(
                f"  [batch:{worker_id}] Market audit (web search + LLM): {query[:50]}",
                flush=True,
            )
            audit_stats = await audit_market_exports(
                output_dir=output_dir,
                market_query=query,
                csv_path=csv_path if csv_path.is_file() else None,
                xlsx_path=xlsx_path if xlsx_path.is_file() else None,
                json_path=json_path if json_path.is_file() else None,
                result=result,
                market_type=market_type,
                concurrent=audit_concurrent,
                fmcg_xlsx=fmcg_xlsx,
            )
            if audit_stats.get("remove") and result.get("relevant_companies") is not None:
                n = len(result.get("relevant_companies") or [])
                print(f"  [batch:{worker_id}] After audit: {n} companies remain ({slug})", flush=True)

        if gap_fill:
            from pathlib import Path

            from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_market

            json_path = Path(result.get("_json_path") or output_dir / f"{slug}.json")
            print(
                f"  [batch:{worker_id}] Gap-fill (Google AI→DDGS→…→OpenAI residual): {query[:50]}",
                flush=True,
            )
            gf_stats = await gap_fill_market(
                result,
                json_path=json_path,
                output_dir=output_dir,
                market_type=market_type,
                fmcg_xlsx=fmcg_xlsx,
                concurrent=gap_fill_concurrent,
                sources=gap_fill_sources,
            )
            print(
                f"  [batch:{worker_id}] Gap-fill: {gf_stats.get('updated', 0)} updated, "
                f"{gf_stats.get('still_missing_gaps', 0)} still have gaps ({slug})",
                flush=True,
            )

        if push_neon:
            from vendor_intel.neon_push import push_result_to_neon

            await asyncio.to_thread(
                push_result_to_neon,
                result,
                label=f"neon:{worker_id}",
            )

        if post_market:
            from vendor_intel.pipeline.post_market import run_post_market_pipeline

            pm_kwargs = dict(post_market)
            pm_kwargs.setdefault("resume", True)
            if push_neon and "push_neon" not in pm_kwargs:
                pm_kwargs["push_neon"] = True
            elif not push_neon and pm_kwargs.get("push_neon") is True:
                # Respect explicit --no-push-neon via caller overriding post_market dict.
                pass
            print(
                f"  [batch:{worker_id}] Post-market pipeline: {query[:50]} "
                f"(resume={'on' if pm_kwargs.get('resume', True) else 'off'})",
                flush=True,
            )
            pm_stats = await run_post_market_pipeline(output_dir, **pm_kwargs)
            if pm_stats.get("error"):
                print(
                    f"  [batch:{worker_id}] WARN post-market: {pm_stats['error']}",
                    flush=True,
                )
                # Do not mark the market done when post-market could not even load JSON —
                # otherwise Founded/HQ stay empty forever while status=done.
                raise RuntimeError(f"post-market failed: {pm_stats['error']}")
            excel = pm_stats.get("excel") or {}
            primary = (excel.get("primary_landscape") or {}) if isinstance(excel, dict) else {}
            if primary.get("primary_xlsx"):
                result["_xlsx_path"] = primary["primary_xlsx"]
            if primary.get("primary_csv"):
                result["_csv_path"] = primary["primary_csv"]
            result["post_market_pipeline"] = pm_stats

        elapsed = time.perf_counter() - t0
        n = len(result.get("relevant_companies") or [])
        await asyncio.to_thread(
            db.mark_market_done,
            market_id,
            companies_exported=n,
            csv_path=result.get("_csv_path", ""),
            xlsx_path=result.get("_xlsx_path", ""),
            docx_path=result.get("_docx_path", ""),
            json_path=result.get("_json_path", ""),
        )
        print(f"  [batch:{worker_id}] OK — {n} companies in {elapsed / 60:.1f} min ({slug})", flush=True)
    except Exception as exc:
        import traceback

        traceback.print_exc()
        kind = classify_failure(exc)
        await asyncio.to_thread(
            db.mark_market_failed,
            market_id,
            error=f"{type(exc).__name__}: {exc}",
            kind=kind,
        )
        print(f"  [batch:{worker_id}] FAILED ({kind}): {exc} ({slug})", flush=True)
