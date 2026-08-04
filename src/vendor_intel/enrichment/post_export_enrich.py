"""Post-export enrichment: deep crawl, LinkedIn, Owler, re-export per market."""
from __future__ import annotations

import asyncio
import json
import socket
import time
from pathlib import Path
from typing import Any

from vendor_intel.enrichment.linkedin_enrich import enrich_company_linkedin
from vendor_intel.enrichment.owler_enrich import enrich_companies_owler


def is_internet_available() -> bool:
    try:
        socket.setdefaulttimeout(3)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("8.8.8.8", 53))
        return True
    except OSError:
        return False


async def wait_for_internet(idx: int, total: int, company_name: str) -> None:
    from vendor_intel.utils.scrape_logging import alog

    if is_internet_available():
        return
    await alog(
        f"[{idx}/{total}] INTERNET DISCONNECTED while processing {company_name}! Pausing execution...",
        level="WARN",
    )
    while not is_internet_available():
        await asyncio.sleep(5)
    await alog(f"[{idx}/{total}] Internet connection restored! Resuming execution...", level="NOTE")


def norm_domain(row: dict) -> str:
    dom = str(row.get("domain") or row.get("website") or "").strip().lower()
    return dom.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0]


def crawl_is_deep(row: dict) -> bool:
    from vendor_intel.export.fmcg_landscape import is_plausible_location

    src = str(row.get("enrichment_source") or "").lower()
    if src == "smart_crawl":
        return True
    crawl = row.get("_crawl") or {}
    for field in ("headquarters", "contact_address"):
        val = str(crawl.get(field) or "").strip()
        if val and is_plausible_location(val, allow_country_only=True):
            return True
    for field in ("offices", "manufacturing_locations"):
        val = str(crawl.get(field) or "").strip()
        if val and is_plausible_location(val, allow_country_only=False):
            return True
    return False


def has_good_hq(smart_data: dict) -> bool:
    from vendor_intel.export.fmcg_landscape import crawl_snapshot, is_plausible_location

    snap = crawl_snapshot(smart_data)
    for field in ("headquarters", "contact_address"):
        val = str(snap.get(field) or "").strip()
        if val and is_plausible_location(val, allow_country_only=True):
            return True
    return False


def apply_crawl(row: dict, smart_data: dict) -> None:
    from vendor_intel.export.fmcg_landscape import crawl_snapshot
    from vendor_intel.pipeline.csv_fields import format_data_sources

    row["_crawl"] = crawl_snapshot(smart_data)
    row["_crawl_raw"] = smart_data
    src = str(smart_data.get("source") or "smart_crawl")
    if smart_data.get("error") and not row["_crawl"].get("crawl_text"):
        row["enrichment_source"] = f"error:{smart_data.get('error')}"[:40]
    else:
        row["enrichment_source"] = src if src != "ssc" else "smart_crawl"

    products = str(row["_crawl"].get("products") or "").strip()
    if products and not str(row.get("key_products") or "").strip():
        from vendor_intel.pipeline.csv_fields import sanitize_key_products

        name = str(row.get("company") or row.get("brand") or "")
        cleaned = sanitize_key_products(products, company=name, brand=str(row.get("brand") or ""))
        if cleaned:
            row["key_products"] = cleaned

    row["data_sources"] = format_data_sources(
        str(row.get("discovery_source") or ""),
        row["enrichment_source"],
        smart_data,
    )


async def crawl_domain(domain: str, *, supplement: bool) -> dict:
    from vendor_intel.enrichment.smart_enrichment import _load_smart_crawl, supplement_crawl

    smart_crawl = _load_smart_crawl()
    smart_data = await smart_crawl(domain, mode="company")
    if supplement and not has_good_hq(smart_data):
        smart_data = await supplement_crawl(domain, smart_data)
    return smart_data


def format_crawl_summary(smart_data: dict, row: dict) -> str:
    from vendor_intel.export.fmcg_landscape import crawl_snapshot, is_plausible_location

    stats = smart_data.get("extraction_stats") or {}
    snap = crawl_snapshot(smart_data)
    hq = str(snap.get("headquarters") or snap.get("contact_address") or "").strip()
    hq_ok = bool(hq and is_plausible_location(hq, allow_country_only=True))
    pages = smart_data.get("pages_fetched") or len(smart_data.get("pages") or [])
    chunks = stats.get("llm_chunks_sent") or 0
    tin = stats.get("llm_tokens_in") or 0
    tout = stats.get("llm_tokens_out") or 0
    founded = snap.get("founded_year") or ""
    employees = str(snap.get("employee_count") or "").strip()
    products = str(snap.get("products") or "").strip()
    err = str(smart_data.get("error") or "").strip()
    src = str(row.get("enrichment_source") or "")

    parts = [f"source={src}", f"pages={pages}", f"llm_chunks={chunks}"]
    if tin:
        parts.append(f"tokens={tin}+{tout}")
    parts.append(f"HQ={'yes' if hq_ok else 'no'}")
    if founded:
        parts.append(f"founded={founded}")
    if employees:
        parts.append(f"employees={employees[:30]}")
    if products:
        parts.append(f"products={len(products.split(','))} items")
    if err:
        parts.append(f"error={err[:40]}")
    return " | ".join(parts)


def save_checkpoint(result: dict, json_path: Path | None) -> None:
    if json_path:
        try:
            json_path.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
        except Exception:
            pass


async def re_enrich(
    result: dict,
    *,
    limit: int | None = None,
    concurrent: int = 4,
    force: bool = False,
    supplement: bool = True,
    linkedin: bool = True,
    log_progress: bool = True,
    json_path: Path | None = None,
) -> tuple[int, int]:
    """Deep crawl + optional LinkedIn per company."""
    from vendor_intel.utils.output_filter import install_stderr_filter
    from vendor_intel.utils.scrape_logging import alog, append_progress, elapsed_minutes, log, log_section

    install_stderr_filter()

    rows = [r for r in (result.get("relevant_companies") or []) if r.get("is_relevant", True)]
    if not rows:
        log("No relevant companies in JSON.", level="WARN")
        return 0, 0

    todo: list[dict] = []
    for r in rows:
        dom = norm_domain(r)
        if not dom:
            continue
        if not force and crawl_is_deep(r):
            continue
        todo.append(r)

    if limit is not None:
        todo = todo[: max(0, limit)]

    if not todo:
        log("All companies already deep-enriched (use --force to re-crawl).", level="WARN")
        return 0, len(rows)

    total = len(todo)
    log_section("DEEP CRAWL (smart_crawl)")
    log(
        f"Queue: {total} companies (of {len(rows)} relevant) | "
        f"concurrent={concurrent} | supplement={supplement} | linkedin={linkedin}"
    )

    sem = asyncio.Semaphore(max(1, concurrent))
    ok = 0
    fail = 0
    done_lock = asyncio.Lock()

    async def _one(row: dict, idx: int) -> None:
        nonlocal ok, fail
        name = str(row.get("company") or row.get("brand") or "")
        dom = norm_domain(row)
        t_start = time.perf_counter()
        async with sem:
            await alog(f"[{idx}/{total}] START  {name[:55]}  ({dom})")
            try:
                await wait_for_internet(idx, total, name)
                try:
                    smart_data = await crawl_domain(dom, supplement=supplement)
                except Exception:
                    if not is_internet_available():
                        await wait_for_internet(idx, total, name)
                        smart_data = await crawl_domain(dom, supplement=supplement)
                    else:
                        raise
                apply_crawl(row, smart_data)

                if linkedin:
                    try:
                        await enrich_company_linkedin(
                            row,
                            smart_data=smart_data,
                            idx=idx,
                            total=total,
                            log_fn=alog,
                        )
                    except Exception as e:
                        await alog(f"[{idx}/{total}] LinkedIn error: {e}", level="ERROR")

                elapsed = time.perf_counter() - t_start
                summary = format_crawl_summary(smart_data, row)
                async with done_lock:
                    ok += 1
                    n_done = ok + fail
                    rate = elapsed_minutes() / max(n_done, 1)
                    eta_min = rate * (total - n_done)
                    await alog(
                        f"[{idx}/{total}] DONE   {name[:45]}  ({dom})  "
                        f"{elapsed:.0f}s | {summary}"
                    )
                    await alog(
                        f"PROGRESS {n_done}/{total} ok={ok} fail={fail} | "
                        f"elapsed={elapsed_minutes():.1f}m ETA≈{eta_min:.0f}m"
                    )
                    if log_progress:
                        from vendor_intel.export.fmcg_landscape import crawl_snapshot

                        snap = crawl_snapshot(smart_data)
                        append_progress(
                            {
                                "n": n_done,
                                "total": total,
                                "idx": idx,
                                "company": name,
                                "domain": dom,
                                "elapsed_s": round(elapsed, 1),
                                "enrichment_source": row.get("enrichment_source"),
                                "headquarters": snap.get("headquarters") or "",
                                "founded_year": snap.get("founded_year") or "",
                                "employee_count": snap.get("employee_count") or "",
                                "pages_fetched": smart_data.get("pages_fetched"),
                                "llm_chunks": (smart_data.get("extraction_stats") or {}).get(
                                    "llm_chunks_sent"
                                ),
                                "llm_tokens_in": (smart_data.get("extraction_stats") or {}).get(
                                    "llm_tokens_in"
                                ),
                            }
                        )
                    save_checkpoint(result, json_path)
            except Exception as exc:
                elapsed = time.perf_counter() - t_start
                async with done_lock:
                    fail += 1
                    n_done = ok + fail
                    await alog(
                        f"[{idx}/{total}] FAIL   {name[:45]}  ({dom})  "
                        f"{elapsed:.0f}s | error={exc}",
                        level="ERROR",
                    )
                    await alog(
                        f"PROGRESS {n_done}/{total} ok={ok} fail={fail} | "
                        f"elapsed={elapsed_minutes():.1f}m",
                        level="WARN",
                    )
                    save_checkpoint(result, json_path)

    await asyncio.gather(*[_one(r, i) for i, r in enumerate(todo, 1)])
    log_section("DEEP CRAWL COMPLETE")
    log(f"Finished: {ok} ok, {fail} failed (of {total})")
    return ok, fail


def re_export_market(
    result: dict,
    json_path: Path,
    *,
    market_type: str = "general",
    fmcg_xlsx: bool = True,
) -> dict[str, str]:
    """Write JSON, CSV, and optional XLSX after enrichment."""
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx
    from vendor_intel.pipeline.orchestrator import save_pipeline_csv

    paths: dict[str, str] = {}
    json_path.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
    paths["_json_path"] = str(json_path)

    csv_path = json_path.with_suffix(".csv")
    save_pipeline_csv(result, str(csv_path))
    paths["_csv_path"] = str(csv_path)

    if fmcg_xlsx:
        try:
            xlsx_path = json_path.with_suffix(".xlsx")
            paths["_xlsx_path"] = export_fmcg_xlsx(result, xlsx_path, market_type=market_type)
        except Exception as exc:
            print(f"  [post_enrich] Excel export failed: {exc}", flush=True)

    return paths


async def post_export_enrich_market(
    result: dict,
    *,
    json_path: Path,
    market_type: str = "general",
    concurrent: int = 4,
    supplement: bool = True,
    linkedin: bool = True,
    owler: bool = True,
    search_backfill: bool = False,
    fmcg_xlsx: bool = True,
    force: bool = False,
) -> dict[str, Any]:
    """Full post-export pass: deep crawl, LinkedIn, Owler, optional backfill, re-export."""
    from vendor_intel.utils.scrape_logging import log, log_section

    t0 = time.perf_counter()
    log_section(f"POST-EXPORT ENRICH: {json_path.parent.name}")

    crawl_ok, crawl_fail = await re_enrich(
        result,
        concurrent=concurrent,
        force=force,
        supplement=supplement,
        linkedin=linkedin,
        json_path=json_path,
    )
    result["re_enrich_ok"] = crawl_ok
    result["re_enrich_failed"] = crawl_fail

    rows = [r for r in (result.get("relevant_companies") or []) if r.get("is_relevant", True)]

    owler_updated = 0
    if owler and rows:
        log_section("OWLER ENRICHMENT")
        owler_updated, owler_total = await enrich_companies_owler(
            rows, concurrent=min(10, concurrent * 2)
        )
        result["owler_enriched"] = owler_updated
        result["owler_enrich_total"] = owler_total
        save_checkpoint(result, json_path)

    if search_backfill and rows:
        from vendor_intel.enrichment.search_backfill import backfill_companies

        log_section("SEARCH BACKFILL")
        bf_ok, bf_fail = await backfill_companies(result)
        result["search_backfill_ok"] = bf_ok
        result["search_backfill_failed"] = bf_fail
        save_checkpoint(result, json_path)

    result["post_enrich_completed_at"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    result["post_enrich_elapsed_minutes"] = round((time.perf_counter() - t0) / 60, 2)

    export_paths = re_export_market(
        result, json_path, market_type=market_type, fmcg_xlsx=fmcg_xlsx
    )
    result.update(export_paths)

    log(
        f"Post-enrich complete: crawl ok={crawl_ok} fail={crawl_fail}, "
        f"owler updated={owler_updated}, elapsed={result['post_enrich_elapsed_minutes']}m"
    )
    return result
