#!/usr/bin/env python3
r"""Unified high energy laser scrape: market research + company re-enrich + CSV/XLSX export.

Prerequisite: run discovery first:
  .venv\Scripts\python.exe run_high_energy_laser_market.py

Usage:
  .venv\Scripts\python.exe run_high_energy_laser_scrape.py --limit 10   # smoke test
  .venv\Scripts\python.exe run_high_energy_laser_scrape.py              # full run
  .venv\Scripts\python.exe run_high_energy_laser_scrape.py --skip-backfill
  .venv\Scripts\python.exe run_high_energy_laser_scrape.py --skip-market-research

Logs (second PowerShell):
  Get-Content output\high_energy_laser\scrape_live.log -Wait -Tail 40
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

OUTPUT_DIR = ROOT / "output" / "high_energy_laser"
JSON_PATH = OUTPUT_DIR / "global_high_energy_laser_market_global.json"
MARKET_JSON = OUTPUT_DIR / "market_research_high_energy_laser.json"
LIVE_LOG = OUTPUT_DIR / "scrape_live.log"
PROGRESS_LOG = OUTPUT_DIR / "scrape_progress.jsonl"
MARKET_TOPIC = "high energy laser"


def _clean_stale_exports() -> None:
    from vendor_intel.utils.scrape_logging import log

    for pattern in ("*.csv", "*.xlsx"):
        for path in OUTPUT_DIR.glob(pattern):
            try:
                path.unlink()
                log(f"Removed stale export: {path.name}")
            except OSError as exc:
                log(f"Could not remove {path.name}: {exc}", level="WARN")


def _check_env() -> None:
    from backend.config import sync_opencode_to_openai_env
    from vendor_intel.utils.scrape_logging import log

    sync_opencode_to_openai_env()
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    model = (os.getenv("OPENAI_MODEL") or os.getenv("MODEL_EXTRACT") or "gpt-4o-mini").strip()
    if not key:
        log("OPENAI_API_KEY not set — smart_crawl and search backfill will be limited.", level="WARN")
    else:
        log(f"OpenAI configured (model={model})")


def _save_json(result: dict) -> None:
    from vendor_intel.utils.scrape_logging import log

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    JSON_PATH.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
    log(f"Saved JSON: {JSON_PATH.resolve()}")


def _export_all(result: dict, *, also_csv: bool = True) -> None:
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx
    from vendor_intel.pipeline.orchestrator import save_pipeline_csv
    from vendor_intel.utils.scrape_logging import log, log_section

    log_section("EXPORT")
    xlsx_path = JSON_PATH.with_suffix(".xlsx")
    out = export_fmcg_xlsx(result, xlsx_path)
    log(f"Excel: {out}")
    if also_csv:
        csv_path = JSON_PATH.with_suffix(".csv")
        save_pipeline_csv(result, str(csv_path))
        log(f"CSV: {csv_path.resolve()}")


async def _run(
    *,
    limit: int | None,
    concurrent: int,
    skip_market_research: bool,
    skip_backfill: bool,
    force: bool,
) -> None:
    from re_enrich_fmcg import re_enrich
    from scrape_market_research import scrape_market_research
    from vendor_intel.enrichment.search_backfill import backfill_companies
    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.utils.scrape_logging import elapsed_minutes, log, log_section

    apply_env_overrides()
    _check_env()
    _clean_stale_exports()

    if not JSON_PATH.is_file():
        log(
            f"Master JSON not found: {JSON_PATH}\n"
            "Run discovery first: .venv\\Scripts\\python.exe run_high_energy_laser_market.py",
            level="ERROR",
        )
        sys.exit(1)

    result = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    n_companies = len([r for r in (result.get("relevant_companies") or []) if r.get("is_relevant", True)])
    log(f"Loaded {JSON_PATH.name} — {n_companies} relevant companies")

    if not skip_market_research:
        log_section("MARKET RESEARCH")
        market = await scrape_market_research(
            topic=MARKET_TOPIC,
            output_path=MARKET_JSON,
        )
        result["market_research"] = market
        log(f"Market research reports collected: {len(market)}")
    elif MARKET_JSON.is_file():
        result["market_research"] = json.loads(MARKET_JSON.read_text(encoding="utf-8"))
        log(f"Using existing market research: {MARKET_JSON.name}")

    t0 = time.perf_counter()
    ok, fail = await re_enrich(
        result,
        limit=limit,
        concurrent=concurrent,
        force=force,
        supplement=True,
        log_progress=True,
        json_path=JSON_PATH,
    )

    result["re_enrich_completed_at"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    result["re_enrich_ok"] = ok
    result["re_enrich_failed"] = fail
    _save_json(result)

    bf_ok, bf_fail = 0, 0
    if skip_backfill:
        log("Skipping search backfill (--skip-backfill)", level="NOTE")
    else:
        log_section("SEARCH BACKFILL")
        bf_ok, bf_fail = await backfill_companies(result, limit=limit)
        result["search_backfill_ok"] = bf_ok
        result["search_backfill_failed"] = bf_fail
        _save_json(result)

    elapsed = time.perf_counter() - t0
    _export_all(result, also_csv=True)
    log_section("RUN COMPLETE")
    log(
        f"Total {elapsed / 60:.1f} min | crawl {ok} ok / {fail} fail | "
        f"backfill {bf_ok} ok / {bf_fail} fail | wall={elapsed_minutes():.1f}m"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="High energy laser market research + company enrichment")
    parser.add_argument("--limit", type=int, default=None, help="Max companies to enrich (smoke test)")
    parser.add_argument("--concurrent", type=int, default=3, help="Parallel crawls (default 3)")
    parser.add_argument("--skip-market-research", action="store_true", help="Skip IMARC/SkyQuest/R&M scrape")
    parser.add_argument(
        "--skip-backfill",
        action="store_true",
        help="Skip DDGS search backfill (faster; smart_crawl data kept)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-crawl already enriched companies",
    )
    parser.add_argument("--log-file", default=str(LIVE_LOG))
    parser.add_argument("--progress-file", default=str(PROGRESS_LOG))
    args = parser.parse_args()

    from vendor_intel.utils.scrape_logging import setup_scrape_logging

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log_path = setup_scrape_logging(
        Path(args.log_file),
        progress_path=Path(args.progress_file),
    )
    from vendor_intel.utils.scrape_logging import log

    log(f"Market: Global High Energy Laser | Live log: {log_path}")

    asyncio.run(
        _run(
            limit=args.limit,
            concurrent=args.concurrent,
            skip_market_research=args.skip_market_research,
            skip_backfill=args.skip_backfill,
            force=args.force,
        )
    )


if __name__ == "__main__":
    main()
