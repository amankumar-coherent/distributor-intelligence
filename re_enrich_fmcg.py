#!/usr/bin/env python3
"""Deep-enrich an existing pipeline JSON with smart_crawl and regenerate FMCG Excel.

Fills Founded, Headquarters, Employees, Turnover, Core Categories, Key Brands, etc.
without re-running discovery.

Usage:
  python re_enrich_fmcg.py output/avocado_oil/global_avocado_oil_market_global.json
  python re_enrich_fmcg.py output/avocado_oil/global_avocado_oil_market_global.json --limit 20
  python re_enrich_fmcg.py output/avocado_oil/global_avocado_oil_market_global.json --force
  python re_enrich_fmcg.py output/avocado_oil/global_avocado_oil_market_global.json --search-backfill
  python re_enrich_fmcg.py output/avocado_oil/global_avocado_oil_market_global.json --also-csv
  python re_enrich_fmcg.py output/avocado_oil/global_avocado_oil_market_global.json --export-only

Requires LLM API key (same as pipeline). Expect ~1-2 min per company with smart_crawl.
"""
from __future__ import annotations

import argparse
import asyncio
import json
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

# Re-export shared helpers for backward compatibility
from vendor_intel.enrichment.linkedin_enrich import (  # noqa: E402
    apply_linkedin_and_summary,
    enrich_company_linkedin,
    extract_linkedin_info,
    fetch_linkedin_page,
    generate_polished_summary,
    get_linkedin_urls,
)
from vendor_intel.enrichment.post_export_enrich import (  # noqa: E402
    apply_crawl,
    crawl_domain,
    is_internet_available,
    norm_domain,
    re_enrich,
    re_export_market,
    wait_for_internet,
)

_apply_linkedin_and_summary = apply_linkedin_and_summary
_norm_domain = norm_domain
_apply_crawl = apply_crawl
_crawl_domain = crawl_domain


def main() -> None:
    parser = argparse.ArgumentParser(description="Deep-enrich pipeline JSON + regenerate FMCG Excel")
    parser.add_argument("json_path", help="Path to pipeline result JSON")
    parser.add_argument("--limit", type=int, default=None, help="Max companies to crawl")
    parser.add_argument("--concurrent", type=int, default=4, help="Parallel crawls (default 4)")
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-crawl even if enrichment_source is already smart_crawl",
    )
    parser.add_argument(
        "--search-backfill",
        action="store_true",
        help="DDGS + OpenAI snippet backfill for missing HQ/founded/employees",
    )
    parser.add_argument("--also-csv", action="store_true", help="Also write sectioned landscape CSV")
    parser.add_argument("--no-xlsx", action="store_true", help="Skip Excel regeneration")
    parser.add_argument(
        "--export-only",
        action="store_true",
        help="Regenerate Excel from JSON only (no deep crawl)",
    )
    args = parser.parse_args()

    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()

    jp = Path(args.json_path)
    if not jp.is_file():
        print(f"File not found: {jp}", file=sys.stderr)
        sys.exit(1)

    result = json.loads(jp.read_text(encoding="utf-8"))
    t0 = time.perf_counter()
    if args.export_only:
        ok, fail = 0, 0
        print("Export-only mode — skipping deep crawl.", flush=True)
    else:
        ok, fail = asyncio.run(
            re_enrich(
                result,
                limit=args.limit,
                concurrent=args.concurrent,
                force=args.force,
                supplement=True,
                json_path=jp,
            )
        )
        if args.search_backfill:
            from vendor_intel.enrichment.search_backfill import backfill_companies

            bf_ok, bf_fail = asyncio.run(backfill_companies(result, limit=args.limit))
            result["search_backfill_ok"] = bf_ok
            result["search_backfill_failed"] = bf_fail
    elapsed = time.perf_counter() - t0

    if not args.export_only:
        result["re_enrich_completed_at"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
        result["re_enrich_ok"] = ok
        result["re_enrich_failed"] = fail
        jp.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
        print(f"Updated JSON: {jp.resolve()}", flush=True)

    if not args.no_xlsx:
        from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx

        xlsx_path = jp.with_suffix(".xlsx")
        out = export_fmcg_xlsx(result, xlsx_path)
        print(f"Regenerated Excel: {out}", flush=True)

    if args.also_csv:
        from vendor_intel.pipeline.orchestrator import save_pipeline_csv

        csv_path = jp.with_suffix(".csv")
        save_pipeline_csv(result, str(csv_path))
        print(f"Regenerated CSV: {csv_path.resolve()}", flush=True)

    print(f"Finished in {elapsed / 60:.1f} min ({ok} crawled, {fail} failed)", flush=True)


if __name__ == "__main__":
    main()
