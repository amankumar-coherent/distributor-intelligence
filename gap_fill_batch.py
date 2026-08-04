#!/usr/bin/env python3
"""
gap_fill_batch.py — automated gap-fill for any batch (10 or 10,000 markets).

Fills missing Founded, HQ, Cities/Regions, Employees, Turnover, LinkedIn using
free sources: search backfill, LinkedIn, Owler snippets, Crunchbase snippets,
Wikipedia API, SEC EDGAR, locations search. Gap-fill only — never overwrites good data.

Usage:
  python gap_fill_batch.py --batch-id medical-batch-10 --yes
  python gap_fill_batch.py --batch-id mega-batch --max-concurrent 10 --yes
  python gap_fill_batch.py --batch-id medical-batch-10 --force --yes
  python gap_fill_batch.py --batch-id medical-batch-10 --sources backfill,linkedin,owler,wikipedia --yes
  python gap_fill_batch.py --batch-id medical-batch-10 --no-push-neon --yes
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))


def main() -> None:
    parser = argparse.ArgumentParser(description="Batch gap-fill for missing firmographics")
    parser.add_argument("--batch-id", "-b", required=True, help="Batch id (same as run_batch.py)")
    parser.add_argument("--max-concurrent", "-c", type=int, default=3, help="Markets in parallel (default: 3)")
    parser.add_argument(
        "--gap-concurrent",
        type=int,
        default=4,
        help="Companies in parallel per market (default: 4)",
    )
    parser.add_argument(
        "--sources",
        default="all",
        help="Comma sources: backfill,linkedin,owler,crunchbase,wikipedia,sec,locations (default: all)",
    )
    parser.add_argument("--force", action="store_true", help="Re-run gap-fill even if market cache says done")
    parser.add_argument("--no-fmcg-xlsx", action="store_true", help="Skip Excel regeneration")
    parser.add_argument("--no-csv", action="store_true", help="Skip CSV regeneration")
    parser.add_argument("--limit", type=int, default=None, help="Process only N markets (smoke test)")
    parser.add_argument(
        "--no-push-neon",
        action="store_true",
        help="Skip merge-push to Neon after gap-fill (default: push when DATABASE_URL is set)",
    )
    parser.add_argument("--yes", "-y", action="store_true", help="Skip confirmation")
    args = parser.parse_args()

    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.utils.output_filter import install_stderr_filter

    install_stderr_filter()
    apply_env_overrides()

    from vendor_intel.enrichment.gap_fill.batch import gap_fill_status, gap_fill_batch

    status = gap_fill_status(args.batch_id)
    print(f"\n  [gap_fill] Batch: {args.batch_id}")
    print(f"  Done markets in registry: {status['done_markets']}")
    print(f"  Gap-fill completed: {status['gap_fill_completed']}")
    print(f"  Gap-fill pending: {status['gap_fill_pending']}")
    if status.get("no_json"):
        print(f"  Missing JSON exports: {status['no_json']}")

    pending = status["gap_fill_pending"]
    if args.force:
        pending = status["done_markets"] - status.get("no_json", 0)

    if pending == 0 and not args.force:
        print("  [gap_fill] Nothing pending — all done markets already gap-filled.")
        return

    if not args.yes:
        try:
            confirm = input(f"  Run gap-fill on ~{pending} market(s)? [Enter=go]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  Cancelled.")
            return
        if confirm:
            print("  Cancelled.")
            return

    try:
        summary = gap_fill_batch(
            args.batch_id,
            max_concurrent=args.max_concurrent,
            gap_concurrent=args.gap_concurrent,
            sources=None if args.sources == "all" else args.sources,
            force=args.force,
            fmcg_xlsx=not args.no_fmcg_xlsx,
            also_csv=not args.no_csv,
            limit=args.limit,
            push_neon=not args.no_push_neon,
        )
    finally:
        try:
            from vendor_intel.clients.ddg_worker_pool import shutdown_ddg_pool

            shutdown_ddg_pool(wait=True)
        except Exception:
            pass

    print(f"\n  [gap_fill] Finished. Completed: {summary.get('gap_fill_completed')} / {summary.get('done_markets')}")


if __name__ == "__main__":
    main()
