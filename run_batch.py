#!/usr/bin/env python3
"""
run_batch.py — generic multi-market batch driver.

Replaces the hand-copied run_<market>_market.py / run_<market>_scrape.py
scripts: point it at a market-list file and it loads every market into the
SQLite registry (data/batch.db) and runs them (sequentially by default, or
concurrently with --max-concurrent) writing CSV/DOCX/(optional XLSX)/JSON to
output/<market_slug>/ per market. Safe to Ctrl+C and resume: markets already
'done' are skipped, and any left 'running' by a killed process get reaped
back to 'pending' on the next start.

Usage:
  python run_batch.py --input queries/markets_batch.csv --batch-id pilot-50
  python run_batch.py --input queries/markets.txt --batch-id smoke-test --yes --limit 2
  python run_batch.py --batch-id pilot-50   # resume: no --input needed once loaded
  python run_batch.py --batch-id pilot-50 --max-concurrent 4   # run 4 markets at once
  python run_batch.py --input queries/markets.csv --batch-id mega-10k --mode bulk --yes
  python run_batch.py --input queries/medical_markets_batch_10.csv --batch-id medical-batch-10 \\
      --mode quality --max-concurrent 1 --yes
  python run_batch.py --input queries/market.txt --batch-id markets-1000-li-no-apollo \\
      --mode quality --max-concurrent 1 --no-apollo --no-name-resolve --push-neon --yes
  python run_batch.py --input queries/medical_markets_batch_10.csv --batch-id medical-batch-10 \\
      --profile recall --cap maximum --max-concurrent 3 --full-enrich --with-audit --gap-fill --push-neon --yes
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

# Must run before heavy logging: Windows Cursor PTYs raise OSError(22) on
# flush when the console handle breaks under concurrent print spam, which
# previously killed the whole run_batch process ("lost sys.stderr").
try:
    from vendor_intel.utils.safe_stdio import install_safe_stdio

    install_safe_stdio()
except Exception:
    pass


def main() -> None:
    parser = argparse.ArgumentParser(description="Vendor Intelligence — batch market runner")
    parser.add_argument(
        "--input", "-i", default=None,
        help="Market list file (.txt or .csv). Omit to resume an already-loaded --batch-id.",
    )
    parser.add_argument("--batch-id", "-b", required=True, help="Batch identifier (resume by reusing the same id)")
    parser.add_argument(
        "--mode", "-m",
        choices=("bulk", "standard", "quality", "full"),
        default=None,
        help="Pipeline preset: bulk (~5-12 min/market, 10k+ on 1 PC) | standard | quality (post-enrich) | full",
    )
    parser.add_argument("--profile", "-p", choices=("quality", "balanced", "recall"), default=None)
    parser.add_argument("--cap", choices=("bulk", "focused", "standard", "broad", "maximum"), default=None)
    parser.add_argument("--no-fmcg-xlsx", action="store_true", help="Skip the 30-column FMCG Excel export")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip the confirmation prompt")
    parser.add_argument("--limit", type=int, default=None, help="Stop after this many markets (smoke-testing)")
    parser.add_argument(
        "--max-concurrent", "-c", type=int, default=None,
        help="Run this many markets at once (default: mode preset, else 1). Per-market "
        "enrich/classify concurrency is automatically scaled down to compensate.",
    )
    parser.add_argument(
        "--full-enrich",
        action="store_true",
        help="After each market pipeline: deep crawl + LinkedIn + Owler + re-export CSV/XLSX",
    )
    parser.add_argument(
        "--enrich-concurrent",
        type=int,
        default=4,
        help="Parallel company enrichments per market during --full-enrich (default: 4)",
    )
    parser.add_argument(
        "--search-backfill",
        action="store_true",
        help="During --full-enrich, also run DDGS snippet backfill for missing fields",
    )
    parser.add_argument(
        "--with-audit",
        action="store_true",
        help="After each market pipeline (and --full-enrich if set): web search + LLM audit, "
        "remove bad rows from CSV/XLSX/JSON, write audit_report.md",
    )
    parser.add_argument(
        "--audit-concurrent",
        type=int,
        default=8,
        help="Parallel company audits per market during --with-audit (default: 8)",
    )
    parser.add_argument(
        "--gap-fill",
        action="store_true",
        help="After pipeline (and --full-enrich / --with-audit if set): multi-source gap-fill for "
        "Founded, HQ, Cities/Regions, Employees, Turnover, LinkedIn",
    )
    parser.add_argument(
        "--gap-fill-concurrent",
        type=int,
        default=4,
        help="Parallel company gap-fills per market during --gap-fill (default: 4)",
    )
    parser.add_argument(
        "--gap-fill-sources",
        default=None,
        help="Comma sources for --gap-fill (any market): google_ai,backfill,wikipedia,wikidata,"
        "owler,linkedin,crunchbase,sec,locations,contact_page,leadership,openai_residual "
        "(default: google_ai→backfill→…→openai_residual). Works for 1…N markets in a batch.",
    )
    parser.add_argument(
        "--push-neon",
        action="store_true",
        help="After each market finishes (enrich/audit/gap-fill), upsert companies into Neon Postgres.",
    )
    parser.add_argument(
        "--no-apollo",
        action="store_true",
        help="In post-enrich (quality/full): skip Apollo contact-name fill. "
        "Keeps LinkedIn MCP firmographic verify (Founded/HQ) if enabled by mode.",
    )
    parser.add_argument(
        "--no-name-resolve",
        action="store_true",
        help="In post-enrich (quality/full): skip contact name-resolve waterfall. "
        "Keeps LinkedIn MCP firmographic verify (Founded/HQ) if enabled by mode.",
    )
    parser.add_argument(
        "--resume-checkpoint",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Reuse Phase 1 plan + Phase 2 company list from disk when present "
        "(skip rediscovery; continue at enrich/classify). Default: on. "
        "Use --no-resume-checkpoint to force a fresh Phase 1/2.",
    )
    args = parser.parse_args()

    from vendor_intel.pipeline.pipeline_modes import mode_describe, resolve_batch_options, apply_mode_settings

    fmcg_xlsx_override = False if args.no_fmcg_xlsx else None
    batch_opts = resolve_batch_options(
        args.mode,
        max_concurrent=args.max_concurrent,
        cap=args.cap,
        profile=args.profile,
        fmcg_xlsx=fmcg_xlsx_override,
        full_enrich=args.full_enrich,
        with_audit=args.with_audit,
        gap_fill=args.gap_fill,
        push_neon=args.push_neon,
        search_backfill=args.search_backfill,
    )

    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.utils.output_filter import install_stderr_filter
    from vendor_intel.utils.safe_stdio import install_safe_stdio

    install_safe_stdio()
    install_stderr_filter()
    apply_env_overrides()

    from vendor_intel.batch import db

    if args.input:
        from vendor_intel.batch.registry_loader import populate_registry

        reg_profile = batch_opts["profile"]
        reg_cap = batch_opts["cap"]
        summary = populate_registry(args.input, args.batch_id, profile=reg_profile, cap=reg_cap)
        print(
            f"  [batch] Loaded {summary['total']} markets from {args.input} "
            f"({summary['inserted']} new, {summary['updated']} already in registry, "
            f"{summary['skipped_duplicate_in_file']} duplicate lines skipped)"
        )
    else:
        if not db.list_markets(args.batch_id):
            print(f"  ERROR: batch '{args.batch_id}' has no markets loaded — pass --input to load one.")
            sys.exit(1)

    # Crashed/killed workers leave rows as 'running' (and sometimes 'failed').
    # Reap BEFORE the pending check, or resume prints "Nothing pending" forever.
    reaped = db.reap_stale_running(args.batch_id, older_than_seconds=0)
    if reaped:
        print(f"  [batch] Reaped {reaped} stuck running market(s) → pending")
    requeued = db.requeue_failed_markets(args.batch_id)
    if requeued:
        print(f"  [batch] Requeued {requeued} failed market(s) → pending")

    counts = db.batch_status_counts(args.batch_id)
    pending = counts.get("pending", 0)
    print(f"  [batch] Registry status: {counts or '(empty)'}")
    if pending == 0:
        print("  [batch] Nothing pending — every market in this batch is already done/failed.")
        return

    mode_line = f"Mode: {batch_opts['mode']} — {mode_describe(batch_opts['mode'])}" if batch_opts.get("mode") else ""
    if mode_line:
        print(f"  [batch] {mode_line}")
    print(
        f"  [batch] Profile: {batch_opts['profile']} | Cap: {batch_opts['cap']} | "
        f"Concurrency: {batch_opts['max_concurrent']} | Output: output/<market_slug>/"
        + (" | FMCG xlsx: off" if not batch_opts["fmcg_xlsx"] else "")
        + (f" | post-market: on" if batch_opts.get("post_market") else "")
        + (f" | full-enrich: on (enrich-concurrent={args.enrich_concurrent})" if batch_opts["full_enrich"] else "")
        + (f" | with-audit: on (audit-concurrent={args.audit_concurrent})" if batch_opts["with_audit"] else "")
        + (f" | gap-fill: on (gap-concurrent={args.gap_fill_concurrent})" if batch_opts["gap_fill"] else "")
        + (" | push-neon: on" if batch_opts["push_neon"] else "")
        + (" | resume-checkpoint: on" if args.resume_checkpoint else " | resume-checkpoint: OFF")
    )
    if not args.yes:
        try:
            confirm = input(
                f"  Run {pending} pending market(s)? "
                "[Enter to start, any other key to cancel]: "
            ).strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  Cancelled.")
            return
        if confirm:
            print("  Cancelled.")
            return

    from vendor_intel.config import Settings

    profile = batch_opts["profile"]
    cap = batch_opts["cap"]
    base_settings = Settings.load().model_copy(
        update={
            "use_mock_data": False,
            "mock_mode": False,
            "pipeline_profile": profile,
            "pipeline_recall_mode": profile == "recall",
            "pipeline_use_ssc": True,
            "pipeline_min_export_confidence": 0.50,
        }
    )
    if batch_opts.get("mode"):
        base_settings = apply_mode_settings(base_settings, batch_opts["mode"], cap_override=cap)
    else:
        from vendor_intel.pipeline.cap_profiles import apply_cap

        base_settings = apply_cap(base_settings, cap)

    post_market = batch_opts.get("post_market")
    if post_market and batch_opts["push_neon"]:
        post_market = {**post_market, "push_neon": True}
    elif post_market and not batch_opts["push_neon"]:
        post_market = {**post_market, "push_neon": False}
    if post_market and (args.no_apollo or args.no_name_resolve):
        post_market = dict(post_market)
        if args.no_apollo:
            post_market["apollo_names"] = False
        if args.no_name_resolve:
            post_market["resolve_partials"] = False
        print(
            "  [batch] Post-enrich overrides:"
            + (" apollo=OFF" if args.no_apollo else "")
            + (" name_resolve=OFF" if args.no_name_resolve else "")
            + " | linkedin_firmographic_verify="
            + ("ON" if post_market.get("linkedin_firmographic_verify", True) else "OFF"),
            flush=True,
        )
        print(
            "  [batch] Post-enrich steps: gap-fill → LinkedIn Founded/HQ"
            + (" → Apollo" if post_market.get("apollo_names") else "")
            + (" → name_resolve" if post_market.get("resolve_partials") else "")
            + " → brand → market_audit → Excel"
            + (" → Neon" if post_market.get("push_neon") else ""),
            flush=True,
        )

    from vendor_intel.batch.orchestrator import run_batch

    try:
        final_counts = run_batch(
            args.batch_id,
            base_settings,
            max_concurrent_markets=batch_opts["max_concurrent"],
            fmcg_xlsx=batch_opts["fmcg_xlsx"],
            full_enrich=batch_opts["full_enrich"],
            enrich_concurrent=args.enrich_concurrent,
            search_backfill=batch_opts["search_backfill"],
            with_audit=batch_opts["with_audit"],
            audit_concurrent=args.audit_concurrent,
            gap_fill=batch_opts["gap_fill"],
            gap_fill_concurrent=args.gap_fill_concurrent,
            gap_fill_sources=args.gap_fill_sources,
            push_neon=batch_opts["push_neon"] and not post_market,
            post_market=post_market,
            resume_checkpoint=args.resume_checkpoint,
            limit=args.limit,
        )
    finally:
        try:
            from vendor_intel.clients.ddg_worker_pool import shutdown_ddg_pool

            shutdown_ddg_pool(wait=True)
        except Exception:
            pass

    print(f"\n  [batch] Stopped. Final status: {final_counts}")


if __name__ == "__main__":
    main()
