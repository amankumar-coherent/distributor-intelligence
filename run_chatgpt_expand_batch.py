#!/usr/bin/env python3
"""Multi-market ChatGPT expand (steps 0–7) for 1…10k markets.

Same pipeline as ``run_cursor_expand.py --backend chatgpt``, but driven from a
market list + SQLite registry (skip-done, resume, concurrent markets).

Examples (from D:\\vendor-intel):

  # Load ~10k markets then expand (1 market at a time — safest for OpenAI limits)
  .\\.venv\\Scripts\\python.exe run_chatgpt_expand_batch.py ^
      --input queries/neon_markets_all.csv --batch-id expand-10k --yes

  # Resume after Ctrl+C (done markets skipped automatically)
  .\\.venv\\Scripts\\python.exe run_chatgpt_expand_batch.py --batch-id expand-10k --yes

  # Smoke: 2 markets, 2 at once
  .\\.venv\\Scripts\\python.exe run_chatgpt_expand_batch.py ^
      --input queries/neon_markets_all.csv --batch-id expand-smoke ^
      --limit 2 --max-concurrent 1 --yes
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))


def _load_env() -> None:
    try:
        from dotenv import dotenv_values, load_dotenv
    except ImportError:
        return
    load_dotenv(ROOT / ".env")
    cmi_env = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env")
    if cmi_env.exists():
        # Prefer Cmi LLM keys (DeepSeek) over stale D:\ OpenAI keys
        vals = dotenv_values(cmi_env)
        for k, v in (vals or {}).items():
            if not v:
                continue
            if k.startswith(
                ("OPENAI_", "LLM_", "COMPILER_", "CLASSIFIER_", "MARKET_MAP_")
            ):
                os.environ[k] = v
        load_dotenv(cmi_env, override=False)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="ChatGPT expand steps 0–7 for every market in a batch (10k-ready)"
    )
    parser.add_argument(
        "--input",
        "-i",
        default=None,
        help="Market list (.csv with query|name, or .txt). Omit to resume --batch-id.",
    )
    parser.add_argument("--batch-id", "-b", required=True, help="Batch id (reuse to resume)")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip confirmation")
    parser.add_argument("--limit", type=int, default=None, help="Process at most N markets this run")
    parser.add_argument(
        "--max-concurrent",
        "-c",
        type=int,
        default=1,
        help="Markets in parallel (default 1; use 1–2 for large OpenAI batches)",
    )
    parser.add_argument("--target", "-t", type=int, default=320)
    parser.add_argument("--max-fill", type=int, default=320)
    parser.add_argument("--max-queries", type=int, default=120)
    parser.add_argument(
        "--fill-backend",
        choices=("search", "openai"),
        default="search",
        help="search=Google AI→DDGS/SearXNG→…→OpenAI residual (default)",
    )
    parser.add_argument("--ddgs-harvest", action="store_true", help="Enable optional Step 3")
    parser.add_argument("--seeds-only", action="store_true")
    parser.add_argument("--skip-recall", action="store_true")
    parser.add_argument("--no-openai-discover", action="store_true")
    parser.add_argument("--no-web-fill", action="store_true")
    parser.add_argument("--no-apollo", action="store_true")
    parser.add_argument("--no-linkedin-enrich", action="store_true")
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Ignore done checkpoints and redo markets (dangerous on 10k)",
    )
    parser.add_argument(
        "--profile",
        default="quality",
        help="Stored on registry rows only (expand ignores classic pipeline profile)",
    )
    parser.add_argument("--cap", default="maximum", help="Stored on registry rows only")
    args = parser.parse_args()

    _load_env()
    os.chdir(ROOT)

    try:
        from vendor_intel.utils.safe_stdio import install_safe_stdio

        install_safe_stdio()
    except Exception:
        pass

    from vendor_intel.batch import db
    from vendor_intel.batch.expand_orchestrator import run_expand_batch
    from vendor_intel.batch.registry_loader import populate_registry
    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()

    if args.input:
        summary = populate_registry(
            args.input, args.batch_id, profile=args.profile, cap=args.cap
        )
        print(
            f"  [expand-batch] Loaded {summary['total']} markets from {args.input} "
            f"({summary['inserted']} new, {summary['updated']} already in registry, "
            f"{summary['skipped_duplicate_in_file']} duplicate lines skipped)",
            flush=True,
        )
    else:
        if not db.list_markets(args.batch_id):
            print(
                f"  ERROR: batch '{args.batch_id}' has no markets — pass --input to load one.",
                file=sys.stderr,
            )
            return 1

    reaped = db.reap_stale_running(args.batch_id, older_than_seconds=0)
    if reaped:
        print(f"  [expand-batch] Reaped {reaped} stuck running market(s) → pending", flush=True)
    requeued = db.requeue_failed_markets(args.batch_id)
    if requeued:
        print(f"  [expand-batch] Requeued {requeued} failed market(s) → pending", flush=True)

    counts = db.batch_status_counts(args.batch_id)
    pending = counts.get("pending", 0)
    print(f"  [expand-batch] Registry status: {counts or '(empty)'}", flush=True)
    if pending == 0:
        print("  [expand-batch] Nothing pending — every market is already done/failed.", flush=True)
        return 0

    print(
        "  [expand-batch] Per market: "
        "0 seeds+Excel → 1 recall → 2 OpenAI discover → 3 DDGS(opt) → "
        "4 verify → 5 search-stack+OpenAI residual → 6 QA → 7 Excel",
        flush=True,
    )
    print(
        f"  [expand-batch] Will run {pending} pending | concurrent={args.max_concurrent} | "
        f"target={args.target} | fill={args.fill_backend}",
        flush=True,
    )
    if not args.yes:
        try:
            confirm = input(
                f"  Run {pending} pending expand market(s)? "
                "[Enter to start, any other key to cancel]: "
            ).strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  Cancelled.")
            return 0
        if confirm:
            print("  Cancelled.")
            return 0

    final = run_expand_batch(
        args.batch_id,
        max_concurrent_markets=args.max_concurrent,
        limit=args.limit,
        target=args.target,
        max_fill=args.max_fill,
        max_queries=args.max_queries,
        fill_backend=args.fill_backend,
        ddgs_harvest=args.ddgs_harvest,
        seeds_only=args.seeds_only,
        skip_recall=args.skip_recall,
        use_apollo=not args.no_apollo,
        use_linkedin_enrich=not args.no_linkedin_enrich,
        resume=not args.fresh,
        force=args.fresh,
        openai_discover=not args.no_openai_discover,
        web_fill=not args.no_web_fill,
    )
    print(f"  [expand-batch] Finished. Status: {final}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
