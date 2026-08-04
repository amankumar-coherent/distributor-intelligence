#!/usr/bin/env python3
"""Web-search expand — YOU run this command; it does live web search + fill.

Examples (PowerShell, from D:\\vendor-intel):

  # Batch A — half of search queries, merge into FINAL, aim for 320
  .\\.venv\\Scripts\\python.exe run_web_expand.py --query "Global Firewall Market" --target 320 --batch a

  # Batch B — remaining queries
  .\\.venv\\Scripts\\python.exe run_web_expand.py --query "Global Firewall Market" --target 320 --batch b

  # All queries in one go (heavier)
  .\\.venv\\Scripts\\python.exe run_web_expand.py --query "Global Smartwatch Market" --target 320 --batch all

  # Deep enrich via full pipeline instead (LLM + crawl) — heavier
  .\\.venv\\Scripts\\python.exe run_web_expand.py --query "Global Green Chemical Market" --mode pipeline --cap maximum

What harvest mode does:
  1) Loads curated queries from queries/discovery/ + auto geo/role queries
  2) Live web search via FreeSearchRouter (ddgs / Bing / SearXNG)
  3) Extracts company name + domain, filters OEM/junk
  4) Merges into existing FINAL Excel (or creates one)
  5) Fills all 35 landscape columns (defaults + Not publicly disclosed for contacts)
  6) Writes wrap-formatted Excel + audit JSON
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))


def _run_pipeline(args: argparse.Namespace) -> int:
    """Delegate to existing run_query.py quality / maximum discovery."""
    cmd = [
        sys.executable,
        str(ROOT / "run_query.py"),
        "--query",
        args.query,
        "--country",
        args.country,
        "--mode",
        "quality",
        "--cap",
        args.cap,
        "--fmcg-xlsx",
        "--profile",
        "recall",
    ]
    if args.no_push_neon:
        cmd.append("--no-push-neon")
    print("  [web_expand] launching full pipeline:", " ".join(cmd), flush=True)
    return subprocess.call(cmd, cwd=str(ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Web-search own expand to N companies (you only run this command)"
    )
    parser.add_argument("--query", "-q", required=True, help='e.g. "Global Firewall Market"')
    parser.add_argument("--country", "-c", default="global")
    parser.add_argument("--target", "-t", type=int, default=320, help="Target company count")
    parser.add_argument(
        "--batch",
        "-b",
        choices=("a", "b", "all"),
        default="a",
        help="Split discovery queries into half-batches (safer on weak PCs)",
    )
    parser.add_argument(
        "--max-queries",
        type=int,
        default=80,
        help="Max web search queries in this run (default 80)",
    )
    parser.add_argument(
        "--mode",
        choices=("harvest", "pipeline"),
        default="harvest",
        help="harvest = web search + Excel merge (default); pipeline = full run_query",
    )
    parser.add_argument(
        "--cap",
        default="maximum",
        choices=("bulk", "focused", "standard", "broad", "maximum"),
        help="Only for --mode pipeline",
    )
    parser.add_argument(
        "--final",
        default=None,
        help="Path to FINAL xlsx to merge into (default: known run4 FINAL if present)",
    )
    parser.add_argument(
        "--out-dir",
        default=None,
        help="Folder for harvest JSON + audit (default: output/web_expand/<slug>)",
    )
    parser.add_argument(
        "--no-merge",
        action="store_true",
        help="Do not merge existing FINAL; write fresh list from harvest only",
    )
    parser.add_argument("--no-push-neon", action="store_true", help="Pipeline mode only")
    args = parser.parse_args()

    os.chdir(ROOT)

    if args.mode == "pipeline":
        return _run_pipeline(args)

    from vendor_intel.pipeline.web_expand import run_web_expand

    audit = asyncio.run(
        run_web_expand(
            args.query,
            country=args.country,
            target=args.target,
            batch=args.batch,
            max_queries=args.max_queries,
            merge_existing=not args.no_merge,
            final_path=args.final,
            out_dir=args.out_dir,
        )
    )
    print(json.dumps(audit, indent=2), flush=True)
    total = int(audit.get("total_after") or 0)
    target = int(args.target)
    if total < target:
        print(
            f"\n  NOTE: {total}/{target} companies. Run again with --batch b "
            f"or raise --max-queries.",
            flush=True,
        )
        return 0
    print(f"\n  OK: reached {total} >= {target}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
