#!/usr/bin/env python3
"""Finish a market in one command: screen what's left, fill every column, export.

Replaces running these three by hand, in this order:

    run_cursor_expand.py --backend kompass      # screen remaining candidates
    refill_accepted.py --all-columns            # fill columns via Google AI
    refill_accepted.py --deep-only              # crawl company websites
    refill_accepted.py --export-only            # write the Excel

Order matters and is enforced here. The website crawl can only run on rows
that HAVE a website, and the website itself is discovered during the column
fill — so running the crawl first wastes it on blank rows.

Every stage is resumable and writes to the same checkpoint, so stopping this
with Ctrl+C and re-running it continues rather than restarting. The Excel is
written after each stage, not only at the end, so a stopped run still leaves
a usable file.

Usage:
    .\\.venv\\Scripts\\python.exe finish_market.py --query "Wearable Medical Devices"
    .\\.venv\\Scripts\\python.exe finish_market.py --query "..." --limit 100   # measure first
    .\\.venv\\Scripts\\python.exe finish_market.py --query "..." --skip-screen
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

PY = sys.executable


def _run(label: str, argv: list[str]) -> int:
    print(f"\n{'=' * 70}\n{label}\n{'=' * 70}", flush=True)
    started = time.time()
    # Inherit stdio so the child's progress is visible live rather than
    # buffered until it finishes — these stages run for hours.
    proc = subprocess.run(argv, cwd=str(ROOT))
    mins = (time.time() - started) / 60
    print(f"\n-- {label}: exit {proc.returncode} after {mins:.1f} min", flush=True)
    return proc.returncode


def _progress(query: str, country: str) -> tuple[int, int, int]:
    """(screened, total_candidates, accepted) straight from the checkpoint."""
    from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint
    from vendor_intel.pipeline.web_expand import resolve_final_path

    final_path = resolve_final_path(query, country, None, None)
    ckpt = ExpandCheckpoint.load(ExpandCheckpoint.path_for(final_path.parent, "kompass"))
    if ckpt is None:
        return 0, 0, 0
    return (
        len(ckpt.data("kompass_processed_urls", []) or []),
        len(ckpt.data("kompass_candidates", []) or []),
        len(ckpt.data("kompass_accepted", []) or []),
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", required=True)
    ap.add_argument("--country", default="global")
    ap.add_argument(
        "--limit",
        type=int,
        default=None,
        help="cap the refill stages to the first N companies — use this to "
             "measure the gain per row before committing to a long run",
    )
    ap.add_argument("--skip-screen", action="store_true",
                    help="candidates are already fully screened; go straight to filling")
    ap.add_argument("--skip-deep", action="store_true",
                    help="skip the website crawl (Google-only fill)")
    args = ap.parse_args()

    import run_cursor_expand as rce

    rce._load_env()

    before = _progress(args.query, args.country)
    print(f"start: {before[0]}/{before[1]} screened, {before[2]} accepted")

    # 1. Screen whatever candidates are left. Resumes from the checkpoint, so
    #    already-screened companies are skipped.
    if not args.skip_screen:
        screened, total, _ = _progress(args.query, args.country)
        if total and screened >= total:
            print(f"\nall {total} candidates already screened — skipping stage 1")
        else:
            rc = _run(
                f"STAGE 1/3  screen remaining {max(0, total - screened)} candidates",
                [PY, "run_cursor_expand.py", "--query", args.query, "--backend", "kompass"],
            )
            if rc != 0:
                print("\nscreening did not finish cleanly — stopping before the "
                      "fill stages so its checkpoint is not overwritten.")
                return rc

    common = [PY, "refill_accepted.py", "--query", args.query, "--country", args.country]
    limit = ["--limit", str(args.limit)] if args.limit else []

    # 2. Fill every column via Google AI Overview. --all-columns rather than
    #    only-blanks: rows accepted before the newer extractors existed have
    #    empty cells that were never recorded as gaps, so a blanks-only pass
    #    would skip exactly the rows that need it most.
    rc = _run("STAGE 2/3  fill all columns (Google AI Overview)",
              common + ["--all-columns"] + limit)
    if rc != 0:
        print("\ncolumn fill failed — the Excel below reflects what was saved.")

    # 3. Crawl each company's own site for the contact details Google does
    #    not publish for small private firms. Runs AFTER stage 2 because it
    #    needs the website stage 2 discovers.
    if not args.skip_deep:
        rc = _run("STAGE 3/3  crawl company websites (contact details)",
                  common + ["--deep-only"] + limit)
        if rc != 0:
            print("\nwebsite crawl failed — the Excel below reflects what was saved.")

    _run("EXPORT", common + ["--export-only"])

    after = _progress(args.query, args.country)
    print(f"\n{'=' * 70}")
    print(f"screened : {before[0]} -> {after[0]} of {after[1]}")
    print(f"accepted : {before[2]} -> {after[2]} distributors")
    print(f"{'=' * 70}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
