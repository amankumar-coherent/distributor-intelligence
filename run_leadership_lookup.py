#!/usr/bin/env python3
"""Leadership contact lookup pipeline (MD/CEO/Country Head/CTO/AI Head/
Marketing Head/Founder name + email) for a CSV list of companies.

Standalone from every other pipeline in this repo — reads a "company"
column from an input CSV, asks the Google AI Overview scraper one question
per company, and writes a new Excel workbook with the results.

Usage:
    python run_leadership_lookup.py --input "C:\\Users\\aman.kumar\\Downloads\\TO RESEARCH PARTNER - Sheet1.csv"
    python run_leadership_lookup.py --input companies.csv --limit 10
    python run_leadership_lookup.py --input companies.csv --start 50 --limit 100
    python run_leadership_lookup.py --input companies.csv --output output/leadership_lookup/results.xlsx

Requires GOOGLE_AI_SCRAPER_ENABLED=true in .env and the google-ai-scraper
backend + browser extension already running (see
src/vendor_intel/clients/google_ai_scraper.py header for the two commands).
"""
import argparse
import asyncio
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import pandas as pd

from vendor_intel.placeholders.load_keys import apply_env_overrides
from vendor_intel.pipeline.leadership_lookup import Checkpoint, run_lookup, run_lookup_concurrent, write_excel


def load_companies(input_path: str) -> list[str]:
    df = pd.read_csv(input_path)
    col = None
    for candidate in df.columns:
        if candidate.strip().lower() == "company":
            col = candidate
            break
    if col is None:
        col = df.columns[0]
    names = [str(v).strip() for v in df[col].tolist() if str(v).strip() and str(v).strip().lower() != "nan"]
    # de-dupe, keep first-seen order
    seen = set()
    unique = []
    for name in names:
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(name)
    return unique


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, help="CSV with a 'company' column")
    parser.add_argument("--output", default=None, help="Output .xlsx path (default: output/leadership_lookup/leadership_contacts.xlsx)")
    parser.add_argument("--start", type=int, default=0, help="Skip the first N companies")
    parser.add_argument("--limit", type=int, default=None, help="Only process this many companies")
    parser.add_argument("--delay", type=float, default=1.0, help="Seconds to sleep between queries (sequential mode only)")
    parser.add_argument("--concurrency", type=int, default=5, help="How many companies to look up in parallel tabs (default: 5, matches the scraper's designed ~10-tab concurrency). Pass 1 to run sequentially instead.")
    parser.add_argument("--checkpoint", default=None, help="Checkpoint JSON path to resume from (default: output/leadership_lookup/leadership_checkpoint.json). Completed companies are skipped on re-run.")
    parser.add_argument("--no-checkpoint", action="store_true", help="Disable checkpointing entirely (always re-query every company)")
    args = parser.parse_args()

    apply_env_overrides()

    companies = load_companies(args.input)
    companies = companies[args.start :]
    if args.limit is not None:
        companies = companies[: args.limit]

    output_path = args.output or str(ROOT / "output" / "leadership_lookup" / "leadership_contacts.xlsx")
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    checkpoint = None
    if not args.no_checkpoint:
        checkpoint_path = args.checkpoint or str(Path(output_path).parent / "leadership_checkpoint.json")
        checkpoint = Checkpoint(checkpoint_path)
        already_done = sum(1 for c in companies if checkpoint.is_done(c))
        if already_done:
            print(f"Resuming from checkpoint: {already_done}/{len(companies)} companies already done ({checkpoint_path})")

    print(f"Looking up {len(companies)} companies (concurrency={args.concurrency})...")
    if args.concurrency > 1:
        rows = asyncio.run(
            run_lookup_concurrent(companies, concurrency=args.concurrency, checkpoint=checkpoint)
        )
    else:
        rows = run_lookup(companies, delay_seconds=args.delay, checkpoint=checkpoint)

    write_excel(rows, output_path)
    print(f"\nExcel written to: {output_path}")


if __name__ == "__main__":
    main()
