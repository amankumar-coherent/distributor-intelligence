#!/usr/bin/env python3
"""
status_gap_fill.py — gap-fill progress for a batch.

Usage:
  python status_gap_fill.py --batch-id medical-batch-10
  python status_gap_fill.py --batch-id medical-batch-10 --watch
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))


def print_status(batch_id: str) -> dict:
    from vendor_intel.enrichment.gap_fill.batch import gap_fill_status

    s = gap_fill_status(batch_id)
    print(f"\n{'=' * 60}")
    print(f"  Gap-fill status: {batch_id}")
    print(f"{'=' * 60}")
    print(f"  Done markets (registry) : {s['done_markets']}")
    print(f"  Gap-fill completed    : {s['gap_fill_completed']}")
    print(f"  Gap-fill pending      : {s['gap_fill_pending']}")
    if s.get("no_json"):
        print(f"  No JSON export        : {s['no_json']}")

    markets = s.get("markets") or []
    if markets:
        print(f"\n  {'Market':<42} {'Status':<10} {'Updated':>8} {'Still gaps':>11}")
        print(f"  {'-' * 75}")
        for m in markets:
            q = (m.get("query") or m.get("slug") or "")[:42]
            st = m.get("status", "?")
            upd = m.get("updated", "")
            miss = m.get("still_missing", "")
            print(f"  {q:<42} {st:<10} {str(upd):>8} {str(miss):>11}")
    print()
    return s


def main() -> None:
    parser = argparse.ArgumentParser(description="Gap-fill batch progress")
    parser.add_argument("--batch-id", "-b", required=True)
    parser.add_argument("--watch", action="store_true")
    parser.add_argument("--interval", type=int, default=30)
    args = parser.parse_args()

    if args.watch:
        try:
            while True:
                print_status(args.batch_id)
                time.sleep(max(5, args.interval))
        except KeyboardInterrupt:
            print("\n  Stopped watching.")
    else:
        print_status(args.batch_id)


if __name__ == "__main__":
    main()
