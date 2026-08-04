#!/usr/bin/env python3
"""
status_batch.py — lightweight progress view for a run_batch.py batch.

Reads data/batch.db directly — no dashboard, no separate service. Safe to run
in a second terminal alongside an in-progress `run_batch.py`: SQLite's WAL
mode lets a read-only connection see committed writes from the writer process
without blocking it.

Usage:
  python status_batch.py --batch-id pilot-50
  python status_batch.py --batch-id pilot-50 --watch
  python status_batch.py --batch-id pilot-50 --watch --interval 10
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))


def _fmt_duration(seconds: float | None) -> str:
    if seconds is None:
        return "—"
    if seconds < 60:
        return f"{seconds:.0f}s"
    if seconds < 3600:
        return f"{seconds / 60:.1f}m"
    return f"{seconds / 3600:.1f}h"


def print_status(batch_id: str) -> dict[str, int]:
    from vendor_intel.batch import db

    counts = db.batch_status_counts(batch_id)
    summary = db.batch_progress_summary(batch_id)

    total = sum(counts.values())
    pending = counts.get("pending", 0)
    running = counts.get("running", 0)
    done = counts.get("done", 0)
    failed = counts.get("failed", 0)

    print(f"\n{'=' * 60}")
    print(f"  Batch: {batch_id}")
    print(f"{'=' * 60}")
    if total == 0:
        print("  (no markets loaded for this batch-id)")
        print()
        return counts
    print(f"  Total markets    : {total}")
    print(f"  Pending          : {pending}")
    print(f"  Running          : {running}")
    print(f"  Done             : {done}")
    print(f"  Failed           : {failed}")
    print(f"  Companies found  : {summary['companies_exported_total']}")

    avg_elapsed = summary["avg_elapsed_seconds"]
    if avg_elapsed is not None:
        remaining = pending + running
        eta_seconds = remaining * avg_elapsed / summary["concurrency"]
        print(f"  Avg time/market  : {_fmt_duration(avg_elapsed)}")
        print(f"  ETA (remaining)  : {_fmt_duration(eta_seconds)} at concurrency={summary['concurrency']}")
    else:
        print("  ETA              : — (no completed markets yet)")

    failures = db.list_recent_failures(batch_id, limit=5)
    if failures:
        print(f"\n  Recent failures (up to 5):")
        for f in failures:
            reason = (f["last_error"] or "")[:100]
            print(f"    [{f['last_error_kind'] or '?'}] {f['query']} ({f['country']}): {reason}")
    print()
    return counts


def main() -> None:
    parser = argparse.ArgumentParser(description="Vendor Intelligence — batch progress viewer")
    parser.add_argument("--batch-id", "-b", required=True)
    parser.add_argument("--watch", "-w", action="store_true", help="Keep polling until the batch finishes")
    parser.add_argument("--interval", type=float, default=5.0, help="Seconds between refreshes in --watch mode")
    args = parser.parse_args()

    if not args.watch:
        print_status(args.batch_id)
        return

    try:
        while True:
            counts = print_status(args.batch_id)
            if counts.get("pending", 0) == 0 and counts.get("running", 0) == 0:
                print("  [status] Batch finished (no pending/running markets left).")
                break
            time.sleep(args.interval)
    except KeyboardInterrupt:
        print("\n  Stopped watching (the batch process itself keeps running).")


if __name__ == "__main__":
    main()
