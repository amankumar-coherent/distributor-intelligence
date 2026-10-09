#!/usr/bin/env python3
"""Backfill the gapfill_attempts counter for companies processed before that
counter existed.

The gap-fill pass only started recording attempts recently, so a run that
was interrupted before then left no marker of how far it got. Comparing the
current checkpoint against the pre-gap-fill snapshot recovers most of it:
any company whose name count increased was definitely queried.

Companies whose counts are unchanged are ambiguous — either never queried,
or queried and nothing found. Those are left at 0 so they get retried,
which is the safe direction: a wasted query costs time, a wrongly-skipped
company loses data permanently.

Usage:
    python mark_gapfill_attempted.py              # dry run
    python mark_gapfill_attempted.py --apply
"""
import argparse
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.pipeline.leadership_lookup import ROLES

CHECKPOINT = ROOT / "output" / "leadership_lookup" / "leadership_checkpoint.json"
BASELINE = ROOT / "output" / "leadership_lookup" / "leadership_checkpoint.pre_gapfill.bak"


def names_in(entry: dict) -> int:
    values = entry.get("values") or {}
    return sum(1 for role in ROLES if (values.get(role) or ["", ""])[0].strip())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--checkpoint", default=str(CHECKPOINT))
    parser.add_argument("--baseline", default=str(BASELINE))
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    current = json.loads(Path(args.checkpoint).read_text(encoding="utf-8"))
    baseline = json.loads(Path(args.baseline).read_text(encoding="utf-8"))

    marked = 0
    already = 0
    for key, entry in current.items():
        if entry.get("gapfill_attempts"):
            already += 1
            continue
        before = baseline.get(key)
        if before is not None and names_in(entry) > names_in(before):
            entry["gapfill_attempts"] = 1
            marked += 1

    print(f"checkpoint : {len(current)} companies")
    print(f"already marked        : {already}")
    print(f"newly marked attempted: {marked}   (gained names since the baseline)")
    print(f"left at 0 (will retry): {len(current) - already - marked}")

    if not args.apply:
        print("\nDry run. Re-run with --apply to write.")
        return

    backup = Path(args.checkpoint).with_suffix(".json.premark.bak")
    shutil.copy2(args.checkpoint, backup)
    tmp = Path(args.checkpoint).with_suffix(".tmp")
    tmp.write_text(json.dumps(current, indent=2), encoding="utf-8")
    tmp.replace(args.checkpoint)
    print(f"\nWritten. Backup at {backup}")


if __name__ == "__main__":
    main()
