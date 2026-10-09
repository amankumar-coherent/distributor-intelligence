#!/usr/bin/env python3
"""Drop zero-name companies from the leadership-lookup checkpoint so a
resumed run re-queries exactly those, keeping the ones that already worked.

Use this after changing the scraper/extension or the extraction prompt: rows
collected under the old behaviour are skipped as "already done" otherwise,
so fixes never get a chance to apply to them.

Usage:
    python prune_leadership_checkpoint.py                 # show what would be dropped
    python prune_leadership_checkpoint.py --apply         # actually drop them
    python prune_leadership_checkpoint.py --apply --min-names 3
"""
import argparse
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.pipeline.leadership_lookup import ROLES

DEFAULT_CHECKPOINT = ROOT / "output" / "leadership_lookup" / "leadership_checkpoint.json"


def names_in(entry: dict) -> int:
    values = entry.get("values") or {}
    return sum(1 for role in ROLES if (values.get(role) or ["", ""])[0].strip())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--checkpoint", default=str(DEFAULT_CHECKPOINT))
    parser.add_argument(
        "--min-names",
        type=int,
        default=1,
        help="Drop any company with fewer than this many names (default 1 = drop only fully empty rows)",
    )
    parser.add_argument("--apply", action="store_true", help="Write the change (otherwise dry-run)")
    args = parser.parse_args()

    path = Path(args.checkpoint)
    if not path.exists():
        print(f"No checkpoint at {path} — nothing to prune.")
        return

    data = json.loads(path.read_text(encoding="utf-8"))
    keep, drop = {}, []
    for key, entry in data.items():
        if names_in(entry) >= args.min_names:
            keep[key] = entry
        else:
            drop.append((entry.get("company", key), len(entry.get("raw_reply", ""))))

    print(f"Checkpoint: {len(data)} companies")
    print(f"  keep: {len(keep)}   drop (re-query): {len(drop)}\n")
    for company, raw_len in sorted(drop, key=lambda x: x[1]):
        print(f"  drop  rawlen={raw_len:5}  {company}")

    if not drop:
        print("\nNothing to drop.")
        return

    if not args.apply:
        print(f"\nDry run. Re-run with --apply to drop these {len(drop)} companies.")
        return

    backup = path.with_suffix(".json.bak")
    shutil.copy2(path, backup)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(keep, indent=2), encoding="utf-8")
    tmp.replace(path)
    print(f"\nDropped {len(drop)} companies. Backup written to {backup}")
    print("Re-run run_leadership_lookup.py to retry them.")


if __name__ == "__main__":
    main()
