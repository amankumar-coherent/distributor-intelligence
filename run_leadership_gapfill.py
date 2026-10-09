#!/usr/bin/env python3
"""Second pass over the leadership checkpoint: for each company, ask Google
only about the roles that are still blank.

Why a narrower query: re-asking "md,ceo,country head,cto,ai head,marketing
head,founder" when six of the seven are already known spends the answer on
restating what we have, and the AI Overview often replies with a generic
company summary instead of naming the one person wanted. Asking "give me
full name and email of country head of X" is a direct question and answers
far more often.

Existing names are never overwritten — a gap-fill run can only add.

Usage:
    python run_leadership_gapfill.py                    # every company with a blank role
    python run_leadership_gapfill.py --only-empty       # just the fully-empty rows
    python run_leadership_gapfill.py --max-missing 3    # skip near-empty rows
    python run_leadership_gapfill.py --limit 50         # try a small batch first
    python run_leadership_gapfill.py --concurrency 3
"""
import argparse
import asyncio
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides
from vendor_intel.pipeline.leadership_lookup import (
    ROLES,
    Checkpoint,
    build_gap_query,
    missing_roles,
    run_gap_fill,
    write_excel,
)

DEFAULT_CHECKPOINT = ROOT / "output" / "leadership_lookup" / "leadership_checkpoint.json"
DEFAULT_OUTPUT = ROOT / "output" / "leadership_lookup" / "leadership_contacts.xlsx"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--checkpoint", default=str(DEFAULT_CHECKPOINT))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--concurrency", type=int, default=3)
    parser.add_argument("--limit", type=int, default=None, help="Only process this many companies")
    parser.add_argument(
        "--only-empty",
        action="store_true",
        help="Only companies with zero names so far (the truncated ones)",
    )
    parser.add_argument(
        "--max-missing",
        type=int,
        default=None,
        help="Skip companies missing more than this many roles",
    )
    parser.add_argument(
        "--max-names",
        type=int,
        default=None,
        help=(
            "Only companies with FEWER than this many names already found. "
            "e.g. --max-names 3 works the sparse rows (0, 1 or 2 names) and "
            "leaves anything with 3+ alone."
        ),
    )
    parser.add_argument(
        "--max-attempts",
        type=int,
        default=None,
        help=(
            "Only companies the gap-fill has queried fewer than this many "
            "times. Use --max-attempts 1 to resume a stopped run and skip "
            "everything already tried (most remaining gaps are roles the "
            "company genuinely does not have, so re-asking rarely helps)."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the queries that would be sent, without querying",
    )
    args = parser.parse_args()

    apply_env_overrides()

    checkpoint_path = Path(args.checkpoint)
    if not checkpoint_path.exists():
        raise SystemExit(f"no checkpoint at {checkpoint_path}")
    checkpoint = Checkpoint(str(checkpoint_path))
    rows = checkpoint.all_rows()

    def n_names(row) -> int:
        return sum(1 for name, _ in row.values.values() if name)

    candidates = [r for r in rows if missing_roles(r)]
    if args.only_empty:
        candidates = [r for r in candidates if n_names(r) == 0]
    if args.max_missing is not None:
        candidates = [r for r in candidates if len(missing_roles(r)) <= args.max_missing]
    if args.max_names is not None:
        candidates = [r for r in candidates if n_names(r) < args.max_names]
    if args.max_attempts is not None:
        candidates = [r for r in candidates if r.gapfill_attempts < args.max_attempts]
    # Fewest gaps first: those are the cheapest, most likely wins.
    candidates.sort(key=lambda r: len(missing_roles(r)))
    if args.limit is not None:
        candidates = candidates[: args.limit]

    print(f"checkpoint : {len(rows)} companies")
    print(f"selected   : {len(candidates)} for gap-fill")
    print(f"names now  : {sum(n_names(r) for r in rows)}")
    gapped = [r for r in rows if missing_roles(r)]
    never = sum(1 for r in gapped if r.gapfill_attempts == 0)
    print(f"with gaps  : {len(gapped)}  ({never} never gap-filled, {len(gapped) - never} already tried)")

    if args.dry_run:
        print("\n--- queries that would be sent (first 20) ---")
        for row in candidates[:20]:
            miss = missing_roles(row)
            print(f"  [{len(miss)} missing] {build_gap_query(row.company, miss)}")
        print(f"\nDry run — {len(candidates)} queries total. Drop --dry-run to execute.")
        return

    if not candidates:
        print("nothing to do")
        return

    before = sum(n_names(r) for r in rows)
    updated = asyncio.run(
        run_gap_fill(candidates, concurrency=args.concurrency, checkpoint=checkpoint)
    )

    gained = sum(n_names(r) for r in updated) - sum(n_names(r) for r in candidates)
    print(f"\nnew names found: {gained}")

    # Re-read so the export includes both the gap-filled rows and the
    # untouched ones.
    final_rows = Checkpoint(str(checkpoint_path)).all_rows()
    after = sum(n_names(r) for r in final_rows)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    write_excel(final_rows, str(output))

    filled = sum(1 for r in final_rows if n_names(r))
    print(f"names total: {before} -> {after}")
    print(f"companies with >=1 name: {filled} of {len(final_rows)}")
    print("per-role totals:")
    for role in ROLES:
        print(f"  {role:16} {sum(1 for r in final_rows if (r.values.get(role) or ('',''))[0])}")
    print(f"\nExcel written to: {output}")


if __name__ == "__main__":
    main()
