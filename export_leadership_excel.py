#!/usr/bin/env python3
"""Rebuild the leadership Excel workbook from the existing checkpoint.

Lets you re-export (e.g. after changing the formatting, or to take a
snapshot mid-run) without re-querying anything.

Usage:
    python export_leadership_excel.py
    python export_leadership_excel.py --output output/leadership_lookup/final.xlsx
    python export_leadership_excel.py --min-names 1     # skip fully empty rows
    python export_leadership_excel.py --sort-by-names   # most complete rows first
"""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.pipeline.leadership_lookup import ROLES, Checkpoint, write_excel

DEFAULT_CHECKPOINT = ROOT / "output" / "leadership_lookup" / "leadership_checkpoint.json"
DEFAULT_OUTPUT = ROOT / "output" / "leadership_lookup" / "leadership_contacts.xlsx"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--checkpoint", default=str(DEFAULT_CHECKPOINT))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument(
        "--min-names",
        type=int,
        default=0,
        help="Only export companies with at least this many names (default 0 = export everything)",
    )
    parser.add_argument(
        "--sort-by-names",
        action="store_true",
        help="Put the most complete rows first instead of input order",
    )
    parser.add_argument(
        "--names-only",
        action="store_true",
        help="Drop the mostly-empty EMAIL columns for a compact, on-screen sheet",
    )
    args = parser.parse_args()

    checkpoint_path = Path(args.checkpoint)
    if not checkpoint_path.exists():
        raise SystemExit(f"no checkpoint at {checkpoint_path}")

    rows = Checkpoint(str(checkpoint_path)).all_rows()

    def name_count(row) -> int:
        return sum(1 for name, _ in row.values.values() if name)

    total = len(rows)
    if args.min_names > 0:
        rows = [r for r in rows if name_count(r) >= args.min_names]
    if args.sort_by_names:
        rows.sort(key=lambda r: (-name_count(r), r.company.lower()))

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    write_excel(rows, str(output), names_only=args.names_only)

    with_names = sum(1 for r in rows if name_count(r))
    names = sum(name_count(r) for r in rows)
    emails = sum(1 for r in rows for _, email in r.values.values() if email)
    print(f"checkpoint : {total} companies")
    print(f"exported   : {len(rows)} rows -> {output}")
    print(f"  with >=1 name : {with_names}")
    print(f"  total names   : {names}")
    print(f"  total emails  : {emails}")
    for role in ROLES:
        filled = sum(1 for r in rows if (r.values.get(role) or ("", ""))[0])
        print(f"    {role:16} {filled}")


if __name__ == "__main__":
    main()
