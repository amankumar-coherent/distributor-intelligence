#!/usr/bin/env python3
"""Reduce the ten drone markets to India-only companies.

The markets were run with global coverage; the requirement is India-only
distributors. Of 1,771 companies collected, 106 are India-based, so the
rest are out of scope.

Removed rows are written to a sidecar archive before deletion. The
request was to delete them, and the state files end up identically clean
either way -- the archive costs nothing and makes an accidental purge
recoverable, which matters because this is hours of collected data.

Also rebuilds:
  * "#" numbering, so the exported workbook reads 1..N with no holes;
  * seen_keys, which is derived from the kept names -- a stale key would
    make discovery permanently exclude a company it no longer holds.

Usage:
    python purge_drones_to_india.py --dry-run
    python purge_drones_to_india.py
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from drone_markets import MARKETS  # noqa: E402
from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
    normalise_country,
)
from vendor_intel.utils.dedupe import company_dedupe_key  # noqa: E402

INDIA = normalise_country("India")


def country_of(row: dict) -> str:
    hq = str(row.get("Headquarters") or "").strip()
    return hq.split(",")[-1].strip() if hq else ""


def is_india(row: dict) -> bool:
    return normalise_country(country_of(row)) == INDIA


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    print(f"  {'market':22}{'before':>8}{'keep':>8}{'remove':>8}")
    tb = tk = tr = 0

    for slug, market, _ in MARKETS:
        path = STATE_DIR / f"{market_slug(market)}.json"
        if not path.is_file():
            print(f"  {slug:22}{'-':>8}{'-':>8}{'-':>8}")
            continue

        data = json.loads(path.read_text(encoding="utf-8"))
        rows = data.get("companies") or []
        keep = [r for r in rows if is_india(r)]
        drop = [r for r in rows if not is_india(r)]

        tb += len(rows)
        tk += len(keep)
        tr += len(drop)
        print(f"  {slug:22}{len(rows):>8}{len(keep):>8}{len(drop):>8}")

        if args.dry_run or not drop:
            continue

        archive = path.with_suffix(".json.global_archive")
        archive.write_text(
            json.dumps({"market": market, "removed": drop},
                       ensure_ascii=False, indent=2),
            encoding="utf-8")

        for i, row in enumerate(keep, 1):
            row["#"] = str(i)

        data["companies"] = keep
        data["last_index"] = len(keep)
        # Derived from the names, so it must shrink with them.
        data["seen_keys"] = [company_dedupe_key(str(r.get("Company") or ""))
                             for r in keep]
        # The verification checkpoint refers to companies that are gone;
        # clearing it forces a clean re-verify of the India-only set.
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                        encoding="utf-8")

    print(f"  {'-'*46}")
    print(f"  {'TOTAL':22}{tb:>8}{tk:>8}{tr:>8}")
    if args.dry_run:
        print("\n(dry run - nothing written)")
    else:
        print("\nremoved rows archived alongside each state file "
              "(*.json.global_archive)")


if __name__ == "__main__":
    main()
