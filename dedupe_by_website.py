#!/usr/bin/env python3
"""Collapse rows that are the SAME company listed under name variants.

Measured on the finished workbooks:

    Dry Eyes : 373 rows -> 322 unique companies (51 redundant, 13%)
    Solar    : 387 rows -> 306 unique companies (81 redundant, 20%)

The pipeline already dedupes on a normalised company name, but these slip
past because the names genuinely differ:

    powernsun.com   PowernSun
                    Powernsun (Power n Sun Solar Systems)
                    Power & Sun
                    Power n Sun                     <- all one Dubai company

    solarhive.com.sg  SolarHive Pte Ltd
                      Solar Hive Pte. Ltd.
                      Solar Hive Pte Ltd
                      Solar Hive                    <- all one Singapore company

A shared website is much stronger evidence of identity than a name, since
a company controls exactly one primary domain.

BUT a shared domain is not proof on its own. A multinational group runs
one domain across genuinely distinct national distributors:

    krannich-solar.com  Krannich Solar GmbH & Co. KG   Weil der Stadt, Germany
                        Krannich Solar Turkey          Istanbul, Turkey
                        Krannich Solar d.o.o.          Zagreb, Croatia

Those are three real distributors with different geography and contacts,
and collapsing them would destroy information.

The rule used here: rows sharing a domain are merged only when they are
in the SAME COUNTRY. Different countries means different national
entities and they are kept apart. Within a merged group the most complete
row wins, and any field it leaves blank is filled from its siblings, so
merging never loses a value.

Usage:
    python dedupe_by_website.py --market "Global Solar Rooftop Market - Distributors Only" --dry-run
    python dedupe_by_website.py --all --export
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)


def domain_of(website: str) -> str:
    """Bare registrable-ish domain, for identity comparison."""
    text = str(website or "").strip().lower()
    text = re.sub(r"^https?://", "", text)
    text = re.sub(r"^www\.", "", text)
    return text.split("/")[0].split("?")[0].strip()


def country_of(row: dict) -> str:
    """Country from Headquarters, normalised for comparison.

    The country is what separates "the same company written four ways"
    from "a group's national subsidiaries", so it decides whether a
    shared domain means a merge.
    """
    hq = str(row.get("Headquarters") or "")
    part = hq.split(",")[-1] if "," in hq else hq
    # Reuse the pipeline's alias table so this tool and the collection-time
    # check agree. "UAE" and "United Arab Emirates" are one country; not
    # collapsing them lets a duplicate through under the other spelling.
    from vendor_intel.pipeline.market_research.state import normalise_country
    return normalise_country(part)


def completeness(row: dict) -> int:
    """How many fields carry a real value. Picks the survivor."""
    return sum(1 for v in row.values() if str(v or "").strip())


def merge_group(rows: list[dict]) -> dict:
    """Most complete row, with blanks filled from its siblings.

    Merging must never lose a value: if one variant recorded a contact
    person and another recorded a phone number, the survivor gets both.
    """
    best = max(rows, key=completeness)
    merged = dict(best)
    for other in rows:
        if other is best:
            continue
        for key, value in other.items():
            if not str(merged.get(key) or "").strip() and str(value or "").strip():
                merged[key] = value
    return merged


def plan(companies: list[dict]) -> tuple[list[dict], list[tuple[str, list[str]]]]:
    """Returns (deduped rows, [(domain, [names merged]) ...])."""
    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    passthrough: list[dict] = []

    for row in companies:
        dom = domain_of(row.get("Website"))
        if not dom:
            passthrough.append(row)
            continue
        # Same domain AND same country. A different country means a
        # distinct national entity, which is kept.
        groups[(dom, country_of(row))].append(row)

    out: list[dict] = []
    merges: list[tuple[str, list[str]]] = []
    for (dom, _country), rows in groups.items():
        if len(rows) == 1:
            out.append(rows[0])
            continue
        merges.append((dom, [str(r.get("Company") or "") for r in rows]))
        out.append(merge_group(rows))

    out.extend(passthrough)
    # Renumber so the "#" column stays sequential.
    for i, row in enumerate(out, 1):
        if "#" in row:
            row["#"] = i
    return out, merges


def process(market: str, *, dry_run: bool) -> int:
    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        print(f"  no state file: {path.name}")
        return 0

    data = json.loads(path.read_text(encoding="utf-8"))
    companies = data.get("companies") or []
    deduped, merges = plan(companies)
    removed = len(companies) - len(deduped)

    print(f"{market}")
    print(f"  {len(companies)} rows -> {len(deduped)} unique "
          f"({removed} redundant across {len(merges)} domains)")
    for dom, names in sorted(merges, key=lambda x: -len(x[1]))[:8]:
        print(f"    {dom}")
        for nm in names:
            print(f"        {nm[:52]}")

    if dry_run or not removed:
        if dry_run and removed:
            print("  DRY RUN — state unchanged")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = path.with_suffix(f".before-website-dedupe-{stamp}.json")
    shutil.copy2(path, backup)
    data["companies"] = deduped
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                    encoding="utf-8")
    print(f"  backup: {backup.name}")
    print(f"  removed {removed} redundant rows")
    return removed


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--market", action="append", default=[])
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--export", action="store_true")
    args = ap.parse_args()

    markets = list(args.market)
    if args.all:
        from run_all_markets import MARKETS  # noqa: PLC0415

        markets = list(MARKETS)
    if not markets:
        ap.error("pass --market or --all")

    changed = {}
    for market in markets:
        changed[market] = process(market, dry_run=args.dry_run)
        print()

    if args.export and not args.dry_run:
        from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

        for market, n in changed.items():
            if n:
                print(f"re-exported: {export_xlsx(market)}")


if __name__ == "__main__":
    main()
