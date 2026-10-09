#!/usr/bin/env python3
"""Enforce the rule that Operational Presence is a SUPERSET of Continent/Geography.

The two columns have definitions you gave explicitly:

  * Continent / Geography  = countries where the company has a PHYSICAL
                             SITE (HQ, office, warehouse, plant)
  * Operational Presence   = every country it does business in, which is
                             WIDER

So a country with a physical site is necessarily a country it does
business in. Operational Presence listing FEWER countries is therefore
impossible, not merely unusual.

Found in the Dry Eyes verified file (2 of 373 rows):

    Zuellig Pharma Vietnam
      CG: APAC(Vietnam, Singapore, Philippines, Thailand, Malaysia, Indonesia)
      OP: Vietnam; Cambodia; Laos          <- 3 countries, but 6 sites

Fixed deterministically by unioning the two lists, which needs no model
call: every country already appears in one of the columns, so nothing is
invented. Countries only in Operational Presence stay there (the company
sells somewhere it has no site — entirely normal); countries only in
Continent/Geography are ADDED to Operational Presence.

Deliberately does NOT touch Continent/Geography. Removing a site claim
would be a factual assertion this script has no basis for, whereas adding
to Operational Presence follows from the definitions alone.

Usage:
    python fix_geography_superset.py --market "Global Dry Eyes Treatment Market - Distributors Only" --dry-run
    python fix_geography_superset.py --market "..." --export
    python fix_geography_superset.py --all --export
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
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

CG = "Continent / Geography"
OP = "Operational Presence"

# Region labels that wrap the country list, e.g. "APAC(India, Japan)".
_REGIONS = {
    "apac", "emea", "eu", "europe", "na", "north america", "latam",
    "south america", "mea", "middle east", "africa", "asia", "oceania",
    "anz", "americas", "asia pacific", "caribbean", "nordics", "benelux",
    "dach", "sea", "southeast asia", "south asia", "cis", "gcc", "global",
    "worldwide",
    # Short forms actually present in the live data. Missing these made
    # the first dry run "fix" 132 rows by appending "ME" and "AF" as if
    # they were countries -- worse than the problem being fixed.
    "me", "af", "as", "eu", "sa", "latin america", "middle east & africa",
    "mena", "apj", "emea & apac", "row", "rest of world", "n america",
    "s america", "central america", "east africa", "west africa",
    "north africa", "southern africa", "sub-saharan africa", "levant",
    "balkans", "baltics", "iberia", "adriatic", "andean", "mercosur",
}

# A token that is a bare acronym (2-4 capitals) is a region label, not a
# country: no country name is written that way in this data. Catching the
# shape as well as the list means a new abbreviation cannot silently
# become a fake country.
_ACRONYM = re.compile(r"^[A-Z]{2,4}$")


def countries_in(value: str) -> list[str]:
    """Country names from either column's format.

    Handles the region-wrapped form ("APAC(India, Japan); EU(Germany)")
    and the flat semicolon form ("India; Japan; Germany"), and drops bare
    region labels so "APAC" never becomes a country.
    """
    text = str(value or "")
    # Walk the string left to right so the result is in DOCUMENT order.
    #
    # Extracting all parenthesised groups first and then the remainder
    # reordered the output: "ME(Bahrain); Saudi Arabia; AF(Kenya)" came
    # back as Bahrain, Kenya, Saudi Arabia. Harmless for a membership
    # test, but it breaks the promise that an untouched value keeps its
    # original ordering, which is what lets a reviewer see at a glance
    # whether a row really changed.
    parts: list[str] = []
    pos = 0
    for match in re.finditer(r"\(([^)]*)\)", text):
        parts.extend(re.split(r"[;,]", text[pos : match.start()]))
        parts.extend(match.group(1).split(","))
        pos = match.end()
    parts.extend(re.split(r"[;,]", text[pos:]))

    out: list[str] = []
    seen: set[str] = set()
    for raw in parts:
        name = raw.strip(" \t.;,/-")
        if not name or name.lower() in _REGIONS or _ACRONYM.match(name):
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(name)
    return out


def widen(row: dict) -> str | None:
    """Operational Presence including every Continent/Geography country,
    or None when it is already a superset."""
    sites = countries_in(row.get(CG))
    ops = countries_in(row.get(OP))
    if not sites:
        return None

    have = {c.lower() for c in ops}
    missing = [c for c in sites if c.lower() not in have]
    if not missing:
        return None
    # Keep the existing order and append, so the value stays recognisable
    # rather than being reshuffled.
    return "; ".join(ops + missing)


def process(market: str, *, dry_run: bool) -> int:
    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        print(f"  no state file: {path.name}")
        return 0

    data = json.loads(path.read_text(encoding="utf-8"))
    companies = data.get("companies") or []

    fixes: list[tuple[dict, str]] = []
    for row in companies:
        new = widen(row)
        if new:
            fixes.append((row, new))

    print(f"{market}")
    print(f"  {len(companies)} companies, {len(fixes)} with OP narrower than CG")

    for row, new in fixes[:10]:
        print(f"    {str(row.get('Company'))[:32]:34}")
        print(f"      sites : {str(row.get(CG))[:74]}")
        print(f"      was   : {str(row.get(OP))[:74]}")
        print(f"      now   : {new[:74]}")

    if dry_run or not fixes:
        if dry_run and fixes:
            print("  DRY RUN — state unchanged")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = path.with_suffix(f".before-geo-superset-{stamp}.json")
    shutil.copy2(path, backup)
    for row, new in fixes:
        row[OP] = new
    data["companies"] = companies
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                    encoding="utf-8")
    print(f"  backup: {backup.name}")
    print(f"  fixed {len(fixes)} rows")
    return len(fixes)


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
