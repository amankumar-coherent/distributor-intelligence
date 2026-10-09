#!/usr/bin/env python3
"""Country-wise sweep for the AP Automation channel-partner market.

One unrestricted global query returns the same few dozen household names
however many times you ask: the model answers "the biggest AP automation
partners" and stops. Restricting each query to ONE country makes it
answer "the biggest partners IN GERMANY", which surfaces the regional
resellers and ERP integrators a global query never reaches. This is the
same state/city rotation that took the fire and drone markets from
stalling in the 80s to 250+, applied to countries instead.

Order matters: the AP automation channel is densest in North America and
Western Europe, so those run first and the long tail follows.

Usage:
    python run_ap_by_country.py                    # all countries, repeat passes
    python run_ap_by_country.py --target 250
    python run_ap_by_country.py --only Germany --only France
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from ap_automation_market import MARKET, SCOPE, KIND, configure  # noqa: E402

# Ordered by where the AP automation channel actually is. The US and
# Western Europe carry most vendor partner programmes; the rest still
# yield national ERP/accounting integrators that global queries miss.
COUNTRIES = (
    "United States",
    "United Kingdom",
    "Canada",
    "Germany",
    "Netherlands",
    "France",
    "Australia",
    "India",
    "Sweden",
    "Norway",
    "Denmark",
    "Finland",
    "Belgium",
    "Switzerland",
    "Ireland",
    "Spain",
    "Italy",
    "Poland",
    "Austria",
    "Singapore",
    "United Arab Emirates",
    "South Africa",
    "New Zealand",
    "Japan",
    "Brazil",
    "Mexico",
)


def sweep(country: str, batch_size: int, target: int) -> int:
    """One country's pass. Returns companies gained."""
    from vendor_intel.pipeline.market_research import batch_runner, discovery
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    before = len(MarketResearchState.load(MARKET).companies)
    if before >= target:
        return 0

    # Re-register the scope with this country's focus, and restrict the
    # answer to it. Both matter: the scope steers what to look for, the
    # restriction stops the model drifting back to US names.
    discovery.set_product_scope(
        MARKET,
        SCOPE + f" FOCUS THIS SEARCH ON {country.upper()}: return companies "
                f"based there. National and regional partners are wanted, "
                f"not only the global names.",
    )
    discovery.set_distributor_definition(MARKET, KIND)
    discovery.set_country_restriction(MARKET, [country])

    print(f"\n--- {country} (have {before}) ---", flush=True)
    try:
        # Ask for a modest amount per country rather than the full target:
        # the point is breadth across countries, not depth in one.
        goal = min(target, before + 40)
        batch_runner.run_until_target(MARKET, target=goal,
                                      batch_size=batch_size)
    except Exception as exc:  # noqa: BLE001
        print(f"  {country}: {type(exc).__name__}: {exc}", flush=True)

    after = len(MarketResearchState.load(MARKET).companies)
    gained = after - before
    print(f"  {country}: +{gained} (total {after})", flush=True)
    return gained


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--passes", type=int, default=3,
                    help="how many times to cycle the country list")
    ap.add_argument("--only", action="append", default=[],
                    help="restrict to these countries (repeatable)")
    args = ap.parse_args()

    configure()  # sets the profile dir + registers scope/definition

    countries = args.only or list(COUNTRIES)
    print(f"market : {MARKET}")
    print(f"target : {args.target}")
    print(f"sweep  : {len(countries)} countries x {args.passes} passes")
    print("", flush=True)

    from vendor_intel.pipeline.market_research.state import MarketResearchState

    total = 0
    try:
        for p in range(1, args.passes + 1):
            print(f"\n===== pass {p}/{args.passes} =====", flush=True)
            gain = 0
            for c in countries:
                gain += sweep(c, args.batch_size, args.target)
                if len(MarketResearchState.load(MARKET).companies) >= args.target:
                    print(f"\nreached target {args.target}", flush=True)
                    raise SystemExit(0)
            total += gain
            print(f"\npass {p} gained {gain}", flush=True)
            # A full pass finding nothing means the country list is spent;
            # more passes would only re-ask the same exhausted queries.
            if gain == 0:
                print("a whole pass found nothing new — stopping", flush=True)
                break
    except KeyboardInterrupt:
        print("\ninterrupted — progress is saved", flush=True)
    except SystemExit:
        pass

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx
    n = len(MarketResearchState.load(MARKET).companies)
    print("")
    print("=" * 60)
    print(f"SWEEP GAINED {total} — total discovered {n}")
    print(export_xlsx(MARKET), flush=True)


if __name__ == "__main__":
    main()
