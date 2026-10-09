#!/usr/bin/env python3
"""Sweep Dangjo pepper distributors ONE COUNTRY AT A TIME.

The global query has gone dry at 38 companies, but the distribution is
lopsided in a way that says the market is not actually exhausted:

    Hong Kong 6, Australia 5, Singapore 3 ... but
    South Korea 1, China 1, Japan 1, Myanmar 0, Sri Lanka 0

South Korea GROWS this pepper and should be the densest source of all.
One company there means the global query stopped early, not that Korea
has one distributor.

Why one query per country finds more than one global query:

  * The model answers a broad question with the most famous names it
    knows, and once those are excluded it returns [] rather than digging
    into a specific market.
  * "Distributors in Myanmar" is a different, narrower question than
    "distributors worldwide, excluding these 38" -- it pushes the model
    into local trade names, importers and wholesalers it would never
    volunteer globally.
  * Each country query carries the same exclusion list, so anything
    already found is skipped rather than re-offered.

Every company is still gated by the same country restriction and the
same Dangjo product scope, so this widens the SEARCH without loosening
the criteria.

Usage:
    python run_dangjo_by_country.py                # all 16, 2 passes
    python run_dangjo_by_country.py --passes 3
    python run_dangjo_by_country.py --only "South Korea" --only Japan
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

from run_dangjo import MARKET, TARGET_COUNTRIES, KNOWN_EXAMPLES  # noqa: E402

# Korea first: it is where the cultivar is grown, processed and exported
# from, so it has the most to give and its answers teach the model what a
# real Dangjo company looks like for the countries that follow.
SWEEP_ORDER = (
    "South Korea", "China", "Japan", "Taiwan", "Hong Kong",
    "Singapore", "Malaysia", "Thailand", "Vietnam", "Indonesia",
    "Philippines", "India", "Australia", "New Zealand",
    "Myanmar", "Sri Lanka",
)


def sweep(country: str, *, batch_size: int) -> int:
    """Run one country-scoped discovery pass. Returns companies gained."""
    from vendor_intel.pipeline.market_research import batch_runner, discovery
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    before = len(MarketResearchState.load(MARKET).companies)

    # Narrow the restriction to this ONE country for the duration of the
    # pass. The prompt then says "only <country> counts", which is a much
    # sharper instruction than a 16-country list, and the collection gate
    # enforces the same thing.
    discovery.set_country_restriction(MARKET, (country,))
    discovery.set_known_examples(MARKET, KNOWN_EXAMPLES)

    print("")
    print("=" * 66)
    print(f"  {country}  (have {before} overall)")
    print("=" * 66, flush=True)

    try:
        batch_runner.run_until_target(
            MARKET, target=before + batch_size, batch_size=batch_size,
        )
    except KeyboardInterrupt:
        raise
    except Exception as exc:  # noqa: BLE001
        print(f"  ! {type(exc).__name__}: {exc}", flush=True)

    after = len(MarketResearchState.load(MARKET).companies)
    gained = after - before
    print(f"  -> {country}: +{gained} (total {after})", flush=True)
    return gained


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--passes", type=int, default=2,
                    help="times to cycle the whole country list")
    ap.add_argument("--batch-size", type=int, default=10,
                    help="companies requested per country pass")
    ap.add_argument("--only", action="append", default=[], metavar="COUNTRY")
    args = ap.parse_args()

    countries = [c for c in SWEEP_ORDER
                 if not args.only or c in args.only]
    if not countries:
        print(f"no match; valid: {', '.join(SWEEP_ORDER)}")
        raise SystemExit(2)

    total = 0
    try:
        for p in range(1, args.passes + 1):
            print("")
            print("#" * 66)
            print(f"# PASS {p}/{args.passes}")
            print("#" * 66, flush=True)
            pass_gain = 0
            for country in countries:
                pass_gain += sweep(country, batch_size=args.batch_size)
            total += pass_gain
            print(f"\npass {p} gained {pass_gain}", flush=True)
            if pass_gain == 0:
                print("a whole pass found nothing new — stopping", flush=True)
                break
    except KeyboardInterrupt:
        print("\ninterrupted — progress is saved", flush=True)
    finally:
        # Restore the full restriction so a later global run is unaffected.
        from vendor_intel.pipeline.market_research import discovery
        discovery.set_country_restriction(MARKET, TARGET_COUNTRIES)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    print("")
    print("=" * 66)
    print(f"SWEEP GAINED {total} — total "
          f"{len(MarketResearchState.load(MARKET).companies)}")
    print("=" * 66)
    print("export: " + str(export_xlsx(MARKET)))


if __name__ == "__main__":
    main()
