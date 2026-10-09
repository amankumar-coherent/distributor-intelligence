#!/usr/bin/env python3
"""City-by-city sweep for the Dangjo countries that returned little or nothing.

A country-level question saturates the same way a national one does: ask
"Dangjo pepper distributors in Vietnam" twice and the second answer is
the first answer's exclusion list. Asking "in Ho Chi Minh City" is a
different question -- it reaches wholesale markets, importers and Korean
grocery chains that a national query never surfaces.

This is the technique that worked everywhere else in this project:
India state-by-state took drone_software from 145 to 585 and Dangjo
itself from 38 to 112, and South Korea city-level took it from 1 to 41.

Cities are chosen for where imported Korean food actually trades: port
and wholesale hubs, capitals, and districts with a Korean expatriate
population -- which is the real channel for a Korean cultivar.

Usage:
    python dangjo_cities.py                      # all weak countries
    python dangjo_cities.py --only Vietnam
    python dangjo_cities.py --batch-size 8
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

apply_env_overrides()

from run_dangjo import MARKET, KNOWN_EXAMPLES  # noqa: E402

# Dangjo registers no product scope of its own -- the market name has
# carried the definition so far. The city sweep needs a scope to attach
# the geography to, so state the cultivar rule explicitly here. It must
# match how the market was collected, or the city passes would widen the
# definition that the 59 wrongly-collected generic-chili companies proved
# has to stay narrow.
BASE_SCOPE = (
    "PRODUCT SCOPE (required): the DANGJO cultivar of Korean pepper "
    "(Capsicum annuum L. cv. Dangjo) in any traded form - fresh, dried, "
    "powder, flake, paste, extract or seed. Generic Korean chili, "
    "gochugaru or Cheongyang pepper alone does NOT qualify.")

# Country -> the cities where imported Korean produce actually changes
# hands. Ordered most-likely-first so an interrupted run still covers
# the best ground.
CITIES: dict[str, tuple[str, ...]] = {
    "China": (
        "Shanghai", "Beijing", "Qingdao", "Guangzhou", "Shenzhen",
        "Yanji and Jilin (Korean autonomous prefecture)", "Dalian",
        "Tianjin", "Shenyang",
    ),
    "Vietnam": (
        # Vietnam returned 19 companies but only one survived: the rest
        # carry gochujang and generic Korean chili, not the Dangjo
        # cultivar. These angles target the Korean-expat supply chain
        # where a specific cultivar actually moves, rather than the
        # general Korean-food importers already collected and rejected.
        "Phu My Hung and District 7 Korean district, Ho Chi Minh City",
        "My Dinh and Trung Hoa Korean district, Hanoi",
        "Ho Chi Minh City Korean grocery importers",
        "Hanoi Korean food wholesale importers",
        "Binh Tay and Thu Duc wholesale produce markets, Ho Chi Minh City",
        "Long Bien wholesale market, Hanoi",
        "Da Nang", "Hai Phong",
    ),
    "Philippines": (
        "Manila and Metro Manila", "Quezon City", "Cebu", "Davao",
        "Angeles City and Clark", "Makati and Bonifacio Global City",
    ),
    "Myanmar": ("Yangon", "Mandalay", "Naypyidaw"),
    "Sri Lanka": ("Colombo", "Kandy", "Galle", "Negombo"),
    "New Zealand": ("Auckland", "Wellington", "Christchurch", "Hamilton"),
    "Taiwan": ("Taipei", "Kaohsiung", "Taichung", "Tainan", "Taoyuan"),
    "Thailand": ("Bangkok", "Chiang Mai", "Pattaya and Chonburi",
                 "Phuket", "Samut Prakan"),
    "India": ("Delhi NCR", "Mumbai", "Bengaluru", "Chennai",
              "Pune", "Hyderabad", "Kolkata"),
}


def sweep_city(country: str, city: str, *, batch_size: int) -> int:
    """One city-scoped pass. Returns companies gained."""
    from vendor_intel.pipeline.market_research import batch_runner, discovery
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    before = len(MarketResearchState.load(MARKET).companies)

    # The collection gate still enforces the COUNTRY -- a city is a
    # search hint, not a new rule, so a company found via "Shanghai" is
    # still validated as being in China.
    #
    # run_until_target takes no region argument; the drone runners steer
    # geography by appending it to the PRODUCT SCOPE, which is the text
    # that rides in the prompt. Same mechanism here.
    discovery.set_country_restriction(MARKET, (country,))
    discovery.set_known_examples(MARKET, KNOWN_EXAMPLES)
    discovery.set_product_scope(
        MARKET,
        f"{BASE_SCOPE} FOCUS THIS SEARCH ON {city.upper()}, "
        f"{country.upper()}: return distributors, importers and wholesalers "
        f"based in or trading through this city, including Korean grocery "
        f"chains, wholesale produce markets and food importers. Local "
        f"city-level companies count, not only national names.")

    print(f"    {city} ...", end="", flush=True)
    try:
        batch_runner.run_until_target(
            MARKET, target=before + batch_size, batch_size=batch_size,
        )
    except KeyboardInterrupt:
        raise
    except Exception as exc:  # noqa: BLE001
        print(f" ! {type(exc).__name__}", flush=True)
        return 0

    gained = len(MarketResearchState.load(MARKET).companies) - before
    print(f" +{gained}", flush=True)
    return gained


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--only", action="append", default=[], metavar="COUNTRY")
    ap.add_argument("--batch-size", type=int, default=8)
    args = ap.parse_args()

    countries = args.only or list(CITIES)
    grand = 0
    for country in countries:
        cities = CITIES.get(country)
        if not cities:
            print(f"  no city list for {country!r} - skipped")
            continue
        print("")
        print("=" * 66)
        print(f"  {country}: {len(cities)} cities")
        print("=" * 66, flush=True)
        total = 0
        for city in cities:
            total += sweep_city(country, city, batch_size=args.batch_size)
            time.sleep(2)
        grand += total
        print(f"  -> {country}: +{total}", flush=True)

    print("")
    print(f"TOTAL GAINED: +{grand}")


if __name__ == "__main__":
    main()
