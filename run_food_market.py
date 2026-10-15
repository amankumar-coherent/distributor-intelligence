#!/usr/bin/env python3
"""Run ONE market from the 75-market ProFood India batch, in its own Chromium.

Same architecture as run_fire_market.py: a dedicated profile per market
(data/food_profile_<slug>) so many can run at once -- a Chromium profile
takes an exclusive lock, and sharing one produces about:blank tabs.

Distributor / hybrid / software rule per market comes from food_markets.py.
India-only, enforced in the prompt and again at the collection gate.

Usage:
    python run_food_market.py --slug dairy_products --target 250
    python run_food_market.py --slug food_enzymes --show-prompt
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from dotenv import load_dotenv  # noqa: E402

load_dotenv(ROOT / ".env")

from food_markets import SIBLINGS_FOOD, by_slug_food  # noqa: E402


def configure(slug: str, region: str | None = None,
              profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope."""
    name = f"food_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(ROOT / "data" / name)

    slug_, market, scope, kind = by_slug_food(slug)
    from vendor_intel.pipeline.market_research import discovery

    full = scope
    siblings = SIBLINGS_FOOD.get(slug_, "")
    if siblings:
        full += (f" This is NOT the {siblings} market; a company belongs "
                 f"here only if it offers what is named above.")
    if region:
        full += (f" FOCUS THIS SEARCH ON {region.upper()}: return companies "
                 f"based there. Regional and state-level firms are wanted, "
                 f"not only national names.")

    discovery.set_product_scope(market, full)
    kinds = {"distributor": discovery.DEF_PHYSICAL,
             "service": discovery.DEF_SERVICE,
             "hybrid": discovery.DEF_HYBRID,
             "software": discovery.DEF_SOFTWARE}
    discovery.set_distributor_definition(market, kinds[kind])
    discovery.set_country_restriction(market, ("India",))
    return market


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slug", required=True)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--region", default=None)
    ap.add_argument("--profile-suffix", default=None,
                    help="run this market in an ISOLATED profile, e.g. w2, "
                    "so multiple workers can cover it at once")
    ap.add_argument("--show-prompt", action="store_true")
    args = ap.parse_args()

    market = configure(args.slug, args.region, args.profile_suffix)
    kind = by_slug_food(args.slug)[3]

    from vendor_intel.pipeline.market_research import batch_runner, discovery

    if args.show_prompt:
        print(f"market: {market}  [{kind}]")
        print(discovery.build_draft_prompt(
            market, excluded=[], count=10, collected=[]))
        return

    print(f"market : {market}")
    print(f"rule   : {kind}")
    print(f"profile: {os.environ['GOOGLE_AI_MODE_PROFILE_DIR']}")
    print(f"target : {args.target}, coverage: "
          f"{args.region or 'INDIA (all states)'}")
    print("", flush=True)

    print(batch_runner.run_until_target(
        market, target=args.target, batch_size=args.batch_size), flush=True)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    print("export: " + str(export_xlsx(market)), flush=True)


if __name__ == "__main__":
    main()
