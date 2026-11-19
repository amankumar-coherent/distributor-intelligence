#!/usr/bin/env python3
"""Run ONE of the 6 Potassium Sorbate / Sorbic Acid / Caffeine markets (Global + India) in its own Chromium profile.

Global markets: worldwide country by country; India markets: India HQ only
(state by state): own profile per browser (data/sorb_profile_<slug>[_wN]),
--region names a country so each round covers a different country.
Company type (distributor rule)
and scope come from sorb_markets.py.

Usage:
    python run_sorb_market.py --slug sorb01_global_potassium_sorbate --target 400
    python run_sorb_market.py --slug sorb01_global_potassium_sorbate --region Germany --show-prompt
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

from sorb_markets import SIBLINGS_SORB, by_slug_sorb, INDIA_SLUGS  # noqa: E402


def configure(slug: str, region: str | None = None,
              profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope."""
    name = f"sorb_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(ROOT / "data" / name)

    slug_, market, scope, kind = by_slug_sorb(slug)
    from vendor_intel.pipeline.market_research import discovery

    full = scope
    siblings = SIBLINGS_SORB.get(slug_, "")
    if siblings:
        full += f" This is NOT the {siblings} market."
    if region and region.lower().startswith("worldwide"):
        # Global round (user, 2026-10-30): no country focus at all.
        full += " Search WORLDWIDE: companies headquartered in any country, on every continent."
    elif region:
        full += f" FOCUS ON {region.upper()}: companies headquartered there."

    discovery.set_product_scope(market, full)
    # Simple query mode (2026-11-13): AI Mode refuses the long rule-heavy prompt.
    from sorb_markets import SIMPLE_REQUEST
    india = slug_ in INDIA_SLUGS
    if india:
        # India module (user 2026-11-16): India-HQ companies only, state by state.
        discovery.set_country_restriction(market, ("India",))
        where = f"in {region}, India" if region else "in India"
    else:
        where = ("in any country" if not region or region.lower().startswith("worldwide")
                 else f"in {region}")
    discovery.set_simple_query(market, f"{SIMPLE_REQUEST[slug_]} headquartered {where}")
    kinds = {"distributor": discovery.DEF_PHYSICAL,
             "service": discovery.DEF_SERVICE,
             "hybrid": discovery.DEF_HYBRID,
             "software": discovery.DEF_SOFTWARE,
             "retail": discovery.DEF_RETAIL}
    discovery.set_distributor_definition(market, kinds[kind])
    # User 2026-11-16: Global module (no restriction) + India module (India HQ) per product.
    return market


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slug", required=True)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--region", default=None)
    ap.add_argument("--profile-suffix", default=None,
                    help="run this market in an ISOLATED profile, e.g. w2, "
                    "so multiple workers can cover it at once")
    ap.add_argument("--show-prompt", action="store_true")
    args = ap.parse_args()

    market = configure(args.slug, args.region, args.profile_suffix)
    kind = by_slug_sorb(args.slug)[3]

    from vendor_intel.pipeline.market_research import batch_runner, discovery

    if args.show_prompt:
        print(f"market: {market}  [{kind}]")
        print(discovery.build_draft_prompt(
            market, excluded=[], count=10, collected=[]))
        return

    print(f"market : {market}")
    print(f"rule   : {kind}")
    print(f"profile: {os.environ['GOOGLE_AI_MODE_PROFILE_DIR']}")
    print(f"target : {args.target}, coverage: {args.region or 'WORLDWIDE'}")
    print("", flush=True)

    print(batch_runner.run_until_target(
        market, target=args.target, batch_size=args.batch_size), flush=True)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    print("export: " + str(export_xlsx(market)), flush=True)


if __name__ == "__main__":
    main()

