#!/usr/bin/env python3
"""Run ONE drone / geospatial / digital-twin market in its own Chromium profile.

Same architecture as run_chemical_market.py: worldwide (no country
restriction), own profile per browser (data/geo_profile_<slug>[_wN]),
--region names a country so each browser covers different ground per round.
The rule (hybrid / software / service) comes from geo_markets.py.

Usage:
    python run_geo_market.py --slug geo_ag_drone --target 400
    python run_geo_market.py --slug geo_gis_analytics --region India --show-prompt
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

from geo_markets import SIBLINGS_GEO, by_slug_geo  # noqa: E402


def configure(slug: str, region: str | None = None,
              profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope."""
    name = f"geo_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(ROOT / "data" / name)

    slug_, market, scope, kind = by_slug_geo(slug)
    from vendor_intel.pipeline.market_research import discovery

    full = scope
    siblings = SIBLINGS_GEO.get(slug_, "")
    if siblings:
        full += f" This is NOT the {siblings} market."
    # "Worldwide", "Worldwide 2", ... = a global pass: no focus sentence, so
    # discovery's own global-coverage steering applies.
    if region and not region.lower().startswith("worldwide"):
        full += f" FOCUS ON {region.upper()}: companies headquartered there."

    discovery.set_product_scope(market, full)
    kinds = {"distributor": discovery.DEF_PHYSICAL,
             "service": discovery.DEF_SERVICE,
             "hybrid": discovery.DEF_HYBRID,
             "software": discovery.DEF_SOFTWARE}
    discovery.set_distributor_definition(market, kinds[kind])
    # User, 2026-10-14: "I want only India companies, not global". Enforced in
    # the prompt and again at the collection gate; the earlier worldwide finds
    # were reduced to India HQ by scripts/purge_to_country.py.
    discovery.set_country_restriction(market, ("India",))
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
    kind = by_slug_geo(args.slug)[3]

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
