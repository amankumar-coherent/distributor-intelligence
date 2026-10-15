#!/usr/bin/env python3
"""Run ONE market from the 3-market global wound-care / medical-films batch.

Same architecture as run_food_market.py (own Chromium profile per market,
data/wound_profile_<slug>), but WORLDWIDE: no country restriction, and
--region names a country instead of an Indian state.

Usage:
    python run_wound_market.py --slug advanced_wound_care --target 250
    python run_wound_market.py --slug medical_films --region Germany --show-prompt
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

from wound_markets import SIBLINGS_WOUND, by_slug_wound  # noqa: E402


def configure(slug: str, region: str | None = None,
              profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope."""
    name = f"wound_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(ROOT / "data" / name)

    slug_, market, scope, kind = by_slug_wound(slug)
    from vendor_intel.pipeline.market_research import discovery

    full = scope
    siblings = SIBLINGS_WOUND.get(slug_, "")
    # Both sentences are SHORT on purpose: discovery sends system prompt +
    # this text in one Google URL (7,550-char budget). Longer wording put
    # Medical Films over Google's 7,800 hard limit, cutting off the JSON
    # reply format, so AI Mode answered in prose and found 0 companies.
    if siblings:
        full += f" This is NOT the {siblings} market."
    if region:
        full += (f" FOCUS ON {region.upper()}: distributors headquartered "
                 f"there, including regional and national ones.")

    discovery.set_product_scope(market, full)
    kinds = {"distributor": discovery.DEF_PHYSICAL,
             "service": discovery.DEF_SERVICE,
             "hybrid": discovery.DEF_HYBRID,
             "software": discovery.DEF_SOFTWARE}
    discovery.set_distributor_definition(market, kinds[kind])
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
    kind = by_slug_wound(args.slug)[3]

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
