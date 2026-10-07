#!/usr/bin/env python3
"""Run ONE market from the 25-market Fire Safety batch, in its own Chromium.

Same architecture as run_drone2_market.py: a dedicated profile per market
so many can run at once (a Chromium profile takes an exclusive lock, and
sharing one produces about:blank tabs and zero results).

12 markets are physical-goods DISTRIBUTORS, 9 are HYBRID (integrated
systems / hardware+software, channel partners and system integrators
count), 4 are SERVICES (fire_prevention, disaster_response,
fire_safety_training, fire_safety_testing) and use the service-provider
definition rather than a reseller rule -- see fire_markets.py's own
docstring for the full classification and reasoning.

India-only, enforced in the prompt and again at the collection gate.

Usage:
    python run_fire_market.py --slug fire_doors --target 250
    python run_fire_market.py --slug fire_safety_testing --show-prompt
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

from fire_markets import SIBLINGS_FIRE, by_slug_fire  # noqa: E402


def configure(slug: str, region: str | None = None,
             profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope.

    `profile_suffix` lets multiple workers cover the SAME market in
    parallel without colliding: a Chromium profile takes an exclusive
    lock, so two processes on fire_profile_<slug> would produce
    about:blank tabs. Each parallel worker gets fire_profile_<slug>_w2
    (etc) instead, with its own captcha-raptor extension staged the same
    way as the base profile.
    """
    name = f"fire_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    profile = ROOT / "data" / name
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(profile)

    slug_, market, scope, kind = by_slug_fire(slug)
    from vendor_intel.pipeline.market_research import discovery

    siblings = SIBLINGS_FIRE.get(slug_, "")
    full = scope
    if siblings:
        full += (f" This is NOT the {siblings} market; a company belongs "
                 f"here only if it offers what is named above.")
    if region:
        full += (f" FOCUS THIS SEARCH ON {region.upper()}: return companies "
                 f"based there. Regional and state-level firms are wanted, "
                 f"not only national names.")

    discovery.set_product_scope(market, full)
    _KIND = {"distributor": discovery.DEF_PHYSICAL,
             "service": discovery.DEF_SERVICE,
             "hybrid": discovery.DEF_HYBRID,
             "software": discovery.DEF_SOFTWARE}
    discovery.set_distributor_definition(market, _KIND[kind])
    discovery.set_country_restriction(market, ("India",))

    # Service markets whose NAME collides with a plain equipment
    # category get searched under a service-led phrasing -- the same
    # fix applied to drone_testing/surveillance/training/inspection,
    # where "Drone Testing Market" read as test-equipment distribution
    # until the search text was rewritten to lead with "SERVICE
    # PROVIDERS ... PERFORM". Only the search text changes; the stored
    # market name is untouched.
    _ALIAS = {
        "fire_prevention":
            "Fire Risk Assessment and Prevention CONSULTING SERVICES in "
            "India - firms that PERFORM fire risk assessment, hazard "
            "identification and prevention planning for third-party "
            "clients",
        "disaster_response":
            "Disaster and Emergency RESPONSE SERVICE providers in India - "
            "organizations that PERFORM search & rescue, hazmat response "
            "and disaster recovery operations for third-party clients or "
            "authorities",
        "fire_safety_training":
            "Fire Safety Training ACADEMIES and INSTITUTES in India that "
            "RUN firefighting, evacuation and fire-safety certification "
            "courses",
        "fire_safety_testing":
            "Fire Safety Testing, Inspection and Certification SERVICE "
            "PROVIDERS in India - accredited laboratories and "
            "certification bodies that PERFORM fire testing and "
            "compliance certification for third-party clients",
    }
    if slug_ in _ALIAS:
        discovery.set_search_alias(market, _ALIAS[slug_])
    return market


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slug", required=True)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--region", default=None)
    ap.add_argument("--profile-suffix", default=None,
                    help="run this slug's market in an ISOLATED profile, "
                    "e.g. w2, so multiple workers can cover it at once")
    ap.add_argument("--show-prompt", action="store_true")
    args = ap.parse_args()

    market = configure(args.slug, args.region, args.profile_suffix)
    kind = by_slug_fire(args.slug)[3]

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
