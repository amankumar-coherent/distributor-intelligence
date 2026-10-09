#!/usr/bin/env python3
"""Run ONE market from the second batch of 15, in its own Chromium.

Same architecture as run_drone_market.py: a dedicated profile per market
so fifteen can run at once (a Chromium profile takes an exclusive lock,
and sharing one produces about:blank tabs and zero results).

Differences from batch one:
  * four markets are SERVICES (inspection, surveillance, testing,
    training) and use the service-provider definition rather than a
    reseller rule;
  * three are hybrid hardware+software (mapping, cybersecurity, cloud).

India-only, enforced in the prompt and again at the collection gate.

Usage:
    python run_drone2_market.py --slug drone_lidar --target 400
    python run_drone2_market.py --slug drone_training --show-prompt
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

from drone_markets2 import SIBLINGS2, by_slug2  # noqa: E402


def configure(slug: str, region: str | None = None,
             profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope.

    `profile_suffix` lets multiple workers cover the SAME market in
    parallel without colliding: a Chromium profile takes an exclusive
    lock, so two processes on drone2_profile_<slug> would produce
    about:blank tabs. Each parallel worker gets drone2_profile_<slug>_w2
    (etc) instead, with its own captcha-raptor extension staged the same
    way as the base profile.
    """
    name = f"drone2_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    profile = ROOT / "data" / name
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(profile)

    slug_, market, scope, kind = by_slug2(slug)
    from vendor_intel.pipeline.market_research import discovery

    siblings = SIBLINGS2.get(slug_, "")
    full = scope
    if siblings:
        full += (f" This is NOT the {siblings} market; a company belongs "
                 f"here only if it offers what is named above.")
    if region:
        full += (f" FOCUS THIS SEARCH ON {region.upper()}: return companies "
                 f"based there. Regional and state-level firms are wanted, "
                 f"not only national names.")

    discovery.set_product_scope(market, full)
    # drone_markets2 uses the business-facing word "distributor"; the
    # library's constant for stocked goods is "physical".
    _KIND = {"distributor": discovery.DEF_PHYSICAL,
             "service": discovery.DEF_SERVICE,
             "hybrid": discovery.DEF_HYBRID,
             "software": discovery.DEF_SOFTWARE}
    discovery.set_distributor_definition(market, _KIND[kind])
    discovery.set_country_restriction(market, ("India",))

    # Service markets whose NAME collides with an equipment category get
    # searched under a service-led phrasing. "Drone Testing Market -
    # Distributors Only" reads as test-equipment distribution: measured,
    # drone_testing passed 19% and returned Caltest Instruments and
    # Sapphire Technologies, both correctly rejected as "distributor of
    # test and measurement instruments, does not provide testing
    # services". drone_inspection has no such collision and passes 59%.
    # Only the search text changes; the stored market name is untouched.
    _ALIAS = {
        "drone_testing":
            "Drone Testing, Certification and Compliance SERVICE PROVIDERS "
            "in India - accredited laboratories and test agencies that "
            "PERFORM testing for third-party clients",
        "drone_surveillance":
            "Drone Surveillance and Aerial Monitoring SERVICE OPERATORS in "
            "India - companies that FLY surveillance and security missions "
            "for third-party clients",
        "drone_training":
            "Drone Pilot Training ACADEMIES and DGCA-approved RPTOs in "
            "India that RUN certification courses",
        "drone_inspection":
            "Drone-based Industrial Inspection SERVICE PROVIDERS in India "
            "that PERFORM asset inspection for third-party clients",
    }
    if slug_ in _ALIAS:
        discovery.set_search_alias(market, _ALIAS[slug_])
    return market


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slug", required=True)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--region", default=None)
    ap.add_argument("--profile-suffix", default=None,
                    help="run this slug's market in an ISOLATED profile, "
                    "e.g. w2, so multiple workers can cover it at once")
    ap.add_argument("--show-prompt", action="store_true")
    args = ap.parse_args()

    market = configure(args.slug, args.region, args.profile_suffix)
    kind = by_slug2(args.slug)[3]

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
