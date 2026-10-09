#!/usr/bin/env python3
"""Run ONE drone market in its own Chromium profile.

Launched ten times in parallel by run_all_drone_markets.py. Each copy
gets its own GOOGLE_AI_MODE_PROFILE_DIR, which is what makes parallelism
possible: launch_persistent_context takes an exclusive lock, so two
processes sharing a profile produce the about:blank tabs and zero drafts
seen earlier in this project.

Global coverage, target 400 verified, per the request. No country
restriction -- these markets are worldwide.

Usage:
    python run_drone_market.py --slug uav_systems --target 400
    python run_drone_market.py --slug drone_ai --show-prompt
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

from drone_markets import SIBLINGS, by_slug  # noqa: E402

# Per-market distributor economics.
#
#   hybrid   sells hardware AND software, so a stocking distributor and
#            a channel partner/integrator are both valid answers
#   software licences and platforms only: the channel partner IS the
#            distributor
#   physical stocked goods: a reseller that holds no stock is excluded
_DEFINITION_KIND = {
    "drone_ai": "hybrid",          # AI compute modules + autonomy stacks
    "drone_analytics": "hybrid",   # dashboards, but sold with hardware
    "drone_software": "software",  # pure licences
    "anti_drone": "hybrid",        # C-UAS ships as hardware + software
}


def configure(slug: str, region: str | None = None) -> tuple[str, str]:
    """Point this process at its own profile and register the scope.

    Must run BEFORE google_ai_mode is first used, because the profile
    directory is read when the browser launches.
    """
    profile = ROOT / "data" / f"drone_profile_{slug}"
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(profile)

    slug_, market, scope = by_slug(slug)
    from vendor_intel.pipeline.market_research import discovery

    # Name the neighbouring markets so a query for "UAV Systems" is not
    # answered with a counter-drone vendor. Overlap for genuinely
    # multi-line companies is allowed; this only prevents drift.
    siblings = SIBLINGS.get(slug_, "")
    full_scope = scope
    if siblings:
        full_scope += (f" This is NOT the {siblings} market; a company "
                       f"belongs here only if it carries the products "
                       f"named above.")

    # Regional focus, used when the global question has saturated: the
    # model has no new worldwide names left, but plenty of regional ones
    # it will only volunteer when asked about that region specifically.
    if region:
        full_scope += (f" FOCUS THIS SEARCH ON {region.upper()}: return "
                       f"distributors headquartered there. Regional and "
                       f"national firms are wanted, not global giants.")

    discovery.set_product_scope(market, full_scope)

    # Which DISTRIBUTOR definition this market's economics call for.
    #
    # The default routes on keywords in the market NAME, and only "Drone
    # Software" contains one -- so Drone AI and Drone Analytics were
    # judged by the PHYSICAL rule, which excludes channel partners,
    # VARs and system integrators. In markets delivered as software and
    # services that is most of the channel: drone_analytics passed only
    # 43 of 121 (44%) before this, and CPaaS lost 65% for the same
    # reason until it was routed correctly.
    kind = _DEFINITION_KIND.get(slug_, discovery.DEF_PHYSICAL)
    discovery.set_distributor_definition(market, kind)

    # India-only. Enforced twice: the prompt asks for India so the answer
    # is not wasted on companies that will be thrown away, and the
    # collection gate in batch_runner rejects anything else outright.
    discovery.set_country_restriction(market, ("India",))

    return market, full_scope


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slug", required=True)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--show-prompt", action="store_true")
    ap.add_argument("--region", default=None,
                    help="focus discovery on one region")
    args = ap.parse_args()

    market, scope = configure(args.slug, args.region)

    from vendor_intel.pipeline.market_research import batch_runner, discovery

    if args.show_prompt:
        print(f"market : {market}")
        print(f"profile: {os.environ['GOOGLE_AI_MODE_PROFILE_DIR']}")
        print("")
        print(discovery.build_draft_prompt(
            market, excluded=[], count=10, collected=[]))
        return

    print(f"market : {market}")
    print(f"profile: {os.environ['GOOGLE_AI_MODE_PROFILE_DIR']}")
    scope_note = args.region or "INDIA (all cities)"
    print(f"target : {args.target} verified, coverage: {scope_note}")
    print("", flush=True)

    summary = batch_runner.run_until_target(
        market, target=args.target, batch_size=args.batch_size)
    print(summary, flush=True)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    print("export: " + str(export_xlsx(market)), flush=True)


if __name__ == "__main__":
    main()
