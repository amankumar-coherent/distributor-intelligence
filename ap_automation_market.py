#!/usr/bin/env python3
"""Accounts Payable (AP) Automation Market — channel-partner mapping.

Finds the RESELLER, SYSTEM INTEGRATOR and TECHNOLOGY PARTNER ecosystem
around AP automation platforms, per the brief: "a map of the reseller,
system integrator, and technology partner ecosystem, highlighting which
partners competitors work with".

WHY A CUSTOM SCOPE RATHER THAN A PLAIN DEF_SOFTWARE MARKET
-----------------------------------------------------------
The built-in software/cloud rule (_DEF_SOFTWARE_SERVICES) says, verbatim:

    STRICTLY EXCLUDE ... SYSTEM INTEGRATORS and consultancies whose
    primary business is building or implementing bespoke solutions
    rather than reselling a platform.

That is right for a pure resale market, but it would discard a third of
what this brief asks for. DEF_HYBRID does admit system integrators, but
it expects the market to sell hardware too ("This market sells BOTH
hardware and software"), which is false here and would pull in hardware
distributors.

So: register the market as DEF_SOFTWARE (the channel partner IS the
distributor — correct economics for AP software) and use the per-market
PRODUCT SCOPE, which is free-form text concatenated into the same prompt,
to re-admit implementation and technology partners. No change to
discovery.py / batch_runner.py / verify_distributors_ai_mode.py, so the
drone and fire pipelines are untouched.

Market definition and taxonomy are the user's, quoted in SCOPE below.

Usage:
    python ap_automation_market.py --show-prompt
    python ap_automation_market.py --target 250
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

SLUG = "ap_automation"
# Deliberately short. The market name is interpolated three times into the
# field instructions, and the full "Global Accounts Payable Automation
# Market - Channel Partners" (60 chars) pushed the query to 7,903 URL
# chars against a 7,800 limit -- so every single query was truncated,
# losing the JSON reply template and coming back as prose. Trimming the
# name is the cheapest 100+ characters available.
MARKET = "Accounts Payable Automation Market"

# The scope carries three jobs: say what the market IS, re-admit the
# partner types the base software rule excludes, and draw the boundary
# the user's own definition draws ("We are not including AP Services").
# ORDER AND LENGTH HERE ARE LOAD-BEARING.
#
# Two constraints shape this text:
#
# 1. The base software rule (rendered ABOVE this scope) tells the model to
#    EXCLUDE system integrators, and the acceptance gate only admits the
#    verdicts {channel_partner, distributor} -- a company labelled
#    "system_integrator" is dropped no matter what the scope says. So the
#    scope cannot merely say "keep SIs"; it must tell the model which
#    VERDICT to use for them. That is what the partner-verdict sentence
#    below does, and it must come FIRST so it survives truncation.
#
# 2. _compact_scope() strips everything between "QUALIFYING...:" and
#    "EXCLUDE" when the prompt overruns the URL budget -- which it did
#    here, silently deleting the entire partner clause (1,460 -> 939
#    chars). So the partner rule must NOT sit in a QUALIFYING block, and
#    the whole scope must be short enough to avoid compaction at all.
SCOPE = (
    "PRODUCT SCOPE (required): ACCOUNTS PAYABLE (AP) AUTOMATION SOFTWARE - "
    "automates invoice capture, PO matching, approval routing, payment "
    "execution, reconciliation and ERP sync. Covers standalone AP "
    "platforms AND AP modules in ERP/procure-to-pay suites. "
    "A company that RESELLS, IMPLEMENTS or INTEGRATES another vendor's AP "
    'software qualifies - return verdict "channel_partner" for it: '
    "resellers, solution providers, implementation and ERP partners "
    "(SAP, Oracle, NetSuite, Dynamics, Sage), certified vendor partners. "
    "EXCLUDE the AP software vendors themselves and their sales arms; "
    "outsourced AP services and BPO invoice processing; referral-only "
    "affiliates. Name the AP vendor(s) it partners with."
)

# Registered as software: in AP automation the channel partner IS the
# distributor — nobody holds physical stock. The scope above re-admits
# the integrator shapes that rule would otherwise drop.
KIND = "software"


def configure(region: str | None = None, profile_suffix: str | None = None) -> str:
    """Point this process at its own profile and register the scope."""
    name = f"ap_profile_{SLUG}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    os.environ["GOOGLE_AI_MODE_PROFILE_DIR"] = str(ROOT / "data" / name)

    from vendor_intel.pipeline.market_research import discovery

    scope = SCOPE
    if region:
        # Kept SHORT on purpose. The long form of this suffix was 140
        # characters, which alone pushed the query 36 chars past the
        # 7,800 URL limit and got every single prompt truncated -- the
        # JSON template was cut off and the model answered in prose.
        scope += f" FOCUS ON {region.upper()}: companies based there."

    discovery.set_product_scope(MARKET, scope)
    discovery.set_distributor_definition(MARKET, KIND)
    return MARKET


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--region", default=None)
    ap.add_argument("--profile-suffix", default=None)
    ap.add_argument("--show-prompt", action="store_true")
    args = ap.parse_args()

    market = configure(args.region, args.profile_suffix)

    from vendor_intel.pipeline.market_research import batch_runner, discovery

    if args.show_prompt:
        print(f"market: {market}  [{KIND}]")
        print(discovery.build_draft_prompt(
            market, excluded=[], count=10, collected=[]))
        return

    print(f"market : {market}")
    print(f"rule   : {KIND} (+ integrator/tech-partner scope)")
    print(f"profile: {os.environ['GOOGLE_AI_MODE_PROFILE_DIR']}")
    print(f"target : {args.target}, coverage: {args.region or 'GLOBAL'}")
    print("", flush=True)

    print(batch_runner.run_until_target(
        market, target=args.target, batch_size=args.batch_size), flush=True)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx
    print(export_xlsx(market), flush=True)


if __name__ == "__main__":
    main()
