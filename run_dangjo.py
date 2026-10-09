#!/usr/bin/env python3
"""Dangjo Pepper (당조고추) ingredient distributors -- STRICT.

Dangjo is a specific functional cultivar of Capsicum annuum L., bred in
South Korea in 2008 for blood-glucose regulation, with quercitrin as its
marker compound. It is NOT a synonym for Korean chili.

An earlier run broadened the term to "Korean chili pepper" so the backend
would answer. It answered about the wrong companies: 59 gochugaru,
gochujang and kimchi importers, of which ZERO mentioned Dangjo anywhere.
Those were discarded rather than verified.

EXPECT A SMALL RESULT. Probing the backend directly before this run:

  * asked what Dangjo pepper is, it answered correctly and in detail
    (cultivar, quercitrin, alpha-glucosidase inhibitors, with
    ScienceDirect/NIH/KCI citations) -- so the term is understood;
  * asked to name companies supplying Dangjo powder or extract, it
    answered exactly "NONE KNOWN";
  * asked the same in Korean, it named only Austem Bio (오스템바이오) and
    Jeil Seed Bio (제일씨드바이오) -- the SEED DEVELOPERS who co-created
    the cultivar, not ingredient distributors.

So a 400-company target is not reachable for this definition, and the
run is expected to stop early on the low-yield rule rather than to fill.
That is the honest answer for this market, and it is the requested one:
only genuine Dangjo ingredient distributors, however few.

Usage:
    python run_dangjo.py --target 400
    python run_dangjo.py --show-prompt
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

MARKET = "Dangjo Pepper (Capsicum annuum L. cv. Dangjo) Ingredients Market - Distributors Only"

# The brief's 16 APAC countries, plus the Western markets its
# country-specific query list also names (USA, Canada, UK, Germany,
# Netherlands, UAE). Korea is where the cultivar is grown and processed,
# so it is the single most likely source of a real hit.
# EXACTLY the 16 countries requested -- no more.
#
# An earlier version added the US, Canada, UK, Germany, the Netherlands
# and the UAE because the brief's keyword list mentioned them in its
# search strings. That was wrong: those were SEARCH terms, while the
# country list is the REQUIREMENT ("I also mention from which country
# only I want distributor"). The mistake let 12 out-of-scope companies
# into the collection before it was caught.
TARGET_COUNTRIES = (
    "South Korea", "China", "Japan", "Taiwan", "Hong Kong", "Indonesia",
    "Malaysia", "Myanmar", "Philippines", "Singapore", "Thailand",
    "Vietnam", "India", "Sri Lanka", "Australia", "New Zealand",
)


# Companies confirmed to trade Dangjo, each with a traceable reason.
# Shown to the backend as examples: asked cold it answered "NONE KNOWN",
# but shown these it produced further real names with sources.
#
# Austembio and Jeil Seed Bio are deliberately absent -- they are the
# SEED developers of the cultivar, which the brief excludes.
# Ordered by how much each teaches the model, because only the first
# ~60 characters fit the URL budget. Short names from DIFFERENT channels
# (fresh wholesale, seed distribution) calibrate better than a long list
# from one channel.
KNOWN_EXAMPLES = (
    "Hantar Fresh",              # Malaysia, wholesales KFresh Korea Dangjo
    "Osborne Seed Company",      # US, distributes Dangjo Cheongyang seed
    "Euro-Atlantic Sdn Bhd",     # Malaysia, K Fresh portfolio lists Dangjo
    "Log House Plants",          # US, Dangjo Cheongyang seedlings
    "Dream of Farmer Co., Ltd.", # Korea, produces and exports Dangjo
    "Freshmart Singapore",       # Singapore, K-Fresh importer
    "Nara Company Co., Ltd.",    # Korea, Dangjo brand distributor
    "H Mart",                    # US/Canada, imports fresh Dangjo
)


def register() -> None:
    """Attach the country restriction and examples. In-process with discovery."""
    from vendor_intel.pipeline.market_research import discovery

    discovery.set_country_restriction(MARKET, TARGET_COUNTRIES)
    discovery.set_known_examples(MARKET, KNOWN_EXAMPLES)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--show-prompt", action="store_true",
                    help="print the first prompt and exit, without querying")
    args = ap.parse_args()

    register()

    from vendor_intel.pipeline.market_research import batch_runner, discovery

    if args.show_prompt:
        print(discovery.build_draft_prompt(
            MARKET, excluded=[], count=10, collected=[]))
        return

    print(f"market : {MARKET}")
    print(f"target : {args.target} verified")
    print(f"countries ({len(TARGET_COUNTRIES)}): {', '.join(TARGET_COUNTRIES)}")
    print("")
    print("NOTE: the backend answered 'NONE KNOWN' when asked directly for")
    print("Dangjo ingredient suppliers, so a small result is the expected")
    print("outcome here, not a failure of the run.")
    print("")

    summary = batch_runner.run_until_target(
        MARKET, target=args.target, batch_size=args.batch_size,
    )
    print(summary)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    print("")
    print("export: " + str(export_xlsx(MARKET)))


if __name__ == "__main__":
    main()
