"""PEM & AEM electrolyzers for green hydrogen -- ONE market, India-HQ only
(user, 2026-10-14).

Rule: HYBRID (distributor / channel partner). Electrolyzers are capital
equipment sold as stacks and as integrated systems, so two shapes qualify:
an authorized distributor / dealer / sales agent of a PEM or AEM electrolyzer
brand, and a system integrator / EPC / balance-of-plant packager that buys a
third-party PEM or AEM stack and delivers it inside its own system. The
electrolyzer makers themselves are excluded (the hybrid rule's "a company
that BUILDS what it sells is a manufacturer" clause), as in every batch.

The user's exclusions (AWE, SOEC, fuel cells, H2 storage / transport /
distribution, grey / blue / brown H2 equipment) are applied to COMPANIES too:
a company whose electrolyzer line is only alkaline or solid oxide, or that
only deals in fuel cells, storage, cylinders, pipelines or SMR / gasifier
equipment, is out.

The scope rides in ONE Google URL with the system prompt (budget 7,550 of a
7,800 hard limit) -- keep it near 600 characters; measure with
`python run_h2_market.py --slug h2_pem_aem_electrolyzer --show-prompt`.
"""
from __future__ import annotations

MARKETS_H2: tuple[tuple[str, str, str, str], ...] = (
    (
        "h2_pem_aem_electrolyzer",
        "India PEM & AEM Electrolyzers for Green Hydrogen Projects Market",
        "PRODUCT SCOPE (required): PEM (proton exchange membrane) and AEM "
        "(anion exchange membrane) WATER ELECTROLYZERS for green hydrogen - "
        "electrolyzer stacks and integrated electrolyzer systems, below 1 MW "
        "to 20 MW+, for green ammonia, refining, steel, chemicals and energy "
        "storage projects. Brands include Siemens Energy, ITM Power, Plug "
        "Power, Accelera (Cummins), Nel PEM, H-TEC, Enapter (AEM), Ohmium. "
        "EXCLUDE the electrolyzer makers themselves, and companies whose line "
        "is only alkaline (AWE) or solid oxide (SOEC) electrolyzers, fuel "
        "cells, hydrogen storage, transport or distribution, or grey/blue "
        "hydrogen equipment. Name the PEM or AEM brands it sells or "
        "integrates.",
        "hybrid",
    ),
)

SIBLINGS_H2: dict[str, str] = {}

# Client-facing definition for row 2 of the final Excel: the user's own
# wording (definition, inclusion taxonomy, exclusions), India-only.
DEFINITIONS: dict[str, str] = {
    "h2_pem_aem_electrolyzer": (
        "The PEM & AEM Electrolyzers for Green Hydrogen Projects Market refers "
        "to the sales of Proton Exchange Membrane (PEM) and Anion Exchange "
        "Membrane (AEM) electrolyzers that use renewable electricity to split "
        "water into hydrogen and oxygen, producing green hydrogen. The market "
        "covers electrolyzer stacks and integrated systems deployed across "
        "industrial and commercial green hydrogen projects. "
        "By Product: PEM Electrolyzers, AEM Electrolyzers. "
        "By Configuration: Electrolyzer Stacks, Integrated Electrolyzer Systems. "
        "By System Capacity: Below 1 MW, 1 MW to Below 5 MW, 5 MW to Below "
        "20 MW, 20 MW and Above. "
        "By Application: Green Ammonia Production, Petroleum Refining, Iron & "
        "Steel Manufacturing, Chemical Manufacturing, Energy Storage & Power, "
        "Other Industrial Applications. "
        "By End User: Green Hydrogen Project Developers, Industrial Hydrogen "
        "Producers, Captive Industrial Users, Utilities & Energy Companies. "
        "Geography: India-headquartered companies only. "
        "EXCLUDED: Alkaline Water Electrolyzers (AWE), Solid Oxide "
        "Electrolyzers (SOEC), hydrogen fuel cells, hydrogen storage, "
        "transportation and distribution, and grey, blue and brown hydrogen "
        "production equipment - and companies that deal only in those."
    ),
}


def by_slug_h2(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_H2:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown h2 market slug: {slug}")

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified company whose profile never matches this goes to
# flagged_for_review.xlsx instead of the main Excel. Tested on the delivered
# list 2026-10-21 (all genuine rows match).
EVIDENCE: dict[str, str] = {
    "h2_pem_aem_electrolyzer": r"electroly[sz]|pem\b|aem\b|green hydrogen|hydrogen (generation|plant|system)",
}
