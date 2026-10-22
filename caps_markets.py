"""Global Empty Capsules Market -- ONE market, worldwide, DISTRIBUTORS (user,
2026-10-15).

Input: "Report Sample_Global Empty Capsules Market_Coherent Market
Insights.pdf" (Downloads). The report has no written definition paragraph
(placeholder text), so the definition is built from its segmentation (capsule
type, material, functionality, shell type, size, purchase mode, application,
end user, region) and its "Distributors Landscape" table (manufacturer ->
distributor -> product), which is the company type collected here:
distributors buying empty capsules from capsule makers (Lonza/Capsugel, ACG,
Qualicaps, Suheung, CapsCanada, Sunil...) and reselling them to pharma,
nutraceutical, CRO and cosmetic manufacturers. Physical-goods DISTRIBUTOR rule.

Scope rides in ONE Google URL with the system prompt (budget 7,550 / hard
7,800); measure with `python run_caps_market.py --slug caps_distributor
--region "United Arab Emirates" --show-prompt`.
"""
from __future__ import annotations

MARKETS_CAPS: tuple[tuple[str, str, str, str], ...] = (
    (
        "caps_distributor",
        "Global Empty Capsules Market",
        "PRODUCT SCOPE (required): EMPTY CAPSULES - empty hard and soft "
        "capsule shells, gelatin (porcine, bovine, fish) and non-gelatin "
        "(HPMC, pullulan, carrageenan, PVA), immediate, delayed and sustained "
        "release, sizes 000 to 5, sold to pharmaceutical, nutraceutical, CRO "
        "and cosmetic manufacturers. Brands include Capsugel (Lonza), ACG, "
        "Qualicaps, Suheung, CapsCanada, Sunil Healthcare, Natural Capsules. "
        "EXCLUDE capsule manufacturers, contract manufacturers filling "
        "capsules, capsule-filling machine sellers, finished supplement "
        "brands and consumer-only pharmacies or shops. Name the capsule "
        "brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_CAPS: dict[str, str] = {}

# The report's regions in its own order, then the rest of the world.
_REPORT_COUNTRIES = (
    "United States", "Canada", "United Kingdom", "Germany", "Italy", "France",
    "Spain", "Russia", "China", "India", "Japan", "South Korea", "Singapore",
    "Malaysia", "Indonesia", "Thailand", "Vietnam", "Philippines", "Australia",
    "Brazil", "Argentina", "Mexico", "United Arab Emirates", "Saudi Arabia",
    "Qatar", "Kuwait", "Oman", "Bahrain", "Israel", "Egypt", "Morocco",
    "Algeria", "Tunisia", "Nigeria", "Kenya", "Ghana", "South Africa",
)
_MORE = (
    "Netherlands", "Belgium", "Switzerland", "Poland", "Turkey", "Ireland",
    "Sweden", "Denmark", "Austria", "Portugal", "Czech Republic", "Romania",
    "Greece", "Hungary", "Pakistan", "Bangladesh", "Sri Lanka", "Taiwan",
    "Hong Kong", "New Zealand", "Chile", "Colombia", "Peru", "Ecuador",
    "Jordan",
)
CAPS_REGIONS: tuple[str, ...] = _REPORT_COUNTRIES + _MORE

# Row 2 of the final Excel.
DEFINITIONS: dict[str, str] = {
    "caps_distributor": (
        "Global Empty Capsules Market: empty (unfilled) capsule shells used to "
        "make capsule dosage forms and supplements. "
        "By Capsule Type: Gelatin Capsules, Non-gelatin Capsules. "
        "By Material: Porcine, Bovine and Marine (Fish) Gelatin Capsules; "
        "Plant-based Capsules (HPMC, Pullulan, others such as Carrageenan); "
        "others (Polyvinyl Alcohol, etc.). "
        "By Functionality: Immediate-release, Delayed-release, "
        "Sustained-release, others (Targeted-release, etc.). "
        "By Shell Type: Hard, Soft. By Size: 000, 00, 0, 1, 2, others (3, 4, "
        "5, etc.). By Purchase Mode: Online, Offline. "
        "By Application: Oral Release, Specialized Delivery, Pre-clinical and "
        "Clinical Solutions, others (Enteric, etc.). "
        "By End User: Biopharmaceutical Companies, Nutraceutical Companies, "
        "Clinical Research Organizations (CROs), others (Cosmetic Companies, "
        "etc.). Geography: global - North America, Europe, Asia Pacific, Latin "
        "America, Middle East, Africa. "
        "COMPANY TYPE IN THIS LIST - EMPTY CAPSULE DISTRIBUTOR: a company that "
        "buys empty capsules from capsule manufacturers (e.g. Lonza/Capsugel, "
        "ACG, Qualicaps, Suheung, CapsCanada, Sunil Healthcare) and resells "
        "them to pharmaceutical, nutraceutical, CRO and cosmetic "
        "manufacturers - pharmaceutical excipient distributors, capsule "
        "importers and stockists, and B2B (online or offline) capsule "
        "suppliers - manufacturing none of the capsules it sells. "
        "EXCLUDED companies: capsule manufacturers, contract manufacturers / "
        "CDMOs that fill capsules, capsule-filling machine sellers, finished "
        "supplement or drug brands, and pharmacies or shops selling only to "
        "consumers."
    ),
}


def by_slug_caps(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_CAPS:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown caps market slug: {slug}")

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified company whose profile never matches this goes to
# flagged_for_review.xlsx instead of the main Excel. Tested on the delivered
# list 2026-10-21 (all genuine rows match).
EVIDENCE: dict[str, str] = {
    "caps_distributor": r"capsule|capsugel|lonza|\bacg\b|qualicaps|suheung|hpmc|gelatin|excipient",
}
