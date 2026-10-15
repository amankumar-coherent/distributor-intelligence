"""Global Antioxidants Market (rubber / polymer antioxidant additives).

Built from the user's brief and the CMI sample report
"Sample_Global_Antioxidants Market_Coherent_Market_Insights.pdf" (2026-10-13):
antioxidant type, rubber/polymer type, application, product form, performance
function, end-use industry, distribution channel (Direct / INDIRECT sales -
the indirect channel is what this list covers) and region.

Physical-goods distributor rule: companies that buy antioxidant additives in
volume from the producers, stock and often repack them, and resell to rubber
compounders, tire, glove, foam, adhesive, cable and plastics manufacturers.
The producers themselves (BASF, SONGWON, SI Group, ADEKA, LANXESS, Rianlon,
Dover, Chitec, Everlight, Double Bond, Mayzo, SABO, Fine Organic, HPL
Additives, Veeral, Everspring) are the quadrant vendors, not distributors.

Deliberately NOT in scope: food / beverage, supplement, pharmaceutical and
cosmetic antioxidants - a generic "antioxidants" search drifts straight into
those, so the scope says so twice (INCLUDE wording and EXCLUDE clause).
"""
from __future__ import annotations

MARKETS_CHEMICAL: tuple[tuple[str, str, str, str], ...] = (
    (
        "antioxidants",
        "Global Antioxidants Market",
        # Compact on purpose: discovery sends system prompt + this text in ONE
        # Google URL (budget 7,550 chars); see wound_markets.py for what
        # over-long scope text does (reply format cut off, prose, 0 companies).
        "PRODUCT SCOPE (required): RUBBER AND POLYMER ANTIOXIDANT ADDITIVES - "
        "antioxidants added to rubber, elastomers and plastics to stop "
        "oxidative degradation: aromatic amine (TMQ, 6PPD), hindered "
        "phenolic, phosphite, phosphonite and thioester grades, for tires, "
        "rubber goods, latex gloves, adhesives, plastic compounding and "
        "cable. Brands include BASF, SONGWON, SI Group, LANXESS. EXCLUDE "
        "food, supplement, pharmaceutical and cosmetic antioxidants, UV "
        "absorbers, light stabilisers and finished rubber or plastic goods. "
        "Name the grades and brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_CHEMICAL: dict[str, str] = {
    # The EXCLUDE clause already rules the food / nutraceutical market out;
    # repeating it as a sibling line would cost URL budget for nothing.
}

# Countries discovery rotates through: the sample report's own region list
# first (North America, Europe, Asia Pacific, Latin America, Middle East,
# Africa), then the other markets with real rubber-chemical distribution.
GLOBAL_REGIONS: tuple[str, ...] = (
    "China", "India", "United States", "Germany", "Taiwan", "South Korea",
    "Japan", "Italy", "United Kingdom", "Thailand", "Malaysia", "Indonesia",
    "Singapore", "Spain", "France", "Netherlands", "Belgium", "Turkey",
    "Brazil", "Mexico", "Poland", "Russia", "Canada", "Vietnam",
    "United Arab Emirates", "Saudi Arabia", "South Africa", "Egypt",
    "Argentina", "Colombia", "Chile", "Czech Republic", "Austria",
    "Switzerland", "Sweden", "Portugal", "Romania", "Hungary", "Greece",
    "Ukraine", "Israel", "Iran", "Pakistan", "Bangladesh", "Sri Lanka",
    "Philippines", "Australia", "New Zealand", "Hong Kong", "Algeria",
    "Morocco", "Nigeria", "Kenya", "Peru", "Ecuador", "Finland", "Denmark",
    "Norway", "Ireland", "Slovakia", "Bulgaria", "Serbia", "Croatia",
)


# Client-facing market definition (the user's own brief, 2026-10-13), shown
# in row 2 of the final Excel instead of a generic rule sentence.
DEFINITIONS: dict[str, str] = {
    "antioxidants": (
        "The Global Antioxidants Market covers chemical antioxidant additives used in "
        "rubber, elastomers, plastics, and polymer formulations to prevent oxidative "
        "degradation during processing and service life. The market includes aromatic "
        "amine, hindered phenolic, phosphite and phosphonite, thioester, and other "
        "antioxidants used across automotive, industrial, healthcare, construction, "
        "electrical, and consumer applications. Excludes food, supplement, "
        "pharmaceutical and cosmetic antioxidants, standalone antiozonants, UV absorbers "
        "and light stabilizers, finished rubber and plastic products, upstream "
        "intermediates, and masterbatches sold as finished compounds."
    ),
}


def by_slug_chemical(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_CHEMICAL:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown chemical market slug: {slug}")
