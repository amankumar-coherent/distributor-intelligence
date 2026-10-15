"""Global Lactoferrin Market (lactoferrin ingredient distribution).

From the user's brief (2026-10-13): bovine / human / recombinant and other
lactoferrin, an iron-binding milk glycoprotein, sold as powder, liquid,
capsules or tablets in standard (70-90%), high (90-98%) and ultra-high
(>=98%) purity, for infant formula, dietary supplements, functional food and
drink, pharmaceutical, animal nutrition and cosmetic use. The brief's
"Indirect Sales via Distributors" channel is what this list covers.

THE SCOPE RISK here is drift, and the brief's exclusions say so explicitly:
a search for "lactoferrin" pulls in general dairy and whey / casein
ingredient traders, and finished supplement or infant-formula brands. A
company qualifies only if LACTOFERRIN ITSELF is a product it distributes -
not milk powder, not whey protein, not a finished capsule brand, and not
dairy or purification equipment. Also excluded by the brief: research-stage
or laboratory work that never results in commercially supplied lactoferrin.

Physical-goods distributor rule: stocking distributors of the ingredient,
reselling to formulators and manufacturers. Producers (Fonterra/NZMP,
FrieslandCampina, Morinaga Milk, Tatua, Synlait, Ingredia, Glanbia, Bega,
Westland, Armor Proteines, Hilmar, Milei) are the vendors, not distributors.
"""
from __future__ import annotations

MARKETS_INGREDIENT: tuple[tuple[str, str, str, str], ...] = (
    (
        "lactoferrin",
        "Global Lactoferrin Market",
        # Compact on purpose: system prompt + this text ride in ONE Google URL
        # (budget 7,550 chars). See wound_markets.py for what over-long scope
        # text does: reply format cut off, prose answer, 0 companies.
        "PRODUCT SCOPE (required): LACTOFERRIN INGREDIENT - bovine, human and "
        "recombinant lactoferrin, an iron-binding milk protein, as powder, "
        "liquid, capsules or tablets, standard to ultra-high purity, for "
        "infant formula, supplements, functional food and drink, pharma, "
        "animal nutrition and cosmetics. Brands include Fonterra NZMP, "
        "FrieslandCampina, Morinaga, Tatua. The company must actually sell "
        "LACTOFERRIN. EXCLUDE dairy, whey and casein traders with no "
        "lactoferrin line, finished products without it, and processing "
        "equipment. Name the lactoferrin grades and brands.",
        "distributor",
    ),
)

SIBLINGS_INGREDIENT: dict[str, str] = {
    # The EXCLUDE clause already rules out whey / casein / dairy traders;
    # a sibling line would repeat it and cost URL budget.
}

# Countries discovery rotates through: nutraceutical and dairy-ingredient
# distribution hubs first (China and East Asia buy most infant-formula
# lactoferrin), then the producing countries and the rest.
GLOBAL_REGIONS: tuple[str, ...] = (
    "China", "United States", "India", "Japan", "South Korea", "Germany",
    "Netherlands", "Australia", "New Zealand", "France", "United Kingdom",
    "Italy", "Spain", "Vietnam", "Indonesia", "Thailand", "Malaysia",
    "Singapore", "Taiwan", "Brazil", "Mexico", "Canada", "Poland", "Turkey",
    "Belgium", "Switzerland", "Denmark", "Ireland", "Sweden", "Austria",
    "Russia", "United Arab Emirates", "Saudi Arabia", "Egypt",
    "South Africa", "Argentina", "Chile", "Colombia", "Peru", "Philippines",
    "Pakistan", "Bangladesh", "Sri Lanka", "Israel", "Greece", "Portugal",
    "Czech Republic", "Hungary", "Romania", "Norway", "Finland", "Ukraine",
    "Morocco", "Nigeria", "Kenya", "Hong Kong", "Qatar", "Kuwait",
)


# Client-facing market definition (the user's own brief, 2026-10-13), shown
# in row 2 of the final Excel instead of a generic rule sentence.
DEFINITIONS: dict[str, str] = {
    "lactoferrin": (
        "The Global Lactoferrin Market covers lactoferrin ingredients and products "
        "derived from bovine, human, recombinant and other sources, used across "
        "nutritional, pharmaceutical, food and beverage, animal nutrition, and personal "
        "care applications. Lactoferrin is an iron-binding glycoprotein valued for its "
        "immune-supporting, antimicrobial and nutritional properties, supplied in "
        "powder, liquid, capsules, tablets and other forms, across standard, high-purity "
        "and ultra-high-purity grades, for infant formula, dietary supplements, "
        "functional foods and beverages, medical nutrition, pharmaceutical formulations, "
        "animal nutrition, and cosmetics and personal care. Excludes general dairy and "
        "milk ingredients, whey protein, casein and lactalbumin sold independently, "
        "products that do not contain lactoferrin, dairy processing and purification "
        "equipment, and research-stage work with no commercially supplied lactoferrin."
    ),
}


def by_slug_ingredient(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_INGREDIENT:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown ingredient market slug: {slug}")
