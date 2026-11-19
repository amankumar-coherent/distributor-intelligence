"""Potassium Sorbate, Sorbic Acid and Caffeine -- each as a GLOBAL market and
a separate INDIA market (user, 2026-11-16: "Global module + India module,
separate analysis for both"). DISTRIBUTORS + WHOLESALERS (traders,
importers, stockists) of these chemicals made by others; producers excluded.
Cap 400 each, one market at a time, 10 browsers per step, through the
Proton VPN guard (scripts/vpn_guard.py).

Global markets: country by country (pivot-style order, then the rest of the
62-country list). India markets: India-HQ only (set_country_restriction),
26 states then the industrial cities.

Discovery uses SIMPLE QUERY mode (Google AI Mode refuses the long prompt);
the verifier enforces company type and exclusions (VERIFY_RULES).
"""
from __future__ import annotations

from esu_markets import ESU_REGIONS
from india_cities import INDIA_CITIES
from india_states import INDIA_STATES

_PRODUCTS = {
    "potassium_sorbate": dict(
        name="Potassium Sorbate",
        scope=("POTASSIUM SORBATE (E202) - granular, powder and solution grades, food, "
               "beverage, cosmetic, pharmaceutical and feed grade preservative"),
        uses=("food and beverages (bakery, dairy, cheese, wine, juices, sauces), cosmetics "
              "and personal care, pharmaceuticals and animal feed"),
        makers="Celanese, Nantong Acetic Acid Chemical, Wanglong, Ningbo Wanglong, Hubei Kexing",
        evidence=r"sorbate|sorbic|e202|preservative|food ingredient|food additive|chemical",
    ),
    "sorbic_acid": dict(
        name="Sorbic Acid",
        scope=("SORBIC ACID (E200) - powder and granular grades, food, feed, cosmetic and "
               "pharmaceutical grade preservative"),
        uses=("food and beverages, bakery, dairy, animal feed and silage, cosmetics and "
              "pharmaceuticals"),
        makers="Celanese, Nantong Acetic Acid Chemical, Wanglong, Daicel",
        evidence=r"sorbic|sorbate|e200|preservative|food ingredient|food additive|chemical",
    ),
    "caffeine": dict(
        name="Caffeine",
        scope=("CAFFEINE - anhydrous caffeine (synthetic and natural), caffeine citrate and "
               "natural caffeine extracts, food, beverage, pharmaceutical and nutraceutical grade"),
        uses=("energy drinks and soft drinks, pharmaceuticals (analgesics, cold and flu), "
              "dietary supplements and sports nutrition, cosmetics"),
        makers="CSPC Pharmaceutical, Shandong Xinhua Pharmaceutical, Kudos Chemie, Aarti Healthcare, BASF",
        evidence=r"caffeine|anhydrous|nutraceutical|pharmaceutical|food ingredient|beverage|chemical|api",
    ),
}

_ORDER = ("potassium_sorbate", "sorbic_acid", "caffeine")

MARKETS_SORB: tuple[tuple[str, str, str, str], ...] = ()
DEFINITIONS: dict[str, str] = {}
VERIFY_RULES: dict[str, str] = {}
EVIDENCE: dict[str, str] = {}
SIMPLE_REQUEST: dict[str, str] = {}
INDIA_SLUGS: set[str] = set()

_markets = []
for i, key in enumerate(_ORDER):
    p = _PRODUCTS[key]
    for j, (geo, label) in enumerate((("global", "Global"), ("india", "India"))):
        slug = f"sorb{i * 2 + j + 1:02d}_{geo}_{key}"
        name = f"{label} {p['name']} Market"
        where = "in India (India-headquartered companies only)" if geo == "india" else "worldwide"
        scope = (f"PRODUCT SCOPE (required): {p['scope']}, sold {where} to manufacturers in "
                 f"{p['uses']}. DISTRIBUTORS, WHOLESALERS, traders, importers or stockists of "
                 f"{p['name'].lower()} made by others, for example buying from producers such as "
                 f"{p['makers']}. EXCLUDE the producers and their own sales arms, and companies "
                 f"selling only finished consumer products.")
        _markets.append((slug, name, scope, "distributor"))
        DEFINITIONS[slug] = (
            f"The {name} covers sales of {p['scope'].split(' - ')[0].lower()} ({p['scope'].split(' - ', 1)[1]}) "
            f"used in {p['uses']}. "
            + ("GEOGRAPHY: India - companies headquartered in India. " if geo == "india"
               else "GEOGRAPHY: global - companies headquartered in any country. ")
            + "COMPANY TYPE IN THIS LIST - DISTRIBUTOR / WHOLESALER: a distributor, wholesaler, "
              f"trader, importer or stockist that buys {p['name'].lower()} from its producers and "
              "resells it to food, beverage, pharmaceutical, cosmetic and feed manufacturers. "
              f"EXCLUDED: {p['name'].lower()} producers and their own sales subsidiaries, and "
              "companies selling only finished consumer products."
        )
        VERIFY_RULES[slug] = (
            f"KEEP distributors, wholesalers, traders or importers that resell {p['name'].lower()} "
            f"made by others" + (" and are headquartered in India" if geo == "india" else "") + ".\n"
            f"DROP {p['name'].lower()} producers and their sales arms, finished-product brands, "
            "and firms not selling it."
        )
        EVIDENCE[slug] = p["evidence"]
        SIMPLE_REQUEST[slug] = (f"distributors, wholesalers, traders or importers of "
                                f"{p['name'].lower()} used in food, beverages, pharmaceuticals "
                                "and cosmetics")
        if geo == "india":
            INDIA_SLUGS.add(slug)
MARKETS_SORB = tuple(_markets)

SIBLINGS_SORB: dict[str, str] = {}

# Per-market cap (user 2026-11-16 10:0x: "in India market keep only 200, in
# global 400"). The launcher reads TARGETS before its --target default.
TARGETS: dict[str, int] = {slug: (200 if slug in INDIA_SLUGS else 400) for slug, *_ in MARKETS_SORB}

_GLOBAL_FIRST = ("China", "India", "United States", "Germany", "Netherlands", "Japan",
                 "South Korea", "Brazil", "Mexico", "United Kingdom", "France", "Italy",
                 "Spain", "Turkey", "Indonesia", "Thailand", "Vietnam", "Malaysia",
                 "United Arab Emirates", "Saudi Arabia", "Egypt", "South Africa",
                 "Australia", "Canada", "Russia", "Poland", "Argentina", "Colombia")
_GLOBAL_REGIONS = _GLOBAL_FIRST + tuple(r for r in ESU_REGIONS if r not in _GLOBAL_FIRST)
_INDIA_REGIONS = tuple(INDIA_STATES) + tuple(INDIA_CITIES)

# Per-market search plan (the launcher accepts {slug: regions}).
SORB_REGIONS: dict[str, tuple[str, ...]] = {
    slug: (_INDIA_REGIONS if slug in INDIA_SLUGS else _GLOBAL_REGIONS) for slug, *_ in MARKETS_SORB
}


def by_slug_sorb(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_SORB:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown sorb market slug: {slug}")
