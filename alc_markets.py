"""India Alcohol Market -- ONE market, India-HQ DISTRIBUTORS and CHANNEL
PARTNERS of alcoholic beverages (user, 2026-11-19: CMI sample report
"Sample_India_Alcohol Market_Coherent_Market_Insights_.pdf" + the user's
definition / segmentation / exclusions). Emphasis on GIN, VODKA and WHISKY:
the client wants potential distributors and channel partners in India for
EXPORT expansion (a foreign brand entering India). Cap 400, 10 browsers,
through the Proton VPN guard.

Who counts: liquor importers, distributors, wholesalers, stockists, licensed
(L1 / FL / CL) trading companies, brand-building / route-to-market partners,
duty-free and travel-retail suppliers, CSD / defence-canteen suppliers, and
the state beverage corporations that act as the wholesale channel (KSBCL,
TASMAC, TSBCL, APSBCL, BEVCO...). Producers that only sell their own brands,
retail shops, bars / restaurants, delivery apps without a distribution licence
and raw-material / packaging / equipment suppliers are excluded.

Searched state by state (report states first), then cities. Discovery uses
SIMPLE QUERY mode (as sorb / bie); the verifier enforces the rules.
"""
from __future__ import annotations

from india_cities import INDIA_CITIES
from india_states import INDIA_STATES

SLUG = "alc_india_alcohol"
MARKET = "India Alcohol Market"

MARKETS_ALC: tuple[tuple[str, str, str, str], ...] = (
    (
        SLUG,
        MARKET,
        "PRODUCT SCOPE (required): ALCOHOLIC BEVERAGES in India - whisky, gin, "
        "vodka, rum, brandy, beer, wine, tequila and fortified wine, imported "
        "(BIO) and Indian-made (IMFL), sold to state beverage corporations, "
        "licensed liquor stores, hotels, restaurants, bars, clubs, duty free and "
        "the defence Canteen Stores Department (CSD). DISTRIBUTORS, IMPORTERS, "
        "wholesalers, stockists, licensed liquor trading companies and "
        "route-to-market channel partners of alcohol brands made by others, "
        "especially imported gin, vodka and whisky, for example Ambrosia, Empire "
        "Spirits, Geist, Talli Drinks. EXCLUDE distilleries, breweries and "
        "wineries selling only their own brands, retail liquor shops, bars and "
        "restaurants, and raw material, packaging or equipment suppliers.",
        "distributor",
    ),
)

SIBLINGS_ALC: dict[str, str] = {}

SIMPLE_REQUEST: dict[str, str] = {
    SLUG: ("liquor importers, distributors, wholesalers or channel partners of "
           "alcoholic beverages (especially imported gin, vodka and whisky, also "
           "rum, wine, beer) holding liquor distribution licences"),
}

_FIRST = ("Uttar Pradesh", "Maharashtra", "Karnataka", "Telangana", "Tamil Nadu",
          "Andhra Pradesh", "Kerala", "West Bengal", "Haryana", "Punjab", "Delhi", "Goa")


def _rank(region: str) -> int:
    for i, s in enumerate(_FIRST):
        if region.lower().startswith(s.lower()):
            return i
    return len(_FIRST)


ALC_REGIONS: tuple[str, ...] = (
    tuple(sorted(INDIA_STATES, key=_rank))
    + tuple(c for c in ("Goa (Panaji, Margao, Vasco da Gama)",) if not any(
        s.lower().startswith("goa") for s in INDIA_STATES))
    + tuple(INDIA_CITIES)
)

DEFINITIONS: dict[str, str] = {
    SLUG: (
        "The India Alcohol Market covers the production, distribution, and sales of "
        "alcoholic beverages across India, including beer, wine, whisky, rum, vodka, "
        "brandy, and other alcoholic beverages such as gin, tequila, and fortified "
        "wine. The market assessment includes demand by product type, alcoholic "
        "strength, packaging type, price point, application type, sales segment, and "
        "state. It also covers pricing dynamics, trade analysis, consumer behaviour, "
        "drinking preferences, distribution networks, and the competitive landscape. "
        "BY PRODUCT TYPE: Beer; Wine; Whisky; Rum; Vodka; Brandy; Others (Gin, "
        "Tequila, Fortified Wine). BY ALCOHOLIC STRENGTH: Low Alcohol (below 10); "
        "Medium Alcohol (10 to 20); High Alcohol (above 20). BY PACKAGING TYPE: "
        "Bottles (less than 200 ml; 200 ml to 750 ml; 750 ml to 1.5 L; above 1.5 L); "
        "Cans (150 ml to 330 ml; 330 ml to 500 ml; above 500 ml); Casks; Other "
        "Packaging Types (5 L to 10 L; 10 L to 50 L; above 50 L). BY PRICE POINT: "
        "Premium or Luxury; Mid Range; Economy or Value. BY APPLICATION TYPE: "
        "Household Consumption; Foodservice & Hospitality; Retail (Off Trade "
        "Consumption); Duty Free & Travel Retail; Event & Catering Services; Others "
        "(Online & E-Commerce Sales, Specialty Alcoholic Beverages). BY SALES "
        "SEGMENT: Civil Distribution Channel - Off Trade (Government Liquor Stores or "
        "Corporations; Private Licensed Liquor Stores; Wine & Beer Shops; Modern "
        "Retail, where permitted; Online or Home Delivery, where permitted) and On "
        "Trade (Hotels; Restaurants; Bars & Pubs; Clubs; Other Licensed On-Premise "
        "Outlets); Defense Distribution Channel (Canteen Stores Department (CSD); Unit "
        "Run Canteens (URCs); Other Authorized Defense Institutional Channels). BY "
        "STATE: Uttar Pradesh; Maharashtra; Karnataka; Telangana; Tamil Nadu; Andhra "
        "Pradesh; Kerala; West Bengal; Haryana; Punjab; Rest of State. "
        "EXCLUSIONS: non-alcoholic beverages and general beverage products outside "
        "the defined alcoholic beverage categories; products and services unrelated "
        "to the production, distribution and sale of alcoholic beverages; standalone "
        "suppliers of raw materials, packaging materials and equipment; "
        "alcohol-related services and businesses outside the India market scope. "
        "COMPANY TYPE IN THIS LIST - DISTRIBUTOR / CHANNEL PARTNER (focus: Gin, Vodka "
        "and Whisky, for export expansion into India): an India-headquartered "
        "importer, distributor, wholesaler, stockist, licensed liquor trading company "
        "or route-to-market partner that brings in, stores and sells alcoholic "
        "beverage brands made by others to state beverage corporations, licensed "
        "retail, on-trade (hotels, restaurants, bars, clubs), duty free / travel "
        "retail or the defence CSD channel; state beverage corporations acting as "
        "the wholesale channel are included. EXCLUDED: distilleries, breweries and "
        "wineries selling only their own brands, retail liquor shops, bars and "
        "restaurants, delivery apps without a distribution licence, and raw "
        "material, packaging and equipment suppliers; media, associations and "
        "directories."
    ),
}

VERIFY_RULES: dict[str, str] = {
    SLUG: (
        "KEEP India-HQ liquor importers, distributors, wholesalers, stockists, "
        "licensed liquor trading firms, route-to-market partners, CSD/defence or "
        "duty-free suppliers and state beverage corporations that sell alcohol "
        "brands made by others (whisky, gin, vodka, rum, wine, beer).\n"
        "DROP distilleries/breweries/wineries selling only own brands, retail "
        "liquor shops, bars/restaurants, delivery-only apps, raw material/"
        "packaging/equipment suppliers, non-alcoholic beverages; HQ outside India."
    ),
}

EVIDENCE: dict[str, str] = {
    SLUG: (
        r"alcohol|liquor|spirit|whisk|gin\b|vodka|rum\b|brandy|wine|beer|tequila|"
        r"imfl|\bbio\b|beverage|bar\b|duty free|canteen|\bcsd\b|import"
    ),
}

TARGETS: dict[str, int] = {SLUG: 400}


def by_slug_alc(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_ALC:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown alc market slug: {slug}")
