"""Two GLOBAL chemical markets, DISTRIBUTORS (user, 2026-11-15):

  1. Global Ethanolamine Market
  2. Global Ethylene Oxide Derivatives Market

Inputs read: the two Coherent Market Insights pivots (segments + countries)
and "Section 6_Customer Intelligence.xlsx" (EO derivative END USERS by
industry - HUL, P&G, Henkel, L'Oreal, Syngenta, PPG, SLB ... - buyers, NOT
list entries; used to say which industries the distributors must serve).
The EO pivot's "Value" tab carries medical-film data (TPU wound-dressing
films) left over from another report; its "Master Sheet Volume" tab has the
real EO segmentation, used here.

Company type = DISTRIBUTOR (the pivots' "Indirect Sales" channel):
chemical distributors, traders, importers, stockists and authorised
resellers of these chemicals made by others. Producers (BASF, Dow, INEOS,
Nouryon, SABIC, Shell, Huntsman, Indorama, Reliance ...) are excluded.

One market at a time with 10 browsers (10 countries per round), every
pipeline step per market before the next (FINALIZE_EACH). Discovery uses
SIMPLE QUERY mode (Google AI Mode refuses the long rule-heavy prompt since
2026-11-13); the verifier enforces the company type and exclusions.
"""
from __future__ import annotations

from esu_markets import ESU_REGIONS

MARKETS_ETH: tuple[tuple[str, str, str, str], ...] = (
    (
        "eth01_ethanolamine",
        "Global Ethanolamine Market",
        "PRODUCT SCOPE (required): ETHANOLAMINES - monoethanolamine (MEA), "
        "diethanolamine (DEA), triethanolamine (TEA) and methyldiethanolamine "
        "(MDEA) in industrial, pharmaceutical, cosmetic and electronic grades, "
        "used for gas sweetening, surfactants and detergents, personal care, "
        "herbicides, cement grinding aids, metalworking fluids, textile "
        "chemicals, pharmaceuticals and water treatment. DISTRIBUTORS, "
        "traders, importers or stockists of ethanolamines made by others. "
        "EXCLUDE the producers of ethanolamines and their own sales arms.",
        "distributor",
    ),
    (
        "eth02_ethylene_oxide_derivatives",
        "Global Ethylene Oxide Derivatives Market",
        "PRODUCT SCOPE (required): ETHYLENE OXIDE DERIVATIVES - ethoxylates "
        "(alcohol, fatty acid, fatty amine, alkylphenol ethoxylates), "
        "ethylene glycols (MEG, DEG, TEG), polyethylene glycols (PEG), "
        "polyether polyols, ethylene oxide sterilants, polyethylene oxide "
        "(PEO) and ethylene carbonate, sold to home care, personal care, "
        "pharmaceutical, agrochemical, coatings, oil and gas, textile, "
        "construction, packaging, food, electronics, paper and water "
        "treatment industries. DISTRIBUTORS, traders, importers or stockists "
        "of these chemicals made by others. EXCLUDE the producers and their "
        "own sales arms.",
        "distributor",
    ),
)

SIBLINGS_ETH: dict[str, str] = {}

# Pivot countries first (both pivots list the same set), then the rest of
# the 62-country ESU order.
_PIVOT = ("United States", "China", "India", "Germany", "Japan", "South Korea",
          "Brazil", "Mexico", "United Kingdom", "France", "Italy", "Spain",
          "Russia", "Canada", "Indonesia", "Thailand", "Vietnam", "Malaysia",
          "Singapore", "Philippines", "Australia", "Saudi Arabia",
          "United Arab Emirates", "Kuwait", "Qatar", "Egypt", "Israel",
          "Turkey", "Argentina", "Colombia", "Chile", "South Africa",
          "Nigeria", "Kenya", "Morocco", "Algeria")
ETH_REGIONS: tuple[str, ...] = _PIVOT + tuple(r for r in ESU_REGIONS if r not in _PIVOT)
assert set(_PIVOT) <= set(ESU_REGIONS)

_CUSTOMERS = (
    "Customer industries (from the customer-intelligence file): home care "
    "(Hindustan Unilever, P&G, Henkel, Colgate-Palmolive, Reckitt), personal "
    "care and cosmetics (L'Oreal, Estee Lauder, Shiseido, Beiersdorf), "
    "agrochemicals (Syngenta, Bayer, Corteva, FMC, UPL), paints, coatings "
    "and inks (PPG, AkzoNobel, Asian Paints, Jotun, Sun Chemical), oil and "
    "gas (SLB, Halliburton, Baker Hughes), textile and leather, automotive, "
    "construction (Sika, Mapei, Saint-Gobain, Pidilite), chemical "
    "manufacturing (Ecolab, Solenis), packaging, food and beverage "
    "processing, pharmaceuticals and healthcare, electronics, pulp and paper "
    "and water treatment (Kemira, Kurita, Veolia).")

DEFINITIONS: dict[str, str] = {
    "eth01_ethanolamine": (
        "The Global Ethanolamine Market covers ethanolamines: "
        "BY PRODUCT TYPE: Monoethanolamine (MEA), Diethanolamine (DEA), "
        "Triethanolamine (TEA), Methyldiethanolamine (MDEA). BY GRADE: "
        "Industrial, Pharmaceutical, Cosmetic, Electronic. BY FUNCTION: "
        "Chemical Intermediate, pH Control Agent, Gas Treatment Agent, "
        "Surfactant Intermediate, Corrosion Inhibitor, Emulsifying Agent, "
        "Others. BY APPLICATION: Gas Sweetening, Surfactants & Detergents, "
        "Personal Care & Cosmetics, Herbicides & Agrochemicals, Cement "
        "Grinding Aids, Metalworking Fluids, Textile Chemicals, "
        "Pharmaceuticals, Water Treatment, Chemical Processing, Others. BY "
        "END-USER INDUSTRY: Oil & Gas, Home & Personal Care, Agriculture, "
        "Construction, Chemical Manufacturing, Pharmaceuticals, Textile, "
        "Water & Wastewater Treatment, Others. BY DISTRIBUTION CHANNEL: "
        "Direct Sales, Indirect Sales. BY REGION: North America, Europe, "
        "Asia Pacific, Latin America, Middle East, Africa. "
        "COMPANY TYPE IN THIS LIST - DISTRIBUTOR (indirect sales channel): a "
        "chemical distributor, trader, importer or stockist that buys "
        "ethanolamines from their producers and resells them to industrial "
        "users. EXCLUDED companies: ethanolamine producers and their own "
        "sales subsidiaries, and companies not selling ethanolamines."
    ),
    "eth02_ethylene_oxide_derivatives": (
        "The Global Ethylene Oxide Derivatives Market covers chemicals made "
        "from ethylene oxide: BY PRODUCT TYPE: Ethoxylates (Alcohol, Fatty "
        "Acid, Fatty Amine, Alkylphenol and Other Ethoxylates), Ethylene "
        "Glycols, Polyethylene Glycols (PEGs), Polyether Polyols, Ethylene "
        "Oxide Sterilants, Polyethylene Oxide (PEO), Ethylene Carbonate, "
        "Other Ethylene Oxide Derivatives. BY FUNCTION: Surfactants, "
        "Solvents, Chemical Intermediates, Humectants, Emulsifiers, "
        "Dispersants, Wetting Agents, Defoamers, Plasticizers, Corrosion "
        "Inhibitors, Gas Treatment Chemicals, Antifreeze & Coolants, "
        "Sterilizing Agents, Others. BY END-USE INDUSTRY: Home Care, "
        "Personal Care & Cosmetics, Pharmaceuticals & Healthcare, "
        "Agrochemicals, Paints Coatings & Inks, Oil & Gas, Textile & "
        "Leather, Automotive, Construction, Chemical Manufacturing, "
        "Packaging, Food & Beverage Processing, Electronics, Pulp & Paper, "
        "Water Treatment, Others. BY FEEDSTOCK SOURCE: Ethylene-Based, "
        "Bio-Ethylene Based. BY SALES CHANNEL: Direct Sales, Indirect Sales. "
        "BY REGION: North America, Europe, Asia Pacific, Latin America, "
        "Middle East, Africa. " + _CUSTOMERS + " "
        "COMPANY TYPE IN THIS LIST - DISTRIBUTOR (indirect sales channel): a "
        "chemical distributor, trader, importer or stockist that buys "
        "ethylene oxide derivatives from their producers and resells them to "
        "the customer industries above. EXCLUDED companies: EO derivative "
        "producers and their own sales subsidiaries, the end-user customers "
        "themselves, and companies not selling EO derivatives."
    ),
}

# Verifier KEEP/DROP (criteria block must stay under ~1,000 chars).
VERIFY_RULES: dict[str, str] = {
    "eth01_ethanolamine": (
        "KEEP chemical distributors, traders, importers or stockists that "
        "resell ethanolamines (MEA, DEA, TEA, MDEA) made by others.\n"
        "DROP ethanolamine producers and their sales arms, end users, and "
        "firms not selling ethanolamines."
    ),
    "eth02_ethylene_oxide_derivatives": (
        "KEEP chemical distributors, traders, importers or stockists that "
        "resell EO derivatives (ethoxylates, MEG/DEG/TEG glycols, PEG, "
        "polyether polyols, PEO, ethylene carbonate) made by others.\n"
        "DROP producers and their sales arms, end users, firms not selling "
        "EO derivatives."
    ),
}

EVIDENCE: dict[str, str] = {
    "eth01_ethanolamine": (
        r"ethanolamine|\bmea\b|\bdea\b|\btea\b|\bmdea\b|amine|"
        r"chemical|solvent|surfactant|gas treat"
    ),
    "eth02_ethylene_oxide_derivatives": (
        r"ethoxylat|glycol|\bmeg\b|\bdeg\b|\bteg\b|\bpeg\b|polyol|"
        r"polyether|ethylene oxide|ethylene carbonate|surfactant|chemical|solvent"
    ),
}

# SIMPLE QUERY phrasing (discovery.set_simple_query).
SIMPLE_REQUEST: dict[str, str] = {
    "eth01_ethanolamine": (
        "chemical distributors, traders or importers of ethanolamines "
        "(monoethanolamine MEA, diethanolamine DEA, triethanolamine TEA, MDEA)"),
    "eth02_ethylene_oxide_derivatives": (
        "chemical distributors, traders or importers of ethylene oxide "
        "derivatives (ethoxylates, ethylene glycols MEG/DEG, polyethylene "
        "glycols PEG, polyether polyols)"),
}


def by_slug_eth(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_ETH:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown eth market slug: {slug}")
