"""India Spices And Seasonings Market -- ONE market, India-HQ only,
DISTRIBUTORS / WHOLESALERS (user, 2026-10-19).

Inputs: user's definition, segmentation and exclusions, and
"Sample_India_Spices And Seasonings Market_Coherent_Market_Insights.pdf"
(Downloads). Its channel segment defines the company type: "Distributors and
Dealers - intermediaries that purchase, stock, and distribute spices and
seasonings to retailers, institutions, and industrial customers" and
"Wholesale Markets - bulk trading channels supplying spices to retailers,
processors, foodservice operators". The report's manufacturers (MDH, Everest,
Badshah, MTR, Aachi...) and suppliers/exporters (Agrocel, Gautam Export,
Swani) are excluded; its "Distributor" examples Reliance Retail / Fabindia /
Urban Platter are retailers and excluded too (retail is its own channel).

User chose a FRESH run (the paused ProFood "spices_and_seasoning_products"
market is not reused). Searched state by state. Scope rides in ONE Google URL
(budget 7,550 / hard 7,800); measure with `python run_spice_market.py --slug
spice_distributor --region "Assam and the North East (Guwahati, Shillong,
Imphal)" --show-prompt`.
"""
from __future__ import annotations

MARKETS_SPICE: tuple[tuple[str, str, str, str], ...] = (
    (
        "spice_distributor",
        "India Spices And Seasonings Market",
        "PRODUCT SCOPE (required): SPICES AND SEASONINGS - whole, ground, "
        "crushed and blended spices (pepper, cardamom, chilli, turmeric, "
        "cumin, coriander, clove and others), masalas and seasoning blends "
        "for households, foodservice and food processors. Spice DISTRIBUTORS, "
        "dealers, stockists, super-stockists and bulk spice wholesalers or "
        "traders that buy from manufacturers and supply retailers, HoReCa "
        "and food processors; brands include MDH, Everest, Badshah, Catch, "
        "MTR, Aachi. EXCLUDE spice manufacturers and brand owners, retailers "
        "and e-commerce stores, commission agents, pure exporters, additive "
        "or flavour firms and equipment suppliers. Name the spice brands or "
        "lines it distributes.",
        "distributor",
    ),
)

SIBLINGS_SPICE: dict[str, str] = {}

DEFINITIONS: dict[str, str] = {
    "spice_distributor": (
        "The India Spices and Seasonings Market covers whole, ground, crushed, "
        "cracked and other processed spice forms, along with blended masalas, "
        "customised seasoning blends and value-added seasoning solutions "
        "supplied for household, foodservice and industrial food-processing "
        "applications. The market assessment covers product demand, pricing, "
        "sourcing, processing, distribution, quality and compliance "
        "requirements, and competitive developments across India. "
        "By Product Type: Black Pepper, Cardamom, Chilli, Cinnamon, Clove, "
        "Coriander, Cumin, Fennel, Fenugreek, Ginger, Nutmeg, Star Anise, "
        "Turmeric, Other Spices. By Product Form: Whole, Ground, Crushed, "
        "Cracked, Tea-Cut, Other. By Product Category: Conventional, Organic, "
        "Steam-Sterilised, Roasted, Gluten-Controlled, Spice Blends and "
        "Seasonings, Other. By Seasoning and Blend Application: Curry Powders "
        "and Regional Masala Blends, Snack Seasonings, Meat, Poultry and "
        "Seafood Seasonings, Ready-to-Eat and Convenience Food Seasonings, "
        "Sauces, Marinades and Condiment Seasonings, Foodservice Seasoning "
        "Blends, Other. By End-Use Industry: Household, Food Processing and "
        "Packaged Food, Snacks and Savoury, Meat, Poultry and Seafood "
        "Processing, Foodservice and HoReCa, Sauces, Condiments and Ready "
        "Meals, Other. By Sourcing Origin: Domestic, Imported, Domestic and "
        "Imported Blended. By Quality and Compliance: Conventional "
        "Specifications, Food Safety and Microbiological, Residue and "
        "Contaminant, Traceability and Origin Verification, Organic "
        "Certification, Customer-Specific, Other. By Distribution Channel: "
        "Direct Sales to Industrial Buyers, Distributors and Wholesalers, "
        "Retail and Modern Trade, Traditional Retail, E-Commerce, Foodservice "
        "Suppliers, Other. By Region: North, South, East and West India. "
        "EXCLUDED from the market: non-spice food ingredients unrelated to "
        "spices and seasonings; standalone food additives and flavouring "
        "ingredients outside the spice and seasoning scope; agricultural "
        "inputs and equipment used for spice cultivation or processing; "
        "finished food products in which spices and seasonings are only "
        "incidental ingredients; other products outside the defined spices "
        "and seasonings market. "
        "COMPANY TYPE IN THIS LIST - SPICE DISTRIBUTOR / WHOLESALER: an "
        "India-headquartered distributor, dealer, stockist, super-stockist or "
        "C&F agent, or bulk spice wholesaler / trader, that purchases, stocks "
        "and distributes spices, masalas and seasoning blends made by others "
        "to retailers, kirana and modern trade, institutions, HoReCa and "
        "industrial food processors. EXCLUDED companies: spice manufacturers, "
        "processors and brand owners (e.g. MDH, Everest, Badshah, MTR, Aachi, "
        "Catch, Tata Sampann, Ramdev, Sakthi), retailers, supermarkets, kirana "
        "stores and e-commerce sellers (e.g. Reliance Retail, Fabindia, Urban "
        "Platter), commission agents and brokers that do not stock goods, pure "
        "exporters, food additive and flavouring companies, agricultural input "
        "and spice-processing equipment suppliers, finished food companies, "
        "and companies headquartered outside India."
    ),
}


def by_slug_spice(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_SPICE:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown spice market slug: {slug}")

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified company whose profile never matches this goes to
# flagged_for_review.xlsx instead of the main Excel. Tested on the delivered
# list 2026-10-21 (all genuine rows match).
EVIDENCE: dict[str, str] = {
    "spice_distributor": r"spice|masala|seasoning|condiment|mdh|everest|catch|badshah|turmeric|chilli|pepper|cardamom|cumin|food|fmcg|grocer",
}
