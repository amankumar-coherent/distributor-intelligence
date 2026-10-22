"""MENA Magnesium Glycinate Market -- ONE market, MENA countries only,
DISTRIBUTORS (user, 2026-10-18).

Inputs: user's definition / segmentation / exclusions / key brands, and
"Sample_MENA_Magnesium Glycinate Market_Coherent_Market_Insights.pdf"
(Downloads). The report's company profile separates Finished Product
Manufacturers, Specialty Nutraceutical Distributors (Ibnsina Pharma, Tamer
Group, Banaja, Salehiya, Fattal, Cigalah, Dallah, Kettaneh, Mersaco, Promedic,
Ubipharm Morocco, AMCO), Brand Owners and Pharmacies / Retailers /
E-commerce -- the DISTRIBUTOR group is what this batch collects.

Scope rides in ONE Google URL with the system prompt (budget 7,550 / hard
7,800); measure with `python run_mg_market.py --slug mg_distributor --region
"United Arab Emirates" --show-prompt`.
"""
from __future__ import annotations

MARKETS_MG: tuple[tuple[str, str, str, str], ...] = (
    (
        "mg_distributor",
        "MENA Magnesium Glycinate Market",
        "PRODUCT SCOPE (required): MAGNESIUM GLYCINATE SUPPLEMENTS - finished "
        "supplements containing magnesium glycinate (bisglycinate): powder, "
        "capsules, tablets, softgels, gummies, liquids and sachets, single "
        "and combination formulas for sleep, stress, muscle, bone, women's "
        "health and sports recovery. Brands include NOW Foods, Doctor's Best, "
        "Thorne, Solgar, Nature Made, KAL, Pure Encapsulations, Life "
        "Extension, Swanson, Nature's Bounty, Garden of Life. EXCLUDE brand "
        "owners and manufacturers, raw-material and ingredient suppliers, "
        "pharmacy chains, retailers and e-commerce stores. Name the "
        "supplement brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_MG: dict[str, str] = {}

# The user's regions, in their order (GCC, Levant, North Africa groups).
MG_COUNTRIES: tuple[str, ...] = (
    "Saudi Arabia", "United Arab Emirates", "Kuwait", "Qatar", "Oman",
    "Bahrain", "Egypt", "Morocco", "Algeria", "Libya", "Sudan", "Tunisia",
    "Jordan", "Lebanon", "Iraq", "Yemen",
)
MG_REGIONS: tuple[str, ...] = MG_COUNTRIES

DEFINITIONS: dict[str, str] = {
    "mg_distributor": (
        "The MENA Magnesium Glycinate Market covers magnesium glycinate "
        "supplements and finished formulations containing magnesium glycinate "
        "as an active nutritional ingredient. The market includes powder, "
        "capsules, tablets, softgels, gummies, oral solutions, liquids, and "
        "sachets, along with single-ingredient and combination formulations "
        "used for magnesium supplementation, general wellness, sleep and "
        "relaxation, stress management, muscle health, women's health, sports "
        "recovery, and other nutritional applications across the Middle East "
        "and North Africa. "
        "By Product Type: Powder, Capsules, Tablets, Softgels, Gummies, Oral "
        "Solutions & Liquids, Sachets, Other. "
        "By Formulation Type: Single-Ingredient; + Vitamins; + Minerals; + "
        "Herbal/Botanical; + Amino Acids; Multi-Magnesium; Other Combinations. "
        "By Application: Magnesium Deficiency, Bone & Muscle Health, Sleep & "
        "Relaxation, Stress & Mood, Cardiovascular, Energy & Metabolism, "
        "Nervous System, Women's Health, Sports Performance & Recovery, "
        "Healthy Aging, General Wellness, Other. "
        "By End User: General Adults, Elderly, Women, Athletes, "
        "Magnesium-Deficient Individuals, Sleep & Stress Consumers, Patients, "
        "Health-Conscious Consumers, Other. By Dosage: Low, Standard, High. "
        "By Distribution Channel: Retail Pharmacies, Hospital & Institutional "
        "Pharmacies, Clinics, E-Commerce, Health & Wellness Stores, "
        "Supermarkets & Hypermarkets, Specialty Nutrition Stores, "
        "Direct-to-Consumer, Other. "
        "By Region: GCC (Saudi Arabia, UAE, Kuwait, Qatar, Oman, Bahrain); "
        "Levant (Egypt, Morocco); North Africa (Algeria, Libya, Sudan, "
        "Tunisia, Jordan, Lebanon, Iraq, Yemen). "
        "EXCLUDED from the market: magnesium supplements without magnesium "
        "glycinate (except for benchmarking); other magnesium salts sold "
        "standalone; general vitamins, minerals and supplements unrelated to "
        "magnesium glycinate; food, beverage and cosmetic products; "
        "pharmaceuticals without magnesium glycinate; upstream raw materials "
        "and ingredient intermediates; other finished supplements without "
        "magnesium glycinate. "
        "COMPANY TYPE IN THIS LIST - SPECIALTY NUTRACEUTICAL DISTRIBUTOR: a "
        "pharmaceutical / nutraceutical distributor, importer or wholesaler "
        "headquartered in the MENA countries above that distributes finished "
        "dietary-supplement brands including magnesium glycinate products "
        "(e.g. NOW Foods, Doctor's Best, Thorne, Solgar, Nature Made, Pure "
        "Encapsulations, Life Extension, Nature's Bounty, Garden of Life) to "
        "pharmacies, hospitals, clinics and retailers - as in the report's "
        "distributor group (Ibnsina Pharma, Tamer Group, Banaja, Salehiya, "
        "Fattal, Cigalah, Dallah, Kettaneh, Mersaco, Promedic, Ubipharm, "
        "AMCO). EXCLUDED companies: supplement brand owners and finished "
        "product manufacturers, raw-material and ingredient suppliers, "
        "pharmacy chains, retailers and e-commerce stores, and companies "
        "headquartered outside these MENA countries."
    ),
}


def by_slug_mg(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_MG:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown mg market slug: {slug}")

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified company whose profile never matches this goes to
# flagged_for_review.xlsx instead of the main Excel. Tested on the delivered
# list 2026-10-21 (all genuine rows match).
EVIDENCE: dict[str, str] = {
    "mg_distributor": r"supplement|nutraceutical|vitamin|magnesium|nutrition|wellness|now foods|solgar|doctor'?s best|thorne|pharma",
}
