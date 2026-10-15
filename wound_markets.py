"""3 global wound-care / medical-film markets, distributors, worldwide.

Built from the user's three CMI report samples (Downloads, 2026-10-12):
  * Global Advanced Wound Care Market
  * Global Negative Pressure Wound Therapy (NPWT) Devices Market
  * Global Medical Films Market
Each scope below restates the report's own "Research Definition and Scope"
segmentation. The reports' "Distributor Intelligence" table is
Manufacturer | Distributor | Product | Contact/Presence, i.e. the deliverable
is third-party distributors of named manufacturers' products, so every
scope asks for the brands each distributor carries.

All three use the physical-goods distributor rule. Medical Films is a
materials market (medical-grade polymers and films sold to converters and
device makers), so its distributors are polymer / specialty-film
distributors, not wound-care dealers.

Each entry is (slug, market_name, product_scope, kind), the shape
market_batches.py expects.
"""
from __future__ import annotations

MARKETS_WOUND: tuple[tuple[str, str, str, str], ...] = (
    (
        "advanced_wound_care",
        "Global Advanced Wound Care Market",
        "PRODUCT SCOPE (required): ADVANCED WOUND CARE PRODUCTS - advanced "
        "wound dressings (bordered and non-bordered foam, alginate and fiber, "
        "hydrocolloid, hydrogel, collagen, antimicrobial and film dressings); "
        "wound therapy devices (negative pressure wound therapy systems, "
        "pressure relief and offloading devices, hyperbaric oxygen therapy "
        "equipment, electrical stimulation devices); biologics and "
        "regenerative wound care (skin substitutes, growth-factor "
        "products); wound closure and adjunctive products (adhesive "
        "closure, sealants, hemostats, compression); wound cleansing and "
        "debridement products. For chronic and acute wounds (diabetic "
        "foot, pressure and venous ulcers, burns, surgical wounds), "
        "supplied to hospitals, clinics and homecare. Manufacturers "
        "include Smith+Nephew, Solventum (3M), Molnlycke, B. Braun and "
        "Coloplast. EXCLUDE "
        "basic gauze, cotton and plain bandage traders and general "
        "pharmacy retailers. Name the advanced wound care brands it "
        "distributes.",
        "distributor",
    ),
    (
        "negative_pressure_wound_therapy",
        "Global Negative Pressure Wound Therapy (NPWT) Devices Market",
        "PRODUCT SCOPE (required): NEGATIVE PRESSURE WOUND THERAPY (NPWT) "
        "DEVICES - single-use (disposable, portable) NPWT devices and "
        "conventional NPWT pumps and systems, with their canisters, foam and "
        "gauze dressing kits; for diabetic foot ulcers, pressure ulcers, "
        "venous leg ulcers, burns, surgical and traumatic wounds; supplied to "
        "hospitals and clinics, ambulatory surgery centers and home care "
        "(including NPWT rental). EXCLUDE distributors of general wound "
        "dressings with no NPWT device line (that is the Advanced Wound Care "
        "market). Name the NPWT device brands it distributes, e.g. Solventum "
        "V.A.C., Smith+Nephew PICO, Molnlycke Avance.",
        "distributor",
    ),
    (
        "medical_films",
        "Global Medical Films Market",
        # Kept SHORT on purpose: the whole prompt rides in a Google URL, and
        # a scope this market's first draft used (~1,250 chars) pushed it to
        # 8,000 chars, over the 7,800 limit, which cut off the JSON reply
        # format and made AI Mode answer in prose (0 companies found).
        "PRODUCT SCOPE (required): MEDICAL-GRADE FILMS AND FILM POLYMERS - "
        "TPU / polyurethane, polyethylene, polypropylene, PET, EVA and "
        "silicone films and film-grade resins sold to medical film "
        "converters and device makers (wound dressing, catheter "
        "securement, surgical drape, NPWT drape, wearable patch, "
        "packaging films). Brands include Lubrizol, Covestro, BASF, "
        "Avient. EXCLUDE finished-dressing distributors and commodity "
        "plastics traders with no medical-grade line. Name the films "
        "or polymers and brands.",
        "distributor",
    ),
)

SIBLINGS_WOUND = {
    "advanced_wound_care": "negative pressure wound therapy devices",
    "negative_pressure_wound_therapy": "advanced wound care dressings",
    # medical_films: its EXCLUDE clause already rules out finished dressings.
}

# Discovery rotates through these when the worldwide question saturates:
# the report's own country list first, then the next-largest medical
# device and polymer markets.
GLOBAL_REGIONS: tuple[str, ...] = (
    "United States", "Germany", "United Kingdom", "France", "Italy", "Spain",
    "Canada", "China", "Japan", "India", "South Korea", "Australia",
    "Brazil", "Mexico", "Netherlands", "Belgium", "Switzerland", "Austria",
    "Poland", "Sweden", "Denmark", "Norway", "Finland", "Ireland",
    "Portugal", "Czech Republic", "Turkey", "Russia", "Saudi Arabia",
    "United Arab Emirates", "Israel", "South Africa", "Egypt", "Nigeria",
    "Kenya", "Argentina", "Chile", "Colombia", "Peru", "Singapore",
    "Malaysia", "Thailand", "Indonesia", "Vietnam", "Philippines",
    "Taiwan", "Hong Kong", "New Zealand", "Pakistan", "Bangladesh",
    "Greece", "Romania", "Hungary", "Qatar", "Kuwait", "Morocco",
)


def by_slug_wound(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_WOUND:
        if entry[0] == slug:
            return entry
    raise KeyError(
        f"unknown wound market slug: {slug}. "
        f"Valid: {', '.join(m[0] for m in MARKETS_WOUND)}"
    )
