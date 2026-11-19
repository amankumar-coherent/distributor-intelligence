"""Global Electrosurgical and Allied Surgical Medical Devices Market -- ONE
market, worldwide, DISTRIBUTORS / DEALERS (user, 2026-10-30).

Inputs: user's definition, product-type and surgical-application segmentation,
18 exclusions, and "Report_Sample_Global_Electrosurgical And Allied Surgical
Medical Devices Market_Coherent_Market_Insights.pdf" (Downloads). The report's
company profiles are MAKERS (Medtronic, Erbe, Olympus, CONMED, BOWA, KARL STORZ,
Ethicon, Stryker, STERIS, KLS Martin, Richard Wolf, Utah Medical, Kirwan, FIAB,
LED SpA, Sutter, ALAN, Deltronix, Encision, MedGyn, Renuvion, TAKTVOLL...) and
its sales channel segment is "Direct Manufacturer Sales" vs "Distributor and
Dealer Sales" (~50-55% of global sales). Distributor intelligence = that second
channel: distributors, dealers and importers reselling these makers' ESUs,
electrosurgical instruments/accessories, CO2 insufflators and surgical smoke
evacuation systems. The user was asleep at setup; same choice the user made for
the biopharma buffer market (distributors only, makers excluded).

The user's exclusions are applied to COMPANIES: makers and their own sales
subsidiaries; distributors selling only manual instruments, ultrasonic / laser /
cryo / microwave energy, thermal cautery, ablation systems, endoscopy or
visualization, robots, trocars / access devices, medical gas, irrigation /
distension, general suction or air filtration, closure / hemostats, general OR
equipment, general consumables, sterilization, TENS / aesthetic devices;
service-only companies.

Global: the report's regions country by country (CAPS_REGIONS' 62 countries),
re-ordered so every round of 10 covers several continents -- if the 400 cap is
reached early the list is still global (buffer-market lesson, 2026-10-28).
Scope rides in ONE Google URL (budget 7,550 / hard 7,800); measure with
scripts/preflight_prompt.py --batch esu.
"""
from __future__ import annotations

from caps_markets import CAPS_REGIONS

MARKETS_ESU: tuple[tuple[str, str, str, str], ...] = (
    (
        "esu_distributor",
        "Global Electrosurgical and Allied Surgical Medical Devices Market",
        "PRODUCT SCOPE (required): electrosurgical generators (ESUs), "
        "electrosurgical pencils, electrodes and bipolar forceps, CO2 "
        "insufflators, surgical smoke evacuators, by Medtronic, Erbe, "
        "Olympus, CONMED, BOWA. DISTRIBUTORS or dealers of these to "
        "hospitals. EXCLUDE device makers, sellers of only manual "
        "instruments, laser, ultrasonic, ablation, endoscopes, OR equipment "
        "or consumables.",
        "distributor",
    ),
)

SIBLINGS_ESU: dict[str, str] = {}

# Rounds of 10 that each span several continents (all 62 CAPS_REGIONS).
ESU_REGIONS: tuple[str, ...] = (
    "United States", "Germany", "China", "India", "Japan", "Brazil",
    "United Arab Emirates", "Saudi Arabia", "South Africa", "United Kingdom",
    "France", "Italy", "Spain", "South Korea", "Australia", "Mexico",
    "Turkey", "Egypt", "Canada", "Russia",
    "Indonesia", "Thailand", "Vietnam", "Malaysia", "Philippines",
    "Singapore", "Argentina", "Chile", "Colombia", "Israel",
    "Poland", "Netherlands", "Belgium", "Switzerland", "Austria", "Sweden",
    "Nigeria", "Kenya", "Morocco", "Pakistan",
    "Taiwan", "Hong Kong", "New Zealand", "Peru", "Ecuador", "Qatar",
    "Kuwait", "Oman", "Bahrain", "Jordan",
    "Ireland", "Denmark", "Portugal", "Czech Republic", "Romania", "Greece",
    "Hungary", "Algeria", "Tunisia", "Ghana",
    "Bangladesh", "Sri Lanka",
)
assert sorted(ESU_REGIONS) == sorted(CAPS_REGIONS)

DEFINITIONS: dict[str, str] = {
    "esu_distributor": (
        "The Global Electrosurgical and Allied Surgical Medical Devices Market "
        "encompasses the commercial production, supply, and sale of "
        "electrosurgical units, disposable and reusable electrosurgical "
        "instruments and accessories, CO2 insufflators, and surgical smoke "
        "evacuation systems and accessories used in human and veterinary "
        "surgical procedures. Electrosurgical products use high-frequency "
        "electrical energy for tissue cutting, coagulation, and related "
        "surgical functions. Allied surgical devices support procedures by "
        "delivering controlled CO2 insufflation to establish and maintain the "
        "operative working space and by capturing and removing surgical smoke. "
        "The market includes standalone equipment, integrated systems, and "
        "associated accessories and consumables used across open and minimally "
        "invasive surgical procedures. "
        "BY PRODUCT TYPE: Electrosurgical Units (ESUs) and Accessories - "
        "Monopolar, Bipolar, and Combined Monopolar and Bipolar Generators; "
        "Disposable and Reusable Electrosurgical Instruments and Accessories; "
        "CO2 Insufflators - Conventional, Advanced, and Integrated CO2 "
        "Insufflation and Smoke Evacuation Systems; Surgical Smoke Evacuation "
        "Systems and Accessories - Standalone Systems, Integrated Modules, and "
        "Disposable and Reusable Accessories. BY SURGICAL APPLICATION: General "
        "and Abdominal Surgery, Gynecological Surgery, Gastrointestinal and "
        "Colorectal Surgery, Urological Surgery, Cardiovascular and Thoracic "
        "Surgery, Orthopedic Surgery, Neurosurgery, ENT Surgery, Plastic and "
        "Reconstructive Surgery, and Other Surgical Applications (Dermatology, "
        "Ophthalmology, and Other Specialties). "
        "EXCLUDED from the market: non-electrosurgical manual instruments "
        "(scalpels, scissors, forceps, clamps, retractors); alternative "
        "surgical energy devices (standalone ultrasonic, laser, cryosurgical "
        "and microwave systems); standalone thermal cautery devices; dedicated "
        "ablation systems (cardiac electrophysiology, tumor ablation, pain "
        "management); endoscopy and surgical visualization equipment; robotic "
        "surgical platforms (separately sold electrosurgical instruments stay "
        "included); general laparoscopic access devices (trocars, ports, "
        "cannulas, Veress needles, retrieval bags); medical CO2 and gas supply "
        "infrastructure; other insufflation and distension equipment (non-CO2 "
        "insufflation, liquid distension, irrigation, fluid management); "
        "general suction and air filtration equipment; surgical closure and "
        "hemostatic products; general operating room equipment; general "
        "surgical consumables and protective products; cleaning and "
        "sterilization products; non-surgical electrical therapy devices (TENS, "
        "muscle stimulators, physiotherapy, non-surgical aesthetic devices); "
        "standalone device-related services; surgical procedure and healthcare "
        "service revenue; intermediate transactions and duplicate system "
        "revenues. "
        "COMPANY TYPE IN THIS LIST - ELECTROSURGICAL DEVICE DISTRIBUTOR: a "
        "distributor, dealer or importer that buys electrosurgical units, "
        "electrosurgical instruments and accessories, CO2 insufflators or "
        "surgical smoke evacuation systems from the manufacturers (e.g. "
        "Medtronic, Erbe, Olympus, CONMED, BOWA, KARL STORZ, Stryker, Ethicon, "
        "Utah Medical, KLS Martin) and resells them to hospitals, ambulatory "
        "surgical centers, clinics and veterinary practices (the report's "
        "\"Distributor and Dealer Sales\" channel). EXCLUDED companies: device "
        "manufacturers and their own sales subsidiaries; distributors selling "
        "only excluded products (manual instruments, laser / ultrasonic / "
        "ablation systems, endoscopy, robots, access devices, medical gas, "
        "general OR equipment, consumables, sterilization, TENS or aesthetic "
        "devices); service-only companies."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; keep short).
VERIFY_RULES: dict[str, str] = {
    "esu_distributor": (
        "KEEP only distributors or dealers of ELECTROSURGICAL units or "
        "instruments, CO2 insufflators or smoke evacuators made by others.\n"
        "DROP device makers and their subsidiaries, sellers of only manual "
        "instruments, laser/ultrasonic/ablation, endoscopes, OR equipment, "
        "consumables, sterilization, service-only firms."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for).
EVIDENCE: dict[str, str] = {
    "esu_distributor": (
        r"electrosurg|electro-surg|diatherm|\besu\b|electrocaut|cauter|"
        r"radiofrequency surg|monopolar|bipolar|generator|insufflat|"
        r"smoke evac|smoke filt|valleylab|erbe|bowa|conmed|elektrochirurg|"
        r"elettrobistur|bisturi|bistouri|electrobistur"
    ),
}


def by_slug_esu(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_ESU:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown esu market slug: {slug}")
