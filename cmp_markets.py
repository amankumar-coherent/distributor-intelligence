"""U.S. 503B Compounding Pharmacies Market -- ONE market, U.S. HQ only,
FDA-registered 503B OUTSOURCING FACILITIES (user, 2026-10-26).

Inputs: user's definition and five exclusions, and "Report_Sample_U.S._503b
Compounding Pharmacies Market_Coherent_Market_In....pdf" (Downloads). The
report's company profiles are 503B outsourcers (Fresenius Kabi, Nephron, B.
Braun / CAPS, SCA Pharma, Fagron Sterile Services, QuVa, Pentec Health,
AnazaoHealth, Empower, Olympia...), so the company TYPE here is the 503B
outsourcing facility itself, not a distributor. User: "dont want 503A", U.S.
HQ only, cap 150 after verify, 10 browsers.

Rule: SERVICE (outsourced compounding performed for hospitals, clinics and
surgery centers). The physical-goods rule would drop every 503B as a
"manufacturer"; the service rule keeps companies that PERFORM the work and
skips the is_distributor gate in verification.

EXCLUDED (user): 503A patient-specific / retail / community compounding
pharmacies; wholesalers, brokers, repackagers and other intermediaries;
conventional NDA/ANDA drug manufacturers; bulk API suppliers. A pharmacy group
that runs a 503A pharmacy AND an FDA-registered 503B facility is kept for its
503B facility (the report profiles Empower and Olympia); a 503A-only pharmacy
is dropped.

Searched state by state (ent_markets.ENT_REGIONS). Scope rides in ONE Google
URL (budget 7,550 / hard 7,800); measure with scripts/preflight_prompt.py.
"""
from __future__ import annotations

from ent_markets import ENT_REGIONS

MARKETS_CMP: tuple[tuple[str, str, str, str], ...] = (
    (
        "cmp_503b",
        "U.S. 503B Compounding Pharmacies Market",
        "PRODUCT SCOPE (required): FDA-REGISTERED 503B OUTSOURCING FACILITIES "
        "- they compound sterile and non-sterile medications in large batches "
        "(prefilled syringes, IV bags, ophthalmic, anesthesia and pain drugs) "
        "WITHOUT patient-specific prescriptions and sell them directly to "
        "hospitals, clinics and surgery centers for office use. Here the "
        "503B facility itself IS the service provider (verdict "
        "service_provider), e.g. Nephron 503B, SCA Pharma, QuVa, Fagron "
        "Sterile Services, Leiters, CAPS. EXCLUDE 503A-only patient-specific "
        "or retail compounding pharmacies, wholesalers, brokers, repackagers, "
        "conventional NDA/ANDA drug makers and bulk API suppliers. Name its "
        "503B facility and products.",
        "service",
    ),
)

SIBLINGS_CMP: dict[str, str] = {}

CMP_REGIONS: tuple[str, ...] = ENT_REGIONS

DEFINITIONS: dict[str, str] = {
    "cmp_503b": (
        "The U.S. 503B Compounding Pharmacies Market encompasses the "
        "commercial and institutional ecosystem of FDA-registered 503B "
        "outsourcing facilities that manufacture large batches of sterile "
        "and non-sterile compounded medications - often without a "
        "patient-specific prescription - to supply hospitals, clinics, "
        "surgery centers and healthcare systems for in-office administration "
        "and stock. "
        "By Molecule: Acetaminophen, Phenylephrine, Midazolam, Vancomycin, "
        "Epinephrine, Fentanyl, Bupivacaine, Morphine, Amiodarone, Heparin, "
        "Ketamine, Hydromorphone, Ketorolac, Ropivacaine, Clonidine, "
        "Mitomycin, Moxifloxacin, Fluconazole, Amphotericin, Gentamicin, "
        "Cyclosporine, Atropine Sulfate, Succinylcholine Chloride, Oxytocin, "
        "Famotidine, Hydralazine HCl, Glutathione, Hydroquinone, Salicylic "
        "Acid, Ascorbic Acid, Benzocaine, Tetracaine, Phenylephrine HCl, "
        "Spironolactone, Glutamine, Thiamine HCl, Niacinamide, Dexpanthenol, "
        "Pyridoxine HCl, Methionine, Cyanocobalamin, Aminophylline, "
        "Acyclovir, Azithromycin, Calcium Gluconate, Doxycycline, Hyaluronic "
        "Acid, Ibuprofen, Naltrexone, Ziconotide, Progesterone, Testosterone, "
        "Gabapentin, Sodium Chloride, Ondansetron, Sildenafil, Tadalafil, "
        "Minoxidil, Lorazepam, Propranolol Hydrochloride, Fluoxetine HCl, "
        "Finasteride, Baclofen, Cyclobenzaprine HCl, Cromolyn Sodium, "
        "Dexamethasone, Diltiazem HCl, Estradiol, Ketoprofen, Lidocaine, "
        "Levocetirizine, Phenol, Others (Dextrose and Others). "
        "By Therapeutic Area: Pain Management, Oncology Supportive Drugs, "
        "Ophthalmology, Cardiology, Orthopedics, Infectious Diseases, "
        "Gastroenterology, Hormone Replacement Therapy (HRT), Neurology, "
        "Dermatology, Others (e.g. endocrinology, respiratory). "
        "EXCLUDED from the market: 503A patient-specific compounding (retail "
        "/ community compounding requiring individual patient "
        "prescriptions); wholesaling and intermediaries (distribution via "
        "third-party wholesalers, brokers or repackagers - products must be "
        "sold directly to healthcare facilities); commercial copies "
        "(mass-producing identical copies of FDA-approved drugs unless the "
        "drug is on the FDA Drug Shortage List); unapproved bulk substances "
        "(bulk APIs not on the FDA 503B Bulks List); conventional "
        "manufacturing (FDA-approved mass pharma products under NDA / ANDA "
        "premarket approval). "
        "COMPANY TYPE IN THIS LIST - 503B OUTSOURCING FACILITY: a company "
        "headquartered in the United States that operates an FDA-registered "
        "503B outsourcing facility and sells its compounded medications "
        "directly to hospitals, clinics, surgery centers and health systems "
        "(including 503B units of larger pharma or pharmacy groups). "
        "EXCLUDED companies: 503A-only compounding pharmacies; drug "
        "wholesalers, distributors, brokers and repackagers; conventional "
        "NDA / ANDA drug manufacturers without a 503B facility; bulk API "
        "suppliers; companies headquartered outside the United States."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; keep short).
VERIFY_RULES: dict[str, str] = {
    "cmp_503b": (
        "KEEP only U.S.-based companies operating an FDA-registered 503B "
        "OUTSOURCING FACILITY that sells compounded drugs directly to "
        "hospitals, clinics or surgery centers (503B units of bigger groups "
        "count).\n"
        "DROP 503A-only or retail compounding pharmacies, wholesalers, "
        "brokers, repackagers, NDA/ANDA drug makers without a 503B, API "
        "suppliers."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for).
EVIDENCE: dict[str, str] = {
    # Profiles describe the products (sterile injectables, prefilled
    # syringes, office-stock vials), rarely the word "503B": a 503B-only
    # pattern flagged 126 of 143 verified companies (2026-10-26). Tested on
    # the verified list: this matches all of them.
    "cmp_503b": (
        r"503\s*-?\s*b|outsourc|compound|sterile|pre-?fill|syringe|admixture|"
        r"iv bag|injectab|ophthalmic|anesthe|office.?(stock|use)|vials?\b|formulation"
    ),
}


def by_slug_cmp(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_CMP:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown cmp market slug: {slug}")
