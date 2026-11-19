"""Mozambique Fire Insurance Market -- ONE market, Mozambique HQ only,
fire / property INSURANCE SELLERS (user, 2026-10-28).

Inputs: user's definition, 7 key inclusions and 8 key exclusions, and "Sample_
Mozambique_Fire Insurance Market_Coherent_Market_Insights_061046.pptx"
(Downloads). The report profiles INSURERS (Empresa Mocambicana de Seguros,
Mocambique Companhia de Seguros, Indico, Imperial, Arko, Palma, Premium,
Confianca, Fidelidade, Global Alliance, Hollard, Britam, Diamond, Phoenix) and
lists the distribution channels insurer direct / brokers / agents /
bancassurance / affinity / institutional partnerships. User, asked insurers-only
vs insurers + intermediaries: "run on the part in my inclusion, make sure the
exclusions never come" -> insurers PLUS the brokers / agents / intermediaries of
inclusion 7 that sell fire or property cover in Mozambique.

Rule: SERVICE (the company sells/underwrites/places the insurance). EXCLUDED:
reinsurers, insurance consultants, loss adjusters / surveyors, fire-prevention
equipment and firefighting firms, companies selling ONLY motor / life / health /
marine / aviation / liability / crop / engineering lines, and companies outside
Mozambique. Banks are not listed unless they are themselves a licensed insurer
or broker.

Searched province by province, then main cities. Scope rides in ONE Google URL
(budget 7,550 / hard 7,800); measure with scripts/preflight_prompt.py --batch moz.
"""
from __future__ import annotations

MARKETS_MOZ: tuple[tuple[str, str, str, str], ...] = (
    (
        "moz_fire_insurance",
        "Mozambique Fire Insurance Market",
        "PRODUCT SCOPE (required): FIRE INSURANCE in Mozambique - basic and "
        "extended standalone fire policies and property multi-risk policies "
        "including fire, for buildings, machinery, stock and contents. The "
        "company SELLS or PLACES this cover (verdict service_provider): "
        "licensed insurers (Hollard, Britam, Fidelidade, Phoenix, Empresa "
        "Mocambicana de Seguros), insurance brokers and agents selling fire "
        "or property cover. EXCLUDE reinsurers, consultants, loss adjusters, "
        "fire-equipment and firefighting firms, and sellers of ONLY motor, "
        "life, health, marine, aviation, liability, crop or engineering "
        "cover.",
        "service",
    ),
)

SIBLINGS_MOZ: dict[str, str] = {}

# Provinces first (Maputo City, Sofala/Beira and Nampula carry most insurers),
# then main cities, so ten browsers cover the whole country in 3 rounds.
MOZ_REGIONS: tuple[str, ...] = (
    "Maputo City", "Maputo Province", "Sofala", "Nampula", "Tete", "Zambezia",
    "Manica", "Cabo Delgado", "Gaza", "Inhambane", "Niassa",
    "Matola", "Beira", "Pemba", "Quelimane", "Chimoio", "Nacala", "Xai-Xai",
    "Lichinga", "Inhambane city", "Maxixe", "Ressano Garcia",
    "Mocuba", "Angoche", "Montepuez", "Chokwe", "Dondo", "Gurue", "Moatize", "Cuamba", "Ilha de Mocambique", "Vilankulo", "Macia", "Marracuene", "Boane", "Namaacha", "Manhica", "Moamba", "Chibuto", "Mueda", "Mocimboa da Praia", "Marromeu", "Gorongosa", "Gondola", "Catandica", "Espungabera", "Milange", "Alto Molocue", "Songo", "Mandimba", "Massinga", "Maniamba", "Bilene", "Ponta do Ouro", "Macomia", "Ribaue", "Monapo", "Nhamatanda", "Mutarara", "Chiure", "Mopeia", "Quissico",
)

DEFINITIONS: dict[str, str] = {
    "moz_fire_insurance": (
        "The Mozambique Fire Insurance Market encompasses insurance policies "
        "designed to provide financial protection against physical loss or "
        "damage to buildings, properties, machinery, equipment, inventories, "
        "and other insured assets caused by fire and associated risks, "
        "including lightning and explosions. The market covers basic "
        "standalone fire insurance, extended standalone fire insurance, and "
        "property multi-risk insurance policies incorporating fire coverage. "
        "It includes residential, commercial, industrial, agricultural, "
        "institutional, infrastructure, and mixed-use properties insured by "
        "individuals, businesses, government entities, and other "
        "organizations across Mozambique. "
        "KEY INCLUSIONS: standalone fire insurance (basic and extended "
        "policies covering fire, lightning, explosions and other associated "
        "perils); property multi-risk insurance (residential, commercial and "
        "industrial policies incorporating fire protection); insured "
        "buildings and physical assets (residential buildings, commercial "
        "establishments, industrial facilities, warehouses, machinery, "
        "equipment, furniture, inventories); all policyholder categories "
        "(individuals and households, micro to large businesses, "
        "state-owned enterprises, government institutions, non-profit "
        "organizations); all property uses and end-user sectors (residential, "
        "agricultural, manufacturing, mining, oil and gas, energy, utilities, "
        "transportation, hospitality, healthcare, education, retail and "
        "others); fire-related additional coverage (smoke damage, "
        "firefighting damage, debris removal, fire-related business "
        "interruption); all specified distribution channels (insurer direct, "
        "insurance brokers, agents, bancassurance, affinity partnerships, "
        "institutional partnerships). "
        "KEY EXCLUSIONS: non-fire property insurance (standalone theft, "
        "burglary, flood, earthquake or storm policies without fire cover); "
        "other insurance categories (life, health, motor, marine, aviation, "
        "personal accident, workers' compensation, standalone liability); "
        "motor and cargo fire coverage provided only under motor, marine "
        "cargo or goods-in-transit policies; standalone business "
        "interruption insurance; specialized agricultural and engineering "
        "insurance (crop, livestock, construction all-risk, machinery "
        "breakdown) outside the fire / property scope; uninsured properties "
        "and losses; reinsurance and ancillary services (reinsurance "
        "transactions, fire prevention equipment, firefighting services, "
        "insurance consulting and other non-insurance activities); insurance "
        "outside Mozambique. "
        "COMPANY TYPE IN THIS LIST - FIRE / PROPERTY INSURANCE SELLER: a "
        "company headquartered in Mozambique that underwrites, sells or "
        "places fire or property multi-risk insurance - licensed insurers "
        "(insurer direct) and the insurance brokers, agents and similar "
        "intermediaries named in the distribution channels. EXCLUDED "
        "companies: reinsurers; insurance consultants, loss adjusters and "
        "surveyors; fire-prevention equipment and firefighting firms; "
        "companies selling only motor, life, health, marine, aviation, "
        "liability, crop or engineering insurance; non-insurance businesses; "
        "companies headquartered outside Mozambique."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; keep short).
VERIFY_RULES: dict[str, str] = {
    "moz_fire_insurance": (
        "KEEP only companies based in Mozambique that underwrite, sell or "
        "place FIRE or property multi-risk insurance (licensed insurers, "
        "brokers, agents).\n"
        "DROP reinsurers, consultants, loss adjusters, fire-equipment or "
        "firefighting firms, sellers of only motor/life/health/marine/"
        "liability/crop insurance, non-insurance firms."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for):
# insurance words in English and Portuguese plus fire / property terms. Test on
# the verified list before finalize.
EVIDENCE: dict[str, str] = {
    "moz_fire_insurance": (
        r"insur|seguro|seguradora|assurance|broker|corretor|mediador|"
        r"incêndio|incendio|\bfire\b|property|propriedade|patrimonial|multirisco|multi-risk"
    ),
}


def by_slug_moz(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_MOZ:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown moz market slug: {slug}")
