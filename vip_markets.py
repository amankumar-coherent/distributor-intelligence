"""North America and Europe Vacuum Insulation Panel Market -- ONE market,
DISTRIBUTORS only, companies headquartered in North America or Europe
(user, 2026-10-25).

No sample report; definition, inclusion/exclusion table and segmentation are
the user's. Company type (user confirmed "Distributors only"): distributors,
stockists, importers and resellers of VIPs and VIP-integrated systems made by
others -- building-insulation, cold-chain / thermal-packaging, industrial and
technical insulation distributors. VIP manufacturers (va-Q-tec, Kingspan
OPTIM-R, Panasonic U-Vacua, Promat/Etex, Porextherm, Kevothermal, ThermoCor...)
and VIP-integrated system makers are EXCLUDED, as are conventional-insulation-
only sellers, core-material / barrier-film suppliers, appliance and vehicle
makers, refrigerant / PCM / cooling-unit sellers, logistics services and
installers.

Uses the 2026-10-21 pipeline standard: country gate (VIP_COUNTRIES), verifier
rule (VERIFY_RULES), product-evidence gate (EVIDENCE). Scope rides in ONE
Google URL (budget 7,550 / hard 7,800); measure with `python run_vip_market.py
--slug vip_distributor --region "Czech Republic" --show-prompt`.
"""
from __future__ import annotations

MARKETS_VIP: tuple[tuple[str, str, str, str], ...] = (
    (
        "vip_distributor",
        "North America and Europe Vacuum Insulation Panel Market",
        "PRODUCT SCOPE (required): VACUUM INSULATION PANELS (VIPs) and "
        "VIP-based systems for buildings, refrigeration, cold-chain packaging "
        "and industry; makers include va-Q-tec, Kingspan OPTIM-R, Panasonic, "
        "Porextherm. DISTRIBUTORS, stockists and resellers of these VIPs. "
        "EXCLUDE VIP makers, conventional-insulation-only sellers (mineral "
        "wool, EPS, PIR, foam), core suppliers and installers. Name the VIP "
        "brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_VIP: dict[str, str] = {}

# User's regions first (U.S., Canada, Germany, U.K., France, Spain, Italy,
# Russia), then the rest of Europe.
VIP_COUNTRIES: tuple[str, ...] = (
    "United States", "Canada", "Germany", "United Kingdom", "France", "Spain",
    "Italy", "Russia", "Netherlands", "Belgium", "Switzerland", "Austria",
    "Poland", "Sweden", "Denmark", "Norway", "Finland", "Ireland", "Portugal",
    "Czech Republic", "Hungary", "Romania", "Greece", "Slovakia", "Slovenia",
    "Croatia", "Bulgaria", "Serbia", "Lithuania", "Latvia", "Estonia",
    "Luxembourg", "Ukraine", "Iceland",
)
VIP_REGIONS: tuple[str, ...] = VIP_COUNTRIES

DEFINITIONS: dict[str, str] = {
    "vip_distributor": (
        "The North America and Europe Vacuum Insulation Panel (VIP) Market "
        "encompasses the manufacturing, commercial supply, and sale of "
        "high-performance thermal insulation panels that use an evacuated, "
        "gas-tight enclosure surrounding a low-conductivity core material to "
        "minimize heat transfer. Vacuum insulation panels provide high thermal "
        "resistance at comparatively low insulation thicknesses, making them "
        "suitable for applications where space efficiency, thermal "
        "performance, and temperature stability are important. A typical "
        "vacuum insulation panel consists of a porous or microporous core "
        "material, a gas-tight barrier envelope, and an evacuated internal "
        "structure; additional protective layers, reinforcing materials, "
        "getters, or desiccants may be incorporated. The market includes VIPs "
        "manufactured using silica, fiberglass, perlite, polymer, and other "
        "suitable core materials, covering standard flat panels, customized "
        "configurations, and VIP-integrated insulation systems for "
        "construction, refrigeration, temperature-controlled packaging, "
        "transportation, industrial equipment, and specialized applications. "
        "Geography: United States and Canada; Germany, United Kingdom, France, "
        "Spain, Italy, Russia (European territory) and the rest of Europe. "
        "By Offering: Standalone VIPs; VIP-Integrated Insulation Systems. By "
        "Configuration: Standard Flat; Custom-Designed. By Core Material: "
        "Silica, Fiberglass, Perlite, Polymer, Other. By Application: Building "
        "& Construction (Roof, Wall & Facade, Floor); Refrigeration & Freezing "
        "Equipment; Temperature-Controlled Packaging & Logistics; "
        "Transportation & Mobility; Industrial Equipment & Process; "
        "Electronics & Other Specialized. By End User: Building & "
        "Construction; Household & Commercial Appliance Manufacturing; "
        "Pharmaceutical & Healthcare; Food & Beverage; Automotive & "
        "Transportation; Aerospace & Defense; Industrial Manufacturing; Others. "
        "INCLUDED: manufactured panels with an evacuated core in a sealed "
        "gas-tight barrier; finished standalone VIPs; commercial assemblies "
        "incorporating genuine VIPs (protected VIP modules, prefabricated "
        "VIP-based systems); flat, shaped, curved, folded and application-"
        "specific panels; VIPs in building envelopes, refrigerators and "
        "freezers, insulated shipping containers and pharmaceutical / food "
        "transport packaging, vehicles, refrigerated transport, aerospace, "
        "industrial machinery, thermal storage, hot-water systems, electronics "
        "and technology installations. "
        "EXCLUDED: conventional non-evacuated insulation boards, blankets and "
        "foams; unfinished core boards and unsealed insulating sheets; "
        "insulation systems with no evacuated VIP component; conventional "
        "boards without vacuum technology; core materials sold for non-VIP "
        "use; mineral wool, EPS, XPS, PIR, PUR and fiberglass insulation "
        "without VIPs; complete refrigerators, freezers, compressors and "
        "refrigeration equipment value beyond the VIP; refrigerants, cooling "
        "units, separately sold phase-change materials, transportation and "
        "logistics services; conventional automotive insulation and complete "
        "vehicles; machinery and electronics value unrelated to the VIP; "
        "demand outside the defined countries; double-counted intermediate "
        "transactions, installation labor and aftermarket services. "
        "COMPANY TYPE IN THIS LIST - VACUUM INSULATION PANEL DISTRIBUTOR: a "
        "company headquartered in North America or Europe that buys VIPs or "
        "VIP-integrated systems from the manufacturers (e.g. va-Q-tec, "
        "Kingspan OPTIM-R, Panasonic U-Vacua, Promat, Porextherm, Kevothermal, "
        "ThermoCor) and resells them - building-insulation and construction-"
        "materials distributors, cold-chain / thermal-packaging distributors, "
        "and industrial or technical insulation distributors, stockists and "
        "importers. EXCLUDED companies: VIP manufacturers and makers of "
        "VIP-integrated systems; sellers of conventional insulation only; "
        "core-material and barrier-film suppliers; appliance, refrigeration "
        "and vehicle manufacturers; refrigerant, cooling-unit and PCM sellers; "
        "logistics and transport service providers; installers and "
        "contractors only; companies headquartered outside North America and "
        "Europe."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; keep short).
VERIFY_RULES: dict[str, str] = {
    "vip_distributor": (
        "KEEP only distributors, stockists or resellers of VACUUM INSULATION "
        "PANELS (VIPs) made by others, based in North America or Europe.\n"
        "DROP VIP makers, conventional-insulation-only sellers, core or film "
        "suppliers, appliance makers, logistics firms, installers."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for).
EVIDENCE: dict[str, str] = {
    "vip_distributor": (
        # Any "vacuum" wording (vacuum insulation / thermal / panels), the VIP
        # makers' brands, and the main European-language terms. Widened after
        # testing on the verified list: "vacuum thermal insulation" and
        # "thin vacuum panels" were missed by a vacuum-insulat-only pattern.
        r"vacuum|\bvips?\b|va-?q-?tec|optim-?r|u-?vacua|porextherm|vacupor|"
        r"kevothermal|thermocor|vakuum|sous vide|panneau.{0,15}vide|al vac[ií]o|"
        r"sottovuoto|vacuüm|próżniow|vakuov"
    ),
}


def by_slug_vip(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_VIP:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown vip market slug: {slug}")
