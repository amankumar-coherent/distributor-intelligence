"""Brazil Industrial Electrification Market -- ONE market, Brazil-based
SOLUTION PROVIDERS (user, 2026-11-18: CMI sample report + taxonomy + the
user's own definition with inclusion / exclusion per segment). Cap 400,
10 browsers, through the Proton VPN guard.

The market covers hardware (electric heating / boilers / heat pumps /
furnaces, electrified machinery and forklifts, motors and drives, switchgear,
transformers, PDUs, cables, actuators), software (industrial energy
management, electrical monitoring and control) and services (electrification
engineering, installation / integration / retrofit, O&M, equipment leasing,
electrification-as-a-service). So the companies are the SELLERS of those:
distributors and authorized dealers, system integrators, engineering / EPC
contractors, O&M and rental firms and the Brazilian makers / units of the
equipment -- the "solution" kind (discovery.DEF_SOLUTION).

Exclusions applied to COMPANIES (out when its work is only in an excluded
area): fossil-fuel boilers / furnaces, diesel machinery, mechanical /
hydraulic / pneumatic drives only, residential and commercial electrical
(building wiring, home electricians, retail electrical stores), general ERP /
business software, general industrial consulting, plain electricity retail /
standalone power generation / utilities without dedicated industrial
electrification, consumer EV charging.

Discovery uses SIMPLE QUERY mode (as sorb / eth); the verifier enforces the
rules (VERIFY_RULES). Brazil-based only (set_country_restriction).
"""
from __future__ import annotations

from water_markets import BRAZIL_REGIONS

SLUG = "bie_industrial_electrification"
MARKET = "Brazil Industrial Electrification Market"

MARKETS_BIE: tuple[tuple[str, str, str, str], ...] = (
    (
        SLUG,
        MARKET,
        "PRODUCT/SOLUTION SCOPE (required): INDUSTRIAL ELECTRIFICATION in Brazil - "
        "electric boilers, heat pumps and furnaces, electric forklifts and mining "
        "machinery, industrial motors and drives (VFD), switchgear, transformers, "
        "panels, cables, actuators, industrial energy management and electrical "
        "control software, and electrification engineering, installation, retrofit, "
        "O&M, leasing and electrification-as-a-service. The company distributes, "
        "integrates, installs, makes or services these for industry, for example "
        "WEG, Comerc, Ecogen, distributors of ABB, Siemens, Schneider, Danfoss. "
        "EXCLUDE fossil or diesel equipment, hydraulic or pneumatic only, "
        "residential or building electrical, retail stores, ERP software and plain "
        "power generation or electricity retail.",
        "solution",
    ),
)

SIBLINGS_BIE: dict[str, str] = {}

SIMPLE_REQUEST: dict[str, str] = {
    SLUG: ("distributors, dealers, system integrators, engineering/installation and "
           "service companies or makers of industrial electrification equipment "
           "(industrial electric motors and drives, switchgear, transformers, "
           "electric boilers and heat pumps, electric forklifts, industrial energy "
           "management) for factories and mining"),
}

# 27 states / metro groups, then industrial / port cities (from Brazil RO),
# plus extra industrial hubs for electrical equipment.
BIE_REGIONS: tuple[str, ...] = tuple(BRAZIL_REGIONS) + (
    "Jundiai", "Piracicaba", "Limeira", "Americana", "Indaiatuba", "Taubate",
    "Jaragua do Sul", "Criciuma", "Caxias do Sul", "Novo Hamburgo", "Maringa",
    "Londrina", "Cascavel", "Uberlandia", "Sete Lagoas", "Montes Claros",
    "Campos dos Goytacazes", "Macae", "Feira de Santana", "Sao Luis", "Manaus industrial pole",
    "Barcarena", "Maraba", "Anapolis", "Rio Verde", "Itabira", "Ouro Preto", "Ipojuca",
)

DEFINITIONS: dict[str, str] = {
    SLUG: (
        "The Brazil Industrial Electrification Market encompasses hardware, software, "
        "and services used to electrify industrial processes, machinery, material "
        "handling equipment, and electrical power distribution infrastructure across "
        "Brazil. The market includes electric heating and steam generation systems, "
        "industrial motors and drives, electrified industrial machinery, electrical "
        "distribution and control equipment, energy management software, and "
        "associated engineering, installation, retrofit, and maintenance services. It "
        "covers both greenfield and brownfield industrial electrification projects "
        "across manufacturing, mining, oil and gas, chemicals, and other industrial "
        "sectors. "
        "BY COMPONENT: Hardware (Industrial Electric Heating Equipment - electric "
        "boilers and steam generation systems, industrial heat pumps, electric "
        "furnaces and process heating systems; Electrified Industrial Machinery and "
        "Equipment - electric excavators, electric wheel loaders, electric forklifts, "
        "electric mining trucks and material handling equipment; Industrial Electric "
        "Motors and Drive Systems; Industrial Electrical Devices and Power "
        "Distribution Equipment - electric actuators, switchgear, power distribution "
        "units, transformers and substations, electrical cables and connection "
        "equipment); Software (Industrial Energy Management and Optimization; "
        "Industrial Electrical Monitoring and Control; Other); Services (Consulting "
        "and Engineering; Installation, Integration and Retrofit; Operation and "
        "Maintenance; Equipment Leasing and Rental; Electrification-as-a-Service and "
        "dedicated renewable electricity supply for industrial electrification). "
        "BY APPLICATION: Process Heating and Steam Generation; Mechanical Drives and "
        "Motion Control; Material Handling and Transportation; Electrical Power "
        "Distribution and Management; Other (industrial HVAC). BY DEPLOYMENT TYPE: "
        "Greenfield; Brownfield and Retrofit. BY END-USE INDUSTRY: Chemicals and "
        "Petrochemicals; Mining and Mineral Processing; Food and Beverages; Pulp and "
        "Paper; Oil and Gas; Automotive and Transportation Equipment Manufacturing; "
        "Construction; Logistics and Warehousing; Textiles; Other (Pharmaceuticals, "
        "Electronics, Water and Wastewater). BY SALES CHANNEL: Online; Offline "
        "(direct manufacturer sales, authorized dealers and distributors, system "
        "integrators, engineering and EPC contractors). "
        "EXCLUSIONS: conventional fossil-fuel boilers, furnaces and steam systems; "
        "diesel excavators, loaders, forklifts and combustion machinery; mechanical, "
        "pneumatic or hydraulic drives without electric motors; residential and "
        "commercial electrical distribution, building HVAC and smart-home products; "
        "general ERP, accounting and business software; general industrial consulting "
        "without electrification; general electricity retail, standalone renewable "
        "power generation and utility services without dedicated industrial "
        "electrification; consumer and passenger vehicle charging. "
        "COMPANY TYPE IN THIS LIST - INDUSTRIAL ELECTRIFICATION SOLUTION PROVIDER: a "
        "company based in Brazil that distributes, resells, integrates, installs, "
        "manufactures, leases or maintains industrial electrification hardware or "
        "software, or provides industrial electrification engineering, retrofit, O&M "
        "or electrification-as-a-service - authorized distributors and dealers, "
        "system integrators, engineering / EPC contractors, O&M and rental companies "
        "and equipment makers serving industrial facilities. EXCLUDED companies: "
        "those working only in residential / commercial building electrical, retail "
        "electrical stores, fossil-fuel or diesel equipment, general software, or "
        "plain power generation and electricity retail; media, associations and "
        "directories."
    ),
}

# Verifier KEEP/DROP (criteria block must stay under ~1,000 chars).
VERIFY_RULES: dict[str, str] = {
    SLUG: (
        "KEEP Brazil-based firms selling INDUSTRIAL electrification: distributors, "
        "dealers, integrators, EPC/installation, retrofit, O&M, rental or makers of "
        "industrial motors/drives, switchgear, transformers, panels, cables, electric "
        "boilers/heat pumps/furnaces, electric forklifts or mining machinery, energy "
        "management/electrical control software, electrification-as-a-service.\n"
        "DROP firms only in residential/commercial building electrical, retail "
        "electrical stores, fossil/diesel equipment, hydraulic/pneumatic only, ERP "
        "software, plain power generation or electricity retail; HQ outside Brazil."
    ),
}

# Product-evidence gate for the final build.
EVIDENCE: dict[str, str] = {
    SLUG: (
        r"electr|eletr|motor|drive|inversor|inverter|\bvfd\b|servo|switchgear|"
        r"transform|substat|subesta|panel|painel|quadro|cable|cabo|actuator|atuador|"
        r"boiler|caldeira|heat pump|bomba de calor|furnace|forno|forklift|empilhadeira|"
        r"energy|energia|automa|power|pot[eê]ncia|scada|\bplc\b|clp|retrofit|"
        r"integrat|integra[cç]|automation|instrument"
    ),
}

TARGETS: dict[str, int] = {SLUG: 400}


def by_slug_bie(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_BIE:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown bie market slug: {slug}")
