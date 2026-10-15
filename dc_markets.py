"""Global Data Center Market -- ONE market, worldwide (user, 2026-10-15).

Inputs: "Report_Sample_Global_Data Center Market_2026-2033 CMI.pptx" and
"Seg Definition, Exclusion and Inclusion.pptx" (Downloads). The segmentation
deck defines the market by component (hardware: servers, storage, network,
power, cooling, racks & enclosures, fire & life safety, physical security;
software: DCIM, BMS, energy, thermal, digital twin, predictive maintenance,
sustainability; services: design, installation & integration, commissioning,
maintenance, managed services, migration), by operator model, IT load and
end-use industry.

COMPANY TYPE: DISTRIBUTOR (user switched to "distributors only" at 02:00
2026-10-15; the solution-provider run is archived in
data/_archived_dc_solution_provider_*). History - first choice was DATA
CENTER SOLUTION PROVIDER / SYSTEM INTEGRATOR, chosen after asking Google AI Mode how
Dell / HPE / Cisco / Schneider / Vertiv / Eaton reach data center customers:
broadline distributors are logistics and credit only (no design or on-site
engineering), VARs sell standard IT and rarely build facility infrastructure,
pure service firms come in after installation -- solution providers / system
integrators design, source multi-vendor equipment, integrate, commission and
support the whole stack, i.e. they deliver the full scope the deck defines.
Runs under the HYBRID rule (resells / integrates OTHER vendors' hardware and
software; companies that build what they sell are manufacturers, excluded).

Scope rides in ONE Google URL with the system prompt (budget 7,550 / hard
7,800); measure with `python run_dc_market.py --slug dc_solution_provider
--region "United Arab Emirates" --show-prompt`.
"""
from __future__ import annotations

MARKETS_DC: tuple[tuple[str, str, str, str], ...] = (
    (
        "dc_distributor",
        "Global Data Center Market",
        "PRODUCT SCOPE (required): DATA CENTER DISTRIBUTORS - broadline, "
        "value-added and specialist distributors that buy from OEMs and "
        "resell data center products: servers, storage and networking; UPS, "
        "PDUs, switchgear and generators; precision and liquid cooling; "
        "racks, containment and structured cabling; fire suppression, access "
        "control and CCTV; DCIM software. OEM lines include Dell, HPE, Cisco, "
        "Schneider Electric, Vertiv, Eaton, Rittal, Legrand. EXCLUDE OEMs and "
        "manufacturers, data center operators (colocation, cloud, "
        "hyperscale), system integrators and solution providers selling "
        "projects to end users, IT resellers and MSPs, and contractors. Name "
        "the OEM brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_DC: dict[str, str] = {}

# Report regions first (North America, Europe, Asia-Pacific, MEA, Latin
# America, in the deck's country order), then the rest of the world.
_REPORT_COUNTRIES = (
    "United States", "Canada", "United Kingdom", "Germany", "France", "Italy",
    "Spain", "Russia", "China", "India", "Japan", "South Korea", "Singapore",
    "Malaysia", "Indonesia", "Thailand", "Vietnam", "Philippines", "Australia",
    "United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Oman",
    "South Africa", "Brazil", "Mexico", "Argentina",
)
_MORE = (
    "Netherlands", "Ireland", "Belgium", "Switzerland", "Sweden", "Norway",
    "Denmark", "Finland", "Poland", "Austria", "Portugal", "Czech Republic",
    "Romania", "Greece", "Turkey", "Israel", "Egypt", "Morocco", "Nigeria",
    "Kenya", "Ghana", "Tanzania", "New Zealand", "Taiwan", "Hong Kong",
    "Bangladesh", "Pakistan", "Sri Lanka", "Chile", "Colombia", "Peru",
    "Ecuador",
)
DC_REGIONS: tuple[str, ...] = _REPORT_COUNTRIES + _MORE

# Row 2 of the final Excel: the market (from the segmentation deck) plus the
# company type chosen, its definition and the reason.
DEFINITIONS: dict[str, str] = {
    "dc_distributor": (
        "Global Data Center Market: data center facilities - enterprise / "
        "institutional, colocation (retail and wholesale), hyperscale / cloud, "
        "telecom and edge - and their development, operation, capacity "
        "expansion and infrastructure. "
        "By Component - Hardware: servers (rack, blade, other), storage (DAS, "
        "NAS, SAN, other), network infrastructure (switches, routers, NICs, "
        "load balancers, security appliances, structured cabling, other), power "
        "infrastructure (UPS, PDUs, switchgear, transfer switches, "
        "transformers, backup generators, battery energy storage, other), "
        "cooling (IT-space air and liquid cooling, central cooling and heat "
        "rejection, coolant distribution, other), racks and enclosures (racks, "
        "cabinets, hot/cold-aisle containment, cable management, other), fire "
        "and life safety (detection, suppression, leak detection, alarms, "
        "EPO, other), physical security (access control, video surveillance, "
        "intrusion detection, other). Software: DCIM, building and facility "
        "management, energy and power, environmental and thermal, digital twin "
        "and capacity planning, predictive maintenance and automation, "
        "sustainability and carbon, other. Services: consulting, design and "
        "engineering; installation and systems integration; testing and "
        "commissioning; maintenance and technical support; managed "
        "infrastructure; migration and modernization; other specialized. "
        "By IT load: below 1 MW to 100 MW and above. End use: IT and telecom, "
        "BFSI, healthcare, retail and e-commerce, manufacturing, energy and "
        "utilities, government, media, education and research, transportation "
        "and logistics, others. Geography: global. "
        "EXCLUDED from the market: general commercial buildings, telecom "
        "towers, network switching facilities and other IT or communication "
        "facilities that are not dedicated data centers. "
        "COMPANY TYPE IN THIS LIST - DATA CENTER DISTRIBUTOR: a company whose "
        "primary business is buying data center products in volume from "
        "OEMs and reselling them onward - broadline IT distributors, "
        "value-added distributors (VADs) and specialist distributors of "
        "servers, storage, networking, power (UPS, PDUs, switchgear, "
        "generators), cooling, racks and cabling, fire and security systems "
        "and DCIM software - supplying resellers, system integrators and "
        "end customers, and manufacturing none of what it sells. "
        "EXCLUDED companies: OEMs and manufacturers (e.g. Dell, HPE, Cisco, "
        "Schneider Electric, Vertiv, Eaton), data center operators "
        "(colocation, cloud, hyperscale, telecom), system integrators and "
        "solution providers that sell projects to end users, IT resellers, "
        "VARs and managed-service providers, maintenance-only firms, and "
        "building or telecom-tower contractors."
    ),
}


def by_slug_dc(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_DC:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown dc market slug: {slug}")
