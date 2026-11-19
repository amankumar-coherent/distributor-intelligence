"""India Industrial Water & Wastewater Market -- ONE market, India-HQ only,
SOLUTION PROVIDERS (user, 2026-11-02; cap 300, 10 browsers).

Inputs: the user's market definition, key inclusions, key exclusions and the
12 segmentation parameters (the report is for SUEZ). The market is REVENUE from
designing, supplying, installing, operating and maintaining industrial water
and wastewater treatment systems in India, so the companies are the SELLERS of
those solutions: EPC / EPCM / DBO contractors, treatment technology and
equipment suppliers (own-make is fine, inclusion 3) and O&M / lifecycle
service companies (inclusions 9-11). That fits none of the older kinds
(distributor drops makers, service drops equipment suppliers) so it uses the
"solution" kind (discovery.DEF_SOLUTION, verdict solution_provider).

The user's 10 EXCLUSIONS are applied to COMPANIES (a company is out when its
work is only in an excluded area): municipal / urban water only, residential
and domestic treatment (household purifiers), agricultural irrigation and
rural water, commercial building / office / hotel treatment, bottled water,
bulk water trading and tanker supply, general plant construction, standalone
chemical / consumables sellers. (Regulatory fines and duplicate revenue
counting are revenue notes, not company types.)

Searched state by state, then industrial cities. The scope rides in ONE Google
URL (budget 7,550 / hard 7,800): no semicolons (stage-4 shrink cuts them);
measure with scripts/preflight_prompt.py --batch iww.
"""
from __future__ import annotations

from india_cities import INDIA_CITIES
from india_states import INDIA_STATES

MARKETS_IWW: tuple[tuple[str, str, str, str], ...] = (
    (
        "iww_industrial",
        "India Industrial Water & Wastewater Market",
        "PRODUCT/SOLUTION SCOPE (required): INDUSTRIAL WATER AND WASTEWATER "
        "TREATMENT for factories and industrial plants in India - process, "
        "boiler feed and high-purity water, cooling water, effluent "
        "treatment (ETP), water reuse and recycling, desalination, zero "
        "liquid discharge (ZLD), brine and sludge management, common "
        "effluent treatment plants (CETP), plus O&M, spares and lifecycle "
        "services. The company designs, supplies, installs, operates or "
        "maintains these systems as an EPC or EPCM contractor, treatment "
        "technology or equipment supplier or O&M provider, for example "
        "Thermax, VA Tech Wabag, Ion Exchange, Veolia, Driplex, Doshion. "
        "EXCLUDE municipal or urban water only, household water purifiers, "
        "irrigation and rural water, building or hotel treatment, bottled "
        "water, tanker or bulk water supply, general construction and "
        "chemical-only sellers.",
        "solution",
    ),
)

SIBLINGS_IWW: dict[str, str] = {}

# 26 states / UT groups (industry-first order) then the 43 industrial cities.
IWW_REGIONS: tuple[str, ...] = tuple(INDIA_STATES) + tuple(INDIA_CITIES)

DEFINITIONS: dict[str, str] = {
    "iww_industrial": (
        "The India Industrial Water & Wastewater Market comprises revenue "
        "generated through the design, engineering, supply, installation, "
        "modernization, operation, maintenance and lifecycle management of "
        "water and wastewater treatment systems serving industrial "
        "facilities in India. It covers industrial process water, "
        "high-purity water, boiler feedwater, cooling water, effluent "
        "treatment, water reuse and recycling, desalination, zero liquid "
        "discharge (ZLD), brine management, sludge treatment and other "
        "specialized industrial water treatment applications. "
        "KEY INCLUSIONS: 1. Greenfield industrial water and wastewater "
        "treatment projects. 2. Brownfield expansion, modernization and "
        "retrofit projects. 3. Industrial water and wastewater treatment "
        "equipment and technology supply. 4. Raw water, process water and "
        "high-purity water treatment. 5. Industrial effluent and wastewater "
        "treatment. 6. Water reuse, recycling and desalination. 7. Zero "
        "liquid discharge (ZLD), brine and sludge management. 8. Cooling "
        "water and boiler water treatment. 9. Operations and maintenance "
        "(O&M) services. 10. Aftermarket spare parts, consumables, repair "
        "and refurbishment. 11. Process optimization, digital monitoring "
        "and lifecycle services. 12. Common Effluent Treatment Plants "
        "(CETPs) and shared industrial water infrastructure. 13. EPC, EPCM, "
        "DBO, BOO and other relevant project and service contracts. "
        "KEY EXCLUSIONS: 1. Municipal and urban water or wastewater "
        "treatment unrelated to industrial facilities. 2. Residential and "
        "domestic water treatment. 3. Agricultural irrigation and rural "
        "water infrastructure. 4. Water treatment for commercial buildings, "
        "offices and hospitality facilities. 5. Bottled-water sales and "
        "packaged drinking water distribution unrelated to industrial "
        "treatment. 6. Bulk water trading, tanker supply and raw water "
        "distribution without treatment. 7. General industrial plant "
        "construction and infrastructure not directly attributable to water "
        "or wastewater treatment. 8. Standalone chemical sales and "
        "consumables unrelated to treatment operations or services. 9. "
        "Regulatory penalties and environmental fines. 10. Duplicate "
        "counting of integrated project, equipment and recurring O&M "
        "revenues. "
        "BY INDUSTRIAL SECTOR: Chemicals and Specialty Chemicals; "
        "Pharmaceuticals and Biotechnology; Oil Refining and "
        "Petrochemicals; Power Generation; Iron, Steel and Other Metals; "
        "Textiles and Apparel Processing; Food and Beverage Processing; "
        "Pulp and Paper; Automotive and Auto Components; Electronics and "
        "Semiconductor Manufacturing; Cement and Building Materials; Mining "
        "and Mineral Processing; Other Industrial Sectors (Fertilizers, "
        "Glass, Rubber and Plastics, Industrial Machinery and Other "
        "Manufacturing). BY TREATMENT SOLUTION: Raw Water and Process Water "
        "Treatment; Demineralized and High-Purity Water Treatment; "
        "Industrial Wastewater and Effluent Treatment; Industrial Water "
        "Reuse and Recycling; Industrial Desalination; Zero Liquid "
        "Discharge and Brine Management; Sludge Treatment and Residuals "
        "Management; Other Industrial Water Treatment Solutions (Cooling "
        "Water Treatment, Boiler Water Conditioning, Specialized "
        "Treatment). BY REVENUE STREAM: New Installation and Greenfield "
        "Project; Brownfield Expansion, Modernization and Retrofit; "
        "Standalone Equipment and Technology Sales; Operations and "
        "Maintenance Services; Aftermarket Spare Parts and Consumables; "
        "Repair, Refurbishment and Lifecycle Technical Services; Digital "
        "Monitoring and Optimization Services. BY TREATMENT PLANT CAPACITY: "
        "Small (below 1 MLD); Medium (1 to below 10 MLD); Large (10 to "
        "below 50 MLD); Very Large (50 MLD and above). BY PRIMARY "
        "APPLICATION: Industrial Process Water Supply; Boiler Feedwater and "
        "High-Purity Water Supply; Cooling Water Supply and Treatment; "
        "Industrial Effluent Treatment and Compliance; Industrial Water "
        "Reuse and Recycling; Industrial Desalination and Freshwater "
        "Substitution; ZLD and Brine Management; Shared Industrial Water "
        "and Wastewater Infrastructure. BY GEOGRAPHY: North, South, East, "
        "West, Central and Northeast India. "
        "COMPANY TYPE IN THIS LIST - INDUSTRIAL WATER & WASTEWATER SOLUTION "
        "PROVIDER: a company headquartered in India that designs, "
        "engineers, supplies, installs, modernizes, operates or maintains "
        "industrial water and wastewater treatment systems, or makes and "
        "supplies the core treatment equipment and technology, or provides "
        "their O&M, spares and lifecycle services - EPC / EPCM / DBO / BOO "
        "contractors, treatment technology and equipment suppliers and O&M "
        "service companies serving industrial facilities. EXCLUDED "
        "companies: those working only in municipal or urban water, "
        "household or domestic purifiers, agricultural irrigation or rural "
        "water, commercial building / office / hospitality treatment, "
        "bottled or packaged drinking water, bulk water trading and tanker "
        "supply, general plant construction, or standalone chemical and "
        "consumables sales; media, associations and directories."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; the whole
# criteria block must stay under ~1,000 chars, so this is tight).
VERIFY_RULES: dict[str, str] = {
    "iww_industrial": (
        "KEEP India firms that design, supply, install or run INDUSTRIAL "
        "water/wastewater treatment (ETP, DM/RO, ZLD, reuse, desalination, "
        "CETP): EPC, technology/equipment makers, O&M.\n"
        "DROP municipal-only, household purifiers, irrigation, "
        "building/hotel plants, bottled/tanker water, general construction, "
        "chemical-only sellers."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for).
EVIDENCE: dict[str, str] = {
    "iww_industrial": (
        r"water|wastewater|effluent|\betp\b|\bstp\b|\bzld\b|zero liquid|"
        r"desalinat|reverse osmosis|\bro\b|demineral|\bdm plant|membrane|"
        r"filtrat|sludge|brine|boiler|cooling|ion exchange|recycl|reuse|"
        r"clarifier|\bcetp\b|sewage|osmosis|aeration|\bmbr\b|\bmbbr\b"
    ),
}


def by_slug_iww(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_IWW:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown iww market slug: {slug}")
