"""9 global drone / geospatial / digital-twin markets (user CSVs, 2026-10-14).

Source: "Report-Producttechnologysegmentation-Servicebusine.csv" and its "(1)"
copy in Downloads. Each row gives product/technology, service/business-model,
application, end-user and geographic segmentation; every row names India as a
focus country, so discovery searches India first.

These are mostly SERVICE and SOFTWARE markets, so one "stocking distributor"
rule would find almost nobody. Same 4-way rule system as fire_markets.py:

  HYBRID   - hardware and/or platforms sold through dealers, resellers and
             system integrators (survey drones + LiDAR, GeoAI, ag drones,
             digital twin platforms)
  SOFTWARE - resellers / distributors / implementation partners of another
             vendor's platform (GIS)
  SERVICE  - companies that PERFORM the work for paying clients (commercial
             drone services, infrastructure inspection, telecom tower and
             power-line / utility inspection)

Scopes are kept around 500 characters: discovery sends system prompt + scope
in ONE Google URL (budget 7,550, hard limit 7,800). See wound_markets.py.
"""
from __future__ import annotations

MARKETS_GEO: tuple[tuple[str, str, str, str], ...] = (
    (
        "geo_drone_survey_lidar",
        "Global Drone Surveying & LiDAR Mapping Market",
        "PRODUCT/SOLUTION SCOPE (required): DRONE SURVEYING AND LIDAR MAPPING - "
        "rotary-wing, fixed-wing and VTOL survey drones, drone LiDAR, "
        "photogrammetry, thermal and multispectral sensors, GNSS/RTK/PPK and "
        "IMU positioning, and mission-planning and point-cloud software, for "
        "topographic, corridor, construction, mining and forestry survey. "
        "Brands include DJI Enterprise, YellowScan, RIEGL, Quantum Systems. "
        "EXCLUDE pure survey-flying operators and consumer drone shops. Name "
        "the survey drone, LiDAR or software brands it sells or integrates.",
        "hybrid",
    ),
    (
        "geo_commercial_drone_services",
        "Global Commercial Drone Services Market",
        "SERVICE SCOPE (required): COMMERCIAL DRONE SERVICES - companies that "
        "PERFORM drone work for paying clients: aerial imaging, mapping and "
        "surveying, inspection, agricultural spraying, data analytics, "
        "delivery, surveillance, BVLOS operations, MRO, training, rental and "
        "fleet management, for agriculture, construction, energy, mining, "
        "telecom, logistics and government. EXCLUDE drone makers and "
        "resellers with no service operations, and hobby photographers. Name "
        "the drone services it performs.",
        "service",
    ),
    (
        "geo_gis_analytics",
        "Global Geospatial Analytics / GIS Market",
        "SOFTWARE SCOPE (required): GEOSPATIAL ANALYTICS AND GIS - desktop, "
        "web, mobile and enterprise GIS software and geospatial data "
        "(satellite, drone, LiDAR, cadastral), resold, licensed or "
        "implemented for government, utilities, telecom, transport, "
        "agriculture and smart cities. Platforms include Esri ArcGIS, "
        "Hexagon, Bentley, Trimble, Maxar, Planet. EXCLUDE the GIS vendors "
        "themselves, field survey crews and IT resellers with no GIS line. "
        "Name the GIS platforms and data it resells or implements.",
        "software",
    ),
    (
        "geo_ai_geospatial",
        "Global AI in Geospatial Imaging & Remote Sensing Market",
        "PRODUCT/SOLUTION SCOPE (required): AI IN GEOSPATIAL IMAGING AND "
        "REMOTE SENSING - GeoAI and computer-vision platforms for drone, "
        "satellite, LiDAR, thermal and hyperspectral imagery, plus edge AI "
        "hardware, for infrastructure inspection, crop health, land use, "
        "utilities, disaster response and defense. Platforms include Esri "
        "GeoAI, Picterra, EOS Data Analytics, Planet, Maxar. EXCLUDE the "
        "platform vendors themselves and AI consultancies with no geospatial "
        "line. Name the GeoAI platforms or imagery it resells or implements.",
        "hybrid",
    ),
    (
        "geo_drone_infra_inspection",
        "Global Drone-Based Infrastructure Inspection Market",
        "SERVICE SCOPE (required): DRONE-BASED INFRASTRUCTURE INSPECTION - "
        "companies that PERFORM drone inspections of buildings, bridges, "
        "roads, railways, tunnels, dams, industrial plants, chimneys, solar "
        "farms and wind turbines, with RGB, zoom, thermal, LiDAR or "
        "ultrasonic payloads, as one-time, recurring or drone-in-a-box "
        "programs with defect reports. EXCLUDE drone resellers with no "
        "inspection service and telecom-tower-only or power-line-only "
        "specialists. Name the inspection services it performs.",
        "service",
    ),
    (
        "geo_telecom_tower_inspection",
        "Global Telecom Tower Inspection & Maintenance Market",
        "SERVICE SCOPE (required): TELECOM TOWER INSPECTION AND MAINTENANCE - "
        "companies that PERFORM tower inspection and maintenance for "
        "operators and tower companies: drone structural, antenna and feeder "
        "inspection, tilt and azimuth audits, corrosion and crack detection, "
        "RF obstruction analysis, 3D tower models and maintenance "
        "verification. EXCLUDE tower owners, telecom equipment vendors and "
        "drone resellers with no tower service. Name the tower inspection or "
        "maintenance services it performs.",
        "service",
    ),
    (
        "geo_powerline_utility_drone",
        "Global Power Line, Utility Asset & Vegetation Management Drone Market",
        "SERVICE SCOPE (required): POWER LINE, UTILITY ASSET AND VEGETATION "
        "MANAGEMENT - companies that PERFORM drone, helicopter or satellite "
        "LiDAR, thermal and corona inspection of transmission and "
        "distribution lines and substations, vegetation encroachment "
        "analysis, wildfire mitigation and storm-damage assessment for "
        "utilities. EXCLUDE the utilities themselves, drone resellers with "
        "no utility service, and tree crews with no aerial inspection. Name "
        "the utility inspection services it performs.",
        "service",
    ),
    (
        "geo_ag_drone",
        "Global Precision Agriculture & Agricultural Drone Market",
        "PRODUCT/SOLUTION SCOPE (required): PRECISION AGRICULTURE AND "
        "AGRICULTURAL DRONES - spraying, mapping, scouting and seeding "
        "drones, spray systems, multispectral sensors and NDVI or "
        "prescription-map software for crop monitoring, precision spraying, "
        "field mapping and yield estimation. Brands include DJI Agras, XAG, "
        "Hylio, Garuda Aerospace. EXCLUDE drone makers' own sales arms, farm "
        "input dealers with no drone line and spraying-only operators. Name "
        "the agricultural drone brands it sells or supports.",
        "hybrid",
    ),
    (
        "geo_digital_twin",
        "Global Digital Twin Market - Infrastructure, Utilities & Smart Cities",
        "PRODUCT/SOLUTION SCOPE (required): DIGITAL TWIN FOR INFRASTRUCTURE, "
        "UTILITIES AND SMART CITIES - asset, network and city twin platforms "
        "combining IoT, AI, BIM, GIS and drone or LiDAR reality capture with "
        "2D/3D dashboards, for construction, operations and maintenance and "
        "urban planning. Platforms include Bentley iTwin, Siemens, Dassault "
        "3DEXPERIENCE, Autodesk Tandem, NVIDIA Omniverse, Esri. EXCLUDE the "
        "platform vendors themselves and factory or product twins. Name the "
        "digital twin platforms it resells or implements.",
        "hybrid",
    ),
)

SIBLINGS_GEO: dict[str, str] = {
    # Overlaps are handled inside each scope's EXCLUDE clause (e.g. the
    # infrastructure-inspection scope excludes tower-only and power-line-only
    # specialists); a separate sibling sentence would cost URL budget.
}

# Every CSV row names India as a focus country; the others named as priority
# markets (U.S., Brazil, Australia, China, Southeast Asia, key EU, Africa)
# come next, then the rest of the world.
GLOBAL_REGIONS: tuple[str, ...] = (
    "India", "United States", "China", "United Kingdom", "Germany",
    "Australia", "Brazil", "France", "Canada", "Japan", "United Arab Emirates",
    "Saudi Arabia", "Indonesia", "Singapore", "Malaysia", "Thailand",
    "Vietnam", "Philippines", "South Korea", "Italy", "Spain", "Netherlands",
    "Belgium", "Switzerland", "Sweden", "Norway", "Denmark", "Finland",
    "Poland", "Turkey", "Israel", "South Africa", "Nigeria", "Kenya", "Egypt",
    "Morocco", "Mexico", "Chile", "Colombia", "Peru", "Argentina",
    "New Zealand", "Ireland", "Austria", "Portugal", "Czech Republic",
    "Romania", "Greece", "Qatar", "Oman", "Kuwait", "Bangladesh", "Pakistan",
    "Sri Lanka", "Taiwan", "Hong Kong", "Ghana", "Tanzania", "Ecuador",
)


# SEARCH PLAN per market (user, 2026-10-14: "in place of country-wise, search
# global, and focus only on the countries each market mentions"). Each market
# is searched WORLDWIDE (three passes - a pass ends after ~60 new companies,
# so one pass alone stops short), then on the focus areas named in ITS OWN
# Geographic column of the CSV, then once per report region for global
# spread. No country-by-country walk.
FOCUS_BY_SLUG: dict[str, tuple[str, ...]] = {
    "geo_drone_survey_lidar": ("India",),
    "geo_commercial_drone_services": ("India", "United States", "European Union"),
    "geo_gis_analytics": ("India",),
    "geo_ai_geospatial": ("India", "Southeast Asia"),
    "geo_drone_infra_inspection": ("India",),
    "geo_telecom_tower_inspection": ("India", "Southeast Asia", "Africa"),
    "geo_powerline_utility_drone": ("India", "Australia", "Brazil", "United States"),
    "geo_ag_drone": ("India", "Brazil", "United States", "China", "Southeast Asia"),
    "geo_digital_twin": ("India",),
}
_REPORT_REGIONS = ("North America", "Europe", "Asia-Pacific", "Latin America",
                   "Middle East & Africa")

# "Worldwide..." regions carry no FOCUS sentence (see run_geo_market.py), so
# the query is genuinely global. Names stay distinct ("Worldwide 2") because
# the launcher records finished regions by name.
REGIONS_BY_SLUG: dict[str, tuple[str, ...]] = {
    slug: ("Worldwide",) + focus + ("Worldwide 2",) + _REPORT_REGIONS + ("Worldwide 3",)
    for slug, focus in FOCUS_BY_SLUG.items()
}


def _definitions() -> dict[str, str]:
    """Client-facing market definition per slug, written from the user's CSV
    columns (queries/geo_markets_csv.json) so each file states ITS market:
    product/technology, business model, applications, end users, geography."""
    import json
    import re
    from pathlib import Path

    rows = json.loads((Path(__file__).resolve().parent / "queries" / "geo_markets_csv.json")
                      .read_text(encoding="utf-8"))

    def clean(s: str) -> str:
        s = re.sub(r"\s*\b[a-z]+\+\d+\s*$", "", (s or "").strip())  # stray "site+1" citations
        return s.rstrip(" .;")

    out = {}
    for (slug, _market, _scope, _kind), r in zip(MARKETS_GEO, rows):
        title = r["Report"].replace("�", "-").strip()
        out[slug] = (
            f"{title}. Product / technology: {clean(r['Product / technology segmentation'])}. "
            f"Business models: {clean(r['Service / business-model segmentation'])}. "
            f"Applications: {clean(r['Application segmentation'])}. "
            f"End users: {clean(r['End-user segmentation'])}. "
            f"Geography: {clean(r['Geographic segmentation'])}."
        )
    return out


DEFINITIONS: dict[str, str] = _definitions()


def by_slug_geo(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_GEO:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown geo market slug: {slug}")
