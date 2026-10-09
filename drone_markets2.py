"""The second batch of 15 drone markets, India-only.

Written from the user's "Exhibitor Function / Relevance" table, one
scope per row. As with the first ten, each scope names the specific
products or services that belong to it AND the neighbouring markets it
must not drift into -- without that, "Drone Mapping", "Drone LiDAR" and
"Drone Navigation" all match a generic survey-equipment query and the
same companies fill three lists.

FOUR OF THESE ARE SERVICES, not distribution:

    inspection, surveillance, testing, training

You cannot "distribute" a pilot-training course or an inspection
contract, so a distributors-only rule finds almost nobody in them --
anti_drone's 86 shows what a structural mismatch costs. Per the user's
decision these four accept SERVICE PROVIDERS; the other eleven stay
distributors-only like every market before them.

Each entry is (slug, market_name, product_scope, kind) where kind is
"distributor", "service" or "hybrid" (hardware + software, so a channel
partner or integrator is as valid as a stocking distributor).
"""
from __future__ import annotations

MARKETS2: tuple[tuple[str, str, str, str], ...] = (
    (
        "drone_lidar",
        "Global Drone LiDAR Market - Distributors Only",
        'PRODUCT SCOPE (required): LiDAR, 3D MAPPING and PRECISION '
        'SURVEYING payloads for drones - airborne LiDAR scanners, laser '
        'rangefinders, point-cloud and 3D reconstruction hardware, survey-'
        'grade RTK/PPK kits sold with LiDAR. Brands like Riegl, YellowScan, '
        'Velodyne, LiDAR USA, GeoCue. EXCLUDE plain camera/imaging sellers '
        'and GIS software firms with no LiDAR line. Name the LiDAR it '
        'distributes.',
        "distributor",
    ),
    (
        "drone_mapping",
        "Global Drone Mapping Market - Distributors Only",
        'PRODUCT SCOPE (required): GIS, AERIAL MAPPING, PHOTOGRAMMETRY and '
        'REMOTE SENSING - photogrammetry and GIS software, orthomosaic and '
        'terrain-model tools, remote-sensing data products and mapping '
        'workflow platforms. Brands like Pix4D, Agisoft Metashape, Esri, '
        'DroneDeploy. EXCLUDE LiDAR hardware specialists and pure '
        'flight-control software. Name the mapping products it handles.',
        "hybrid",
    ),
    (
        "drone_navigation",
        "Global Drone Navigation Market - Distributors Only",
        'PRODUCT SCOPE (required): GPS, NAVIGATION, GUIDANCE and '
        'POSITIONING systems for drones - GNSS receivers and antennas, '
        'RTK/PPK base stations, INS/IMU units, autopilots, guidance '
        'computers and anti-jam positioning. Brands like Trimble, Septentrio, '
        'u-blox, Hemisphere, VectorNav. EXCLUDE mapping-software firms and '
        'complete-aircraft sellers. Name the navigation hardware it '
        'distributes.',
        "distributor",
    ),
    (
        "drone_communications",
        "Global Drone Communications Market - Distributors Only",
        'PRODUCT SCOPE (required): WIRELESS CONNECTIVITY, COMMUNICATIONS '
        'and DATA LINKS for drones - radio and telemetry modules, C2 and '
        'video data links, BVLOS and mesh radios, SATCOM and 4G/5G drone '
        'connectivity, antennas and ground radios. Brands like Doodle Labs, '
        'Microhard, Silvus, uAvionix. EXCLUDE navigation/GNSS specialists '
        'and cloud-platform firms. Name the comms hardware it distributes.',
        "distributor",
    ),
    (
        "drone_batteries",
        "Global Drone Batteries Market - Distributors Only",
        'PRODUCT SCOPE (required): UAV BATTERIES, POWERPACKS and ENERGY '
        'STORAGE - lithium-polymer and Li-ion drone battery packs, '
        'smart batteries, hydrogen fuel cells and tethered power for UAVs, '
        'plus cells and BMS sold for drone use. EXCLUDE chargers and '
        'charging infrastructure (a separate market), and complete-aircraft '
        'sellers. Name the battery brands it distributes.',
        "distributor",
    ),
    (
        "drone_charging",
        "Global Drone Charging Market - Distributors Only",
        'PRODUCT SCOPE (required): CHARGING INFRASTRUCTURE, POWER SUPPLY '
        'and BATTERY MANAGEMENT - drone charging pads and docking stations, '
        'fast and multi-bay chargers, battery-management and swap systems, '
        'field power supplies and generators for UAV fleets. EXCLUDE '
        'battery-cell sellers with no charging line, and drone-in-a-box '
        'robotics firms. Name the charging products it distributes.',
        "distributor",
    ),
    (
        "drone_propulsion",
        "Global Drone Propulsion Market - Distributors Only",
        'PRODUCT SCOPE (required): MOTORS, PROPULSION SYSTEMS, PROPELLERS '
        'and DRIVE TECHNOLOGIES - brushless motors, ESCs, propellers and '
        'rotors, gearboxes, UAV engines and hybrid powertrains, thrust '
        'stands. Brands like T-Motor, Hobbywing, KDE Direct, Mejzlik. '
        'EXCLUDE battery and charging specialists, and complete-aircraft '
        'sellers. Name the propulsion products it distributes.',
        "distributor",
    ),
    (
        "drone_components",
        "Global Drone Components Market - Distributors Only",
        'PRODUCT SCOPE (required): ELECTRONICS, WIRES, CONNECTORS, '
        'SEMICONDUCTORS and COMPONENTS for drones - PCBs and electronic '
        'assemblies, connectors and cable harnesses, semiconductors, '
        'passives, fasteners and mechanical hardware for UAV builds. '
        'EXCLUDE sellers of complete aircraft, and specialists in motors, '
        'batteries or sensors alone. Name the component lines it '
        'distributes.',
        "distributor",
    ),
    (
        "drone_composites",
        "Global Drone Composites Market - Distributors Only",
        'PRODUCT SCOPE (required): CARBON FIBRE, PLASTICS, MOULDING and '
        'LIGHTWEIGHT STRUCTURES - carbon-fibre tubes, sheets and prepreg, '
        'composite airframe parts, engineering plastics and 3D-printing '
        'materials for UAVs, moulding and tooling services for airframes. '
        'EXCLUDE electronics distributors and complete-aircraft sellers. '
        'Name the composite materials it distributes.',
        "distributor",
    ),
    (
        "drone_cybersecurity",
        "Global Drone Cybersecurity Market - Distributors Only",
        'PRODUCT SCOPE (required): DATA SECURITY, NETWORK SECURITY and UAV '
        'CYBERSECURITY - secure C2 links and encryption modules, drone '
        'network and endpoint security, anti-spoofing and secure firmware, '
        'compliance and audit tooling for UAV fleets. EXCLUDE counter-drone '
        'jamming/detection hardware (a separate market) and generic IT '
        'security firms with no UAV line. Name the security products it '
        'handles.',
        "hybrid",
    ),
    (
        "drone_cloud",
        "Global Drone Cloud Platforms Market - Distributors Only",
        'PRODUCT SCOPE (required): CLOUD, SERVERS, DATABASES and DRONE DATA '
        'INFRASTRUCTURE - cloud platforms for drone data, fleet and mission '
        'data storage, imagery pipelines and APIs, edge/server hardware and '
        'hosting sold for UAV data. EXCLUDE analytics-only firms and '
        'flight-planning software specialists. Name the cloud platforms it '
        'resells or deploys.',
        "hybrid",
    ),
    (
        "drone_inspection",
        "Global Drone Inspection Market - India",
        'PRODUCT SCOPE (required): INFRASTRUCTURE, ENERGY, INDUSTRIAL and '
        'FACILITY INSPECTION using drones - power line and substation '
        'inspection, wind turbine and solar farm inspection, oil and gas '
        'and refinery inspection, bridges, towers, buildings and industrial '
        'plant surveys. EXCLUDE pure equipment sellers with no inspection '
        'service, and mapping/survey-only firms. Name the inspection '
        'services it provides.',
        "service",
    ),
    (
        "drone_surveillance",
        "Global Drone Surveillance Market - India",
        'PRODUCT SCOPE (required): MONITORING, SURVEILLANCE, SECURITY and '
        'INTELLIGENCE using drones - perimeter and site security patrols, '
        'crowd and event monitoring, border and police aerial surveillance, '
        'ISR and reconnaissance services, traffic and disaster monitoring. '
        'EXCLUDE counter-drone specialists and camera-hardware sellers with '
        'no surveillance operation. Name the surveillance services it '
        'provides.',
        "service",
    ),
    (
        "drone_testing",
        "Global Drone Testing Market - India",
        'PRODUCT SCOPE (required): TESTING, VALIDATION, CERTIFICATION and '
        'COMPLIANCE for drones - DGCA type certification support, airworthiness '
        'and flight testing, EMI/EMC and environmental testing, test ranges '
        'and labs, regulatory and compliance consultancy for UAV operators. '
        'EXCLUDE manufacturers testing only their own products, and pure '
        'equipment sellers. Name the testing or certification services it '
        'provides.',
        "service",
    ),
    (
        "drone_training",
        "Global Drone Training Market - India",
        'PRODUCT SCOPE (required): PILOT TRAINING, STEM EDUCATION and UAV '
        'SKILL DEVELOPMENT - DGCA-approved remote pilot training '
        'organisations (RPTO), drone pilot certification courses, drone '
        'STEM and school programmes, industrial UAV skilling and '
        'train-the-trainer. EXCLUDE manufacturers offering only product '
        'familiarisation, and general aviation schools with no drone '
        'course. Name the training programmes it runs.',
        "service",
    ),
)

# Sibling markets, to stop drift between neighbours that share vocabulary.
SIBLINGS2 = {
    "drone_lidar": "drone mapping, drone sensors",
    "drone_mapping": "drone LiDAR, drone analytics",
    "drone_navigation": "drone communications, drone components",
    "drone_communications": "drone navigation, drone cloud platforms",
    "drone_batteries": "drone charging, drone propulsion",
    "drone_charging": "drone batteries, drone robotics",
    "drone_propulsion": "drone components, drone batteries",
    "drone_components": "drone propulsion, drone composites",
    "drone_composites": "drone components, UAV airframes",
    "drone_cybersecurity": "anti-drone systems, drone cloud platforms",
    "drone_cloud": "drone analytics, drone software",
    "drone_inspection": "drone surveillance, drone mapping",
    "drone_surveillance": "drone inspection, anti-drone systems",
    "drone_testing": "drone training, drone certification consultancy",
    "drone_training": "drone testing, drone inspection services",
}


def by_slug2(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS2:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown market slug: {slug}. "
                   f"Valid: {', '.join(m[0] for m in MARKETS2)}")
