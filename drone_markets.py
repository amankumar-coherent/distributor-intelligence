"""The ten drone markets, each with a scope that distinguishes it.

These markets overlap by name -- "UAV Systems", "Commercial Drones" and
"Defense Drones" would all match a generic "drone distributor" query, and
"Drone AI", "Drone Software" and "Drone Analytics" are three words for
something a careless prompt would treat as one thing.

The scopes below are written to be MUTUALLY DISCRIMINATING: each names
the specific products that belong to it AND names the neighbouring
markets it must not drift into. Without that, ten runs return one list
ten times.

Per the user's decision, a genuinely multi-line company may appear in
more than one market -- a firm selling both drone AI and drone analytics
belongs in both. What the scopes prevent is the SAME generic drone
reseller filling all ten.

Each entry is (slug, market_name, product_scope).
"""
from __future__ import annotations

MARKETS: tuple[tuple[str, str, str], ...] = (
    (
        "uav_systems",
        "Global UAV Systems Market - Distributors Only",
        'PRODUCT SCOPE (required): COMPLETE UAV PLATFORMS and airframes - '
        'fixed-wing, rotary and VTOL unmanned aircraft sold as whole '
        'systems, plus ground control stations, launch/recovery equipment '
        'and airframe spares. EXCLUDE companies selling only cameras, '
        'sensors, software, analytics services or counter-drone gear with '
        'no airframe line. Name the UAV platforms it distributes.',
    ),
    (
        "commercial_drones",
        "Global Commercial Drones Market - Distributors Only",
        'PRODUCT SCOPE (required): COMMERCIAL and INDUSTRIAL UAV '
        'applications - drones for civil and industrial use: '
        'survey and mapping, agriculture spraying, inspection, delivery, '
        'cinematography, public safety. Brands like DJI Enterprise, '
        'Autel, Parrot, senseFly, Wingtra. EXCLUDE defence/military-only '
        'suppliers, counter-drone vendors, and pure software or analytics '
        'firms. Name the commercial drone brands it distributes.',
    ),
    (
        "defense_drones",
        "Global Defense Drones Market - Distributors Only",
        'PRODUCT SCOPE (required): DEFENCE UAVs, MILITARY applications and '
        'HOMELAND SECURITY - ISR platforms, tactical and loitering '
        'munitions, military-grade payloads, and UAS supplied to armed '
        'forces, police, border, paramilitary and homeland-security '
        'agencies. EXCLUDE consumer/commercial-only drone sellers, '
        'counter-drone specialists and software-only firms. Name the '
        'defence or security UAS it distributes.'
    ),
    (
        "anti_drone",
        "Global Anti-Drone Systems Market - Distributors Only",
        'PRODUCT SCOPE (required): COUNTER-UAS, ANTI-DRONE and DRONE '
        'DETECTION technologies - detection radar, RF and acoustic '
        'sensors, drone-detection systems, RF jammers, spoofers, net '
        'capture, directed-energy and airspace-security systems. Brands '
        'like Dedrone, DroneShield, Robin Radar, MyDefence. EXCLUDE '
        'companies selling drones with no counter-drone or detection '
        'line. Name the C-UAS or detection products it distributes.'
    ),
    (
        "drone_robotics",
        "Global Drone Robotics Market - Distributors Only",
        'PRODUCT SCOPE (required): ROBOTICS, AUTONOMOUS SYSTEMS and '
        'UNMANNED GROUND SYSTEMS - unmanned ground vehicles (UGV), ground '
        'robots, robotic arms and manipulators, docking stations, '
        'drone-in-a-box systems, swarm hardware, flight controllers, '
        'actuators and propulsion. EXCLUDE sellers of complete ready-to-fly '
        'aircraft only, and software or imaging firms. Name the robotics '
        'or unmanned systems it distributes.'
    ),
    (
        "drone_ai",
        "Global Drone AI Market - Distributors Only",
        'PRODUCT SCOPE (required): AI SOFTWARE, INTELLIGENT AUTOMATION and '
        'AUTONOMOUS DECISION-MAKING for drones - onboard AI compute and '
        'edge inference hardware, computer-vision autonomy stacks, object '
        'detection and tracking, autonomous navigation in GPS-denied '
        'environments. EXCLUDE generic flight-planning software, '
        'BI/analytics dashboards and hardware sellers with no AI line. '
        'Name the AI products or platforms it distributes.'
    ),
    (
        "drone_software",
        "Global Drone Software Market - Distributors Only",
        'PRODUCT SCOPE (required): UAV SOFTWARE, APPLICATION DEVELOPMENT '
        'and CONTROL PLATFORMS - flight planning and mission control, '
        'ground control station software, fleet management, UTM/airspace '
        'management, photogrammetry and 3D reconstruction, plus custom '
        'UAV/GIS application development and software implementation. '
        'Resellers of Pix4D, DroneDeploy, UgCS, Auterion or Skyward '
        'qualify, and so do firms that build or deploy drone software for '
        'customers. EXCLUDE hardware-only distributors, drone-as-a-service '
        'and survey/mapping contractors that sell flights rather than '
        'software. Name the software it resells or develops.',
    ),
    (
        "drone_analytics",
        "Global Drone Analytics Market - Distributors Only",
        'PRODUCT SCOPE (required): drone DATA PROCESSING, ANALYTICS and '
        'INTELLIGENCE APPLICATIONS - processing '
        'captured imagery into insight: volumetric and stockpile '
        'measurement, crop health and NDVI, asset-inspection defect '
        'reporting, change detection, digital twins and analytics '
        'dashboards. EXCLUDE flight-control software, AI chips and '
        'hardware sellers. Name the analytics products it distributes.',
    ),
    (
        "drone_sensors",
        "Global Drone Sensors Market - Distributors Only",
        'PRODUCT SCOPE (required): SENSORS, DETECTION SYSTEMS and ONBOARD '
        'SENSING technologies for drones - LiDAR, hyperspectral, '
        'thermal/IR, '
        'magnetometers, gas and radiation detectors, SAR, RTK/GNSS and '
        'IMU modules. Brands like Velodyne, Riegl, YellowScan, MicaSense, '
        'Headwall. EXCLUDE complete-aircraft sellers and software firms. '
        'Name the sensor payloads it distributes.',
    ),
    (
        "drone_imaging",
        "Global Drone Imaging Market - Distributors Only",
        'PRODUCT SCOPE (required): CAMERAS, AERIAL IMAGING, MULTISPECTRAL '
        'and THERMAL IMAGING for drones - gimbal-stabilised RGB and '
        'cinema cameras, multispectral imaging, zoom and low-light '
        'payloads, FLIR/thermal imaging cameras, and aerial imaging '
        'services sold as a product. Brands like Zenmuse, Sony Airpeak, '
        'FLIR Vue, Phase One. EXCLUDE LiDAR/spectral sensor specialists '
        'and analytics firms. Name the imaging payloads it distributes.',
    ),
)

# Each market's prompt tells the model which SIBLING markets exist, so it
# does not answer "UAV Systems" with a counter-drone vendor. Overlap is
# allowed for genuinely multi-line companies; this only stops drift.
SIBLINGS = {
    "uav_systems": "commercial drones, defence drones",
    "commercial_drones": "defence drones, UAV platforms",
    "defense_drones": "commercial drones, counter-drone systems",
    "anti_drone": "drone platforms, drone sensors",
    "drone_robotics": "complete UAV platforms, drone software",
    "drone_ai": "drone software, drone analytics",
    "drone_software": "drone AI, drone analytics",
    "drone_analytics": "drone software, drone AI",
    "drone_sensors": "drone imaging, drone robotics",
    "drone_imaging": "drone sensors, drone analytics",
}


def by_slug(slug: str) -> tuple[str, str, str]:
    for entry in MARKETS:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown market slug: {slug}. "
                   f"Valid: {', '.join(m[0] for m in MARKETS)}")
