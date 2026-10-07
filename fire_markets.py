"""25 Fire Safety Ecosystem markets, India-only.

Built from the user's "Fire Ecosystem Industry Analysis_Title & Taxonomy.xlsx":
the Home tab gives each market's name and one-line "Exhibitor Function /
Relevance" definition, and the Taxonomy tab gives a detailed breakdown
(By System/Product/Technology/Application/End-Use/Distribution
Channel/Region) showing that most markets share a large surface with their
closest neighbours -- Fire Protection, Fire Suppression and Passive Fire
Protection all touch "fire-resistant" and "systems" vocabulary, Fire
Detection and Gas & Flame Detection both cover detectors and alarms, Life
Safety/Emergency Communications/Smoke Control Systems all appear under
each other's End-Use and Application branches. Without an explicit
EXCLUDE clause naming the neighbour, a generic query for any one of these
returns the same companies as its siblings -- this is exactly the
drone_mapping/drone_lidar/drone_navigation overlap problem from the
drone market batches, so every scope below names both what belongs and
which of its closest siblings it must not drift into.

CLASSIFICATION (4-way distributor rule, same system as drone_markets2.py):

  DISTRIBUTOR (12) -- physical goods, stocking/reselling hardware:
    fire_protection, fire_detection, fire_suppression, sprinkler_systems,
    fire_pumps, passive_fire_protection, fire_doors, gas_flame_detection,
    electrical_fire_safety, firefighting_equipment, fire_vehicles,
    rescue_equipment

  HYBRID (9) -- integrated systems / hardware+software, channel
  partners and system integrators count the same as a stocking
  distributor:
    industrial_fire_safety, battery_fire_safety,
    critical_infrastructure_fire_safety, life_safety,
    emergency_communications, smoke_control_systems, smart_fire_safety,
    fire_ai_robotics, fire_drones_gis

  SERVICE (4) -- performed for a paying customer, not sold as a stocked
  good; two of these (disaster_response, fire_safety_training) have no
  "By Distribution Channel" row at all in their own taxonomy, itself a
  sign these are not goods-distribution markets:
    fire_prevention, disaster_response, fire_safety_training,
    fire_safety_testing

Each entry is (slug, market_name, product_scope, kind) -- identical shape
to drone_markets2.MARKETS2, so the same runner/supervisor/verifier code
can drive both batches once market_batches.py generalizes over them.
"""
from __future__ import annotations

MARKETS_FIRE: tuple[tuple[str, str, str, str], ...] = (
    (
        "fire_protection",
        "Fire Protection Market - India",
        'PRODUCT SCOPE (required): FIRE PROTECTION SYSTEMS AND EQUIPMENT - '
        'standalone, integrated and specialized fire protection systems; '
        'fire extinguishers, sprinkler systems, fire pumps, hydrant '
        'systems, hoses and nozzles, fire monitors, alarm panels and '
        'devices. Includes conventional and smart/connected systems. This '
        'is the BROAD protection-systems market; EXCLUDE companies that '
        'only sell fire suppression AGENTS (that is Fire Suppression), '
        'only passive materials like coatings/barriers (Passive Fire '
        'Protection), or only sprinklers/pumps as a single-product '
        'specialist (those are their own markets). Name the fire '
        'protection products it distributes.',
        "distributor",
    ),
    (
        "fire_prevention",
        "Fire Prevention Market - India",
        'SERVICE SCOPE (required): FIRE RISK ASSESSMENT AND PREVENTION '
        'SERVICES - fire risk assessment and management, hazard '
        'identification, fire prevention and mitigation planning, fire '
        'safety management, compliance and preventive-technology '
        'consulting. This is a RISK/CONSULTING service, not equipment '
        'sales - EXCLUDE pure equipment distributors with no assessment/'
        'consulting offering, and EXCLUDE Fire Safety Training providers '
        '(a separate market). Name the prevention/risk service it '
        'performs.',
        "service",
    ),
    (
        "fire_detection",
        "Fire Detection Market - India",
        'PRODUCT SCOPE (required): FIRE DETECTION HARDWARE - smoke '
        'detectors (point, beam, aspirating), heat detectors, flame '
        'detectors, multi-criteria and multi-sensor detectors, fire alarm '
        'control panels and detection devices. EXCLUDE gas-only or '
        'toxic-gas detection specialists (Gas & Flame Detection), and '
        'EXCLUDE companies selling only full suppression/sprinkler systems '
        'with no standalone detection product line. Name the detection '
        'hardware it distributes.',
        "distributor",
    ),
    (
        "fire_suppression",
        "Fire Suppression Market - India",
        'PRODUCT SCOPE (required): FIRE SUPPRESSION SYSTEMS AND AGENTS - '
        'water-based, foam-based, gas-based, dry-chemical and wet-chemical '
        'suppression systems; automatic and manual suppression systems; '
        'suppression agents (water, foam, dry/wet chemicals, clean '
        'agents). EXCLUDE sprinkler-only specialists (Sprinkler Systems), '
        'fire-pump-only specialists (Fire Pumps), and passive/structural '
        'fire protection with no active suppression agent. Name the '
        'suppression system or agent it distributes.',
        "distributor",
    ),
    (
        "sprinkler_systems",
        "Sprinkler Systems Market - India",
        'PRODUCT SCOPE (required): SPRINKLER SYSTEMS AND HYDRANTS - wet '
        'pipe, dry pipe, pre-action, deluge and water-mist sprinkler '
        'systems; conventional, addressable, smart and IoT-enabled '
        'sprinkler systems; fire hydrants and water spray/deluge systems. '
        'EXCLUDE general fire suppression agent sellers with no sprinkler '
        'hardware (Fire Suppression) and EXCLUDE fire-pump-only '
        'specialists (Fire Pumps). Name the sprinkler/hydrant products it '
        'distributes.',
        "distributor",
    ),
    (
        "fire_pumps",
        "Fire Pumps Market - India",
        'PRODUCT SCOPE (required): FIRE PUMPS AND PUMPING SYSTEMS - '
        'electric, diesel and jockey fire pumps; vertical turbine, '
        'horizontal split-case and end-suction pumps; single, multi-pump, '
        'duty & standby and packaged fire pump systems, across low to '
        'very-high capacity. EXCLUDE sprinkler-system-only sellers with no '
        'standalone pump product line (Sprinkler Systems). Name the pump '
        'products it distributes.',
        "distributor",
    ),
    (
        "passive_fire_protection",
        "Passive Fire Protection Market - India",
        'PRODUCT SCOPE (required): PASSIVE FIRE PROTECTION MATERIALS - '
        'fire-resistant barriers, fire compartmentation and firestopping '
        'products, structural fire protection, fire-resistant coatings, '
        'firestop sealants and mortars, fire-resistant boards/panels/'
        'insulation, fire-rated glass. Intumescent, cementitious and '
        'encapsulation technologies. EXCLUDE Fire Doors specialists (a '
        'separate market even though fire-rated doors appear in both '
        'taxonomies) and EXCLUDE active suppression/detection equipment '
        'sellers. Name the passive fire protection product it '
        'distributes.',
        "distributor",
    ),
    (
        "fire_doors",
        "Fire Doors Market - India",
        'PRODUCT SCOPE (required): FIRE-RATED DOORS AND BARRIERS - '
        'single-leaf, double-leaf, sliding, rolling, glass and acoustic '
        'fire doors; emergency exit fire doors; steel, timber, aluminum, '
        'glass and composite fire doors. EXCLUDE general passive-'
        'protection material sellers with no door product line (Passive '
        'Fire Protection). Name the fire door product it distributes.',
        "distributor",
    ),
    (
        "industrial_fire_safety",
        "Industrial Fire Safety Market - India",
        'PRODUCT/SERVICE SCOPE (required): INTEGRATED INDUSTRIAL FIRE '
        'SAFETY SOLUTIONS - companies that supply or integrate fire '
        'detection & alarm, suppression, sprinkler, hydrant, gas/flame '
        'detection, passive protection and smoke control systems '
        'SPECIFICALLY for industrial facilities (manufacturing, '
        'processing, chemical plants, storage & warehousing, utility & '
        'power). System integrators and turnkey project suppliers count '
        'here, not only single-product distributors. EXCLUDE companies '
        'that sell only ONE of these subsystems as a stand-alone market '
        'already covered elsewhere unless they are positioning it for '
        'industrial/process-safety use specifically. Name the industrial '
        'fire safety systems it supplies or integrates.',
        "hybrid",
    ),
    (
        "gas_flame_detection",
        "Gas & Flame Detection Market - India",
        'PRODUCT SCOPE (required): GAS AND FLAME DETECTION HARDWARE - '
        'combustible/toxic/oxygen gas detectors, flame detectors, '
        'multi-gas detectors; point, open-path, portable, fixed and '
        'wireless gas detectors, for hazardous-area and industrial '
        'hazardous-gas monitoring (oil & gas, chemicals, mining, marine). '
        'EXCLUDE smoke/heat detector specialists with no gas/flame '
        'product (Fire Detection). Name the gas/flame detection hardware '
        'it distributes.',
        "distributor",
    ),
    (
        "electrical_fire_safety",
        "Electrical Fire Safety Market - India",
        'PRODUCT SCOPE (required): ELECTRICAL FIRE SAFETY PRODUCTS - '
        'electrical/arc-fault fire detection, thermal monitoring for '
        'electrical faults, fire-resistant cables and wiring, fire-'
        'resistant enclosures, surge & overload protection, electrical '
        'fire suppression aimed specifically at short-circuit/overload/'
        'arc-fault/cable/transformer/motor fire hazards. EXCLUDE general '
        'fire detection or suppression sellers with no electrical-hazard-'
        'specific product. Name the electrical fire safety product it '
        'distributes.',
        "distributor",
    ),
    (
        "battery_fire_safety",
        "Battery Fire Safety Market - India",
        'PRODUCT/SERVICE SCOPE (required): BATTERY AND BESS/EV FIRE SAFETY '
        '- battery fire detection, thermal-runaway detection, battery '
        'fire suppression, explosion prevention and containment for '
        'lithium-ion/BESS/EV battery systems; cell-level, module-level, '
        'pack-level, rack-level and facility-level protection; '
        'temperature/smoke/gas/flame/thermal-imaging/cell-voltage '
        'monitoring technology specific to batteries. Channel partners '
        'and system integrators for BESS/EV OEMs count here. EXCLUDE '
        'general electrical fire safety sellers with no battery/BESS-'
        'specific product. Name the battery fire safety solution it '
        'supplies or integrates.',
        "hybrid",
    ),
    (
        "critical_infrastructure_fire_safety",
        "Critical Infrastructure Fire Safety Market - India",
        'PRODUCT/SERVICE SCOPE (required): FIRE SAFETY FOR CRITICAL '
        'INFRASTRUCTURE - integrated fire detection, suppression, '
        'sprinkler, gas/flame detection, passive protection, smoke '
        'control and emergency communication systems SPECIFICALLY for '
        'power generation & utility, transmission & distribution, '
        'telecom, transportation, water & wastewater and government/'
        'public infrastructure. System integrators and turnkey suppliers '
        'for these sectors count here. EXCLUDE general commercial/'
        'residential fire safety sellers with no critical-infrastructure/'
        'utility focus. Name the critical-infrastructure fire safety '
        'systems it supplies.',
        "hybrid",
    ),
    (
        "firefighting_equipment",
        "Firefighting Equipment Market - India",
        'PRODUCT SCOPE (required): FIREFIGHTING EQUIPMENT AND TOOLS - '
        'fire extinguishers (water, foam, dry-chemical, CO2, wet-chemical, '
        'clean-agent), fire hoses, nozzles, hydrants, monitors, hose '
        'reels, firefighting pumps, foam equipment, firefighting tools, '
        'breathing apparatus. Conventional and smart/connected equipment. '
        'EXCLUDE fixed installed sprinkler/suppression SYSTEMS (those are '
        'their own markets) and EXCLUDE Rescue Equipment (extrication/'
        'rescue tools, a separate market). Name the firefighting '
        'equipment it distributes.',
        "distributor",
    ),
    (
        "fire_vehicles",
        "Fire Vehicles Market - India",
        'PRODUCT SCOPE (required): FIRE AND RESCUE VEHICLES - fire '
        'engines, water/foam tender trucks, rescue vehicles, aerial '
        'ladder/turntable ladder trucks, rapid intervention vehicles, '
        'airport crash tenders, standard/heavy-duty/compact/multi-purpose/'
        'all-terrain/off-road/electric fire vehicles. Authorized dealers '
        'and specialized vehicle suppliers count as distributors here, '
        'not only direct OEM sales. EXCLUDE portable/handheld '
        'firefighting equipment sellers with no vehicle product line. '
        'Name the fire vehicle it distributes or supplies.',
        "distributor",
    ),
    (
        "rescue_equipment",
        "Rescue Equipment Market - India",
        'PRODUCT SCOPE (required): RESCUE AND EXTRICATION EQUIPMENT - '
        'hydraulic/pneumatic/mechanical rescue tools, cutting & spreading '
        'equipment, lifting & stabilization equipment, ropes & rope-'
        'rescue gear, rescue stretchers, evacuation equipment; manual, '
        'powered, battery-powered and robotic rescue equipment, for '
        'buildings, road, industrial, mining, construction, tunnel and '
        'marine rescue. EXCLUDE fire vehicles (Fire Vehicles) and '
        'standard firefighting hand tools already covered under '
        'Firefighting Equipment. Name the rescue equipment it '
        'distributes.',
        "distributor",
    ),
    (
        "life_safety",
        "Life Safety Market - India",
        'PRODUCT/SERVICE SCOPE (required): LIFE SAFETY AND EVACUATION '
        'SYSTEMS - integrated fire alarm, emergency lighting, emergency '
        'communication, smoke control, evacuation and fire detection/'
        'suppression systems bundled specifically for OCCUPANT life-'
        'safety and building evacuation; conventional, addressable, '
        'wireless and IoT-enabled life safety systems. System integrators '
        'for building life-safety projects count here. EXCLUDE single-'
        'product detection/suppression/smoke-control sellers already '
        'covered under their own markets, unless they sell an integrated '
        'multi-system life-safety package. Name the life safety system it '
        'supplies.',
        "hybrid",
    ),
    (
        "emergency_communications",
        "Emergency Communications Market - India",
        'PRODUCT/SERVICE SCOPE (required): EMERGENCY COMMUNICATION '
        'SYSTEMS - public address systems, mass notification systems, '
        'voice evacuation systems, emergency intercom and call systems, '
        'two-way communication and emergency alert systems; wired and '
        'wireless communication technology, for fire emergency, building '
        'evacuation, industrial emergency, disaster response and public '
        'safety/critical-incident use. EXCLUDE general life-safety system '
        'integrators with no standalone communication product, and '
        'EXCLUDE Disaster Response service providers. Name the emergency '
        'communication system it distributes.',
        "hybrid",
    ),
    (
        "disaster_response",
        "Disaster Response Market - India",
        'SERVICE SCOPE (required): DISASTER RESPONSE AND EMERGENCY '
        'RESPONSE SERVICES - companies or organizations that PERFORM fire '
        '& emergency response, search & rescue, medical emergency '
        'response, hazardous material response, evacuation & shelter '
        'response, or disaster recovery response, for fire, flood, '
        'earthquake, cyclone, industrial accident, chemical incident or '
        'wildfire events. This is an operational RESPONSE SERVICE, not '
        'equipment sales - EXCLUDE pure rescue/firefighting/'
        'communication EQUIPMENT distributors with no response-service '
        'offering (those are their own markets). Name the disaster '
        'response service it performs.',
        "service",
    ),
    (
        "smoke_control_systems",
        "Smoke Control Systems Market - India",
        'PRODUCT/SERVICE SCOPE (required): SMOKE CONTROL AND VENTILATION '
        'SYSTEMS - smoke exhaust, extraction, pressurization systems; '
        'natural and mechanical smoke ventilation; combined smoke control '
        'systems; manual, automatic, centralized, zone-based and BMS-'
        'integrated smoke control, for fire evacuation, smoke containment/'
        'extraction and firefighter access. EXCLUDE general life-safety '
        'or fire-detection sellers with no standalone smoke-control/'
        'ventilation product line. Name the smoke control system it '
        'distributes or integrates.',
        "hybrid",
    ),
    (
        "fire_safety_training",
        "Fire Safety Training Market - India",
        'SERVICE SCOPE (required): FIRE SAFETY TRAINING SERVICES - '
        'companies or organizations that RUN fire prevention, fire '
        'detection & alarm, firefighting, evacuation, extinguisher-use, '
        'emergency-response or fire-risk-management TRAINING COURSES; '
        'basic through advanced/specialized/certification/refresher '
        'training; traditional, digital, VR/AR, simulator and AI-enabled '
        'training delivery. This is a TRAINING service, not equipment '
        'sales - EXCLUDE equipment distributors that merely mention '
        'training as a support activity for their own products, and '
        'EXCLUDE Fire Prevention (risk-assessment consulting, a separate '
        'market). Name the training programme it runs.',
        "service",
    ),
    (
        "smart_fire_safety",
        "Smart Fire Safety Market - India",
        'PRODUCT/SERVICE SCOPE (required): IOT-ENABLED / CONNECTED FIRE '
        'SAFETY SYSTEMS - smart fire detection, alarm, suppression, '
        'prevention, emergency communication, evacuation and fire '
        'monitoring systems using Wi-Fi/Bluetooth/Zigbee/LoRaWAN/NB-IoT/'
        '5G connectivity; real-time/remote/cloud-based monitoring, '
        'centralized/mobile control, automated alerts. Channel partners '
        'and platform resellers count here. EXCLUDE conventional (non-'
        'connected) equipment sellers already covered under their own '
        'product markets, unless they specifically sell the IoT/smart/'
        'connected version. Name the smart fire safety product or '
        'platform it distributes.',
        "hybrid",
    ),
    (
        "fire_ai_robotics",
        "Fire AI & Robotics Market - India",
        'PRODUCT/SERVICE SCOPE (required): AI-BASED FIRE DETECTION AND '
        'FIREFIGHTING ROBOTICS - AI-based fire/smoke/flame detection, '
        'predictive fire-risk analytics, AI-assisted emergency response, '
        'fire scene intelligence; firefighting robots, inspection robots, '
        'search & rescue robots, surveillance robots, hazardous-area '
        'robots, autonomous mobile robots, for industrial, oil & gas, '
        'chemical, warehouse, power-plant, high-rise and wildland '
        'environments. EXCLUDE conventional (non-AI, non-robotic) '
        'detection or firefighting equipment sellers, and EXCLUDE Fire '
        'Drones & GIS (a separate market). Name the AI/robotic fire '
        'safety product it distributes.',
        "hybrid",
    ),
    (
        "fire_drones_gis",
        "Fire Drones & GIS Market - India",
        'PRODUCT/SERVICE SCOPE (required): FIREFIGHTING DRONES AND GIS/'
        'MAPPING FOR FIRE - multirotor, fixed-wing, hybrid VTOL, heavy-'
        'lift, tethered and autonomous drones used for fire detection, '
        'monitoring, aerial surveillance, thermal imaging, search & '
        'rescue, damage assessment; fire-perimeter/hotspot/spread/risk-'
        'zone/evacuation-route/damage/infrastructure MAPPING and GIS '
        'software for wildland, urban, industrial, oil & gas, power/'
        'utility, mining and coastal fire response. Drone OEMs, GIS '
        'technology providers, system integrators and specialized fire-'
        'safety suppliers all count as distributors here. EXCLUDE '
        'general Fire AI & Robotics sellers with no drone/GIS product, '
        'and EXCLUDE ground-based rescue robots. Name the drone or GIS '
        'product it distributes.',
        "hybrid",
    ),
    (
        "fire_safety_testing",
        "Fire Safety Testing Market - India",
        'SERVICE SCOPE (required): FIRE SAFETY TESTING, INSPECTION AND '
        'CERTIFICATION SERVICES - companies that PERFORM fire-resistance, '
        'fire-detection, fire-alarm, fire-suppression, smoke & flame, '
        'fire-door or fire-extinguisher TESTING; laboratory, field, on-'
        'site, performance, destructive, non-destructive and '
        'environmental testing; product certification, building fire-'
        'safety compliance, system commissioning, periodic inspection and '
        'regulatory-compliance services. Independent testing '
        'laboratories, certification bodies and fire safety consultants '
        'count here. This is a TESTING/CERTIFICATION service, not '
        'equipment sales - EXCLUDE equipment manufacturers/distributors '
        'that only test their OWN products with no third-party testing '
        'service. Name the testing or certification service it provides.',
        "service",
    ),
)


SIBLINGS_FIRE = {
    "fire_protection": "fire suppression, passive fire protection, sprinkler systems",
    "fire_prevention": "fire safety training",
    "fire_detection": "gas & flame detection",
    "fire_suppression": "sprinkler systems, fire pumps, fire protection",
    "sprinkler_systems": "fire suppression, fire pumps",
    "fire_pumps": "sprinkler systems",
    "passive_fire_protection": "fire doors",
    "fire_doors": "passive fire protection",
    "industrial_fire_safety": "critical infrastructure fire safety",
    "gas_flame_detection": "fire detection",
    "electrical_fire_safety": "battery fire safety",
    "battery_fire_safety": "electrical fire safety",
    "critical_infrastructure_fire_safety": "industrial fire safety",
    "firefighting_equipment": "rescue equipment, fire vehicles",
    "fire_vehicles": "firefighting equipment",
    "rescue_equipment": "firefighting equipment, fire vehicles",
    "life_safety": "emergency communications, smoke control systems",
    "emergency_communications": "life safety, disaster response",
    "disaster_response": "emergency communications, rescue equipment",
    "smoke_control_systems": "life safety",
    "fire_safety_training": "fire prevention",
    "smart_fire_safety": "fire ai & robotics",
    "fire_ai_robotics": "fire drones & GIS, smart fire safety",
    "fire_drones_gis": "fire ai & robotics",
    "fire_safety_testing": "fire prevention",
}


def by_slug_fire(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_FIRE:
        if entry[0] == slug:
            return entry
    raise KeyError(
        f"unknown fire market slug: {slug}. "
        f"Valid: {', '.join(m[0] for m in MARKETS_FIRE)}"
    )
