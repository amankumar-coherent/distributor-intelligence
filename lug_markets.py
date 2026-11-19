"""India Luggage Delivery Service Market -- ONE market, India-HQ only,
luggage delivery SERVICE PROVIDERS (user, 2026-11-05; cap 300, 10 browsers).

Inputs: the user's market definition, included/excluded scope table and nine
MANDATORY companies (Fly my Luggage, Campus2home, Shipkar, Avaan Excess, Blue
Dart, Delhivery, DTDC, Om Packers and Movers, Moving Force India Packers &
Movers). The market is PAID logistics services that collect, transport and
deliver travellers' personal luggage (suitcases, excess baggage, sports
equipment, student belongings, business travel baggage) door to door,
airport / hotel / railway-station transfers, domestic and international
(with an Indian pickup or delivery point), by road, air, rail or multimodal.

Company type = "service" (discovery.DEF_SERVICE): the companies PERFORM the
luggage delivery. The mandatory list shows that general couriers (Blue Dart,
Delhivery, DTDC) and packers & movers (Om, Moving Force) count when they run a
dedicated luggage / excess-baggage / student-baggage shipping service, so the
scope and VERIFY_RULES say so explicitly (the service rule would otherwise drop
them as "one activity among many").

The user's EXCLUSIONS applied to companies: airlines and their checked or
excess baggage fees, passenger transport and ticketing, plain parcel /
e-commerce couriers with no luggage service, commercial freight and bulk
cargo carriers, standalone luggage storage / packaging / insurance, household
relocation (furniture, appliances, full house moves) without a luggage
service, international shippers with no Indian pickup or delivery point.

Searched: first a nationwide round that names the mandatory companies, then
state by state, then cities. The scope rides in ONE Google URL (hard 7,800):
no semicolons; measure with scripts/preflight_prompt.py --batch lug.
"""
from __future__ import annotations

from india_cities import INDIA_CITIES
from india_states import INDIA_STATES

MANDATORY_LUG: tuple[str, ...] = (
    "Fly my Luggage", "Campus2home", "Shipkar", "Avaan Excess", "Blue Dart",
    "Delhivery", "DTDC", "Om Packers and Movers",
    "Moving Force India Packers & Movers",
)

MARKETS_LUG: tuple[tuple[str, str, str, str], ...] = (
    (
        "lug_india",
        "India Luggage Delivery Service Market",
        "SERVICE SCOPE (required): PAID LUGGAGE DELIVERY in India - the "
        "company collects, transports and delivers travellers' personal "
        "luggage, suitcases, excess baggage, sports equipment, student "
        "belongings or business travel baggage, door to door, airport to "
        "home or hotel, hotel to airport, railway station pickup and "
        "delivery, domestic or international shipping with a pickup or "
        "delivery point in India, by road, air, rail or multimodal, booked "
        "online or offline. Dedicated luggage shippers qualify, and so do "
        "couriers, logistics firms and packers and movers that offer a "
        "distinct luggage, excess baggage or student baggage delivery "
        "service, for example Fly my Luggage, Campus2home, Shipkar, Avaan "
        "Excess, Blue Dart, Delhivery, DTDC, Om Packers and Movers, Moving "
        "Force India Packers and Movers. EXCLUDE airlines and their checked "
        "or excess baggage fees, passenger transport and ticketing, parcel "
        "or e-commerce couriers with no luggage service, freight and bulk "
        "cargo carriers, luggage storage, packaging or insurance only, "
        "household shifting of furniture and appliances only.",
        "service",
    ),
)

SIBLINGS_LUG: dict[str, str] = {}

# A nationwide round naming the mandatory companies first, then the 26
# states / UT groups, then the 43 cities.
LUG_REGIONS: tuple[str, ...] = (
    "India nationwide, incl. Fly my Luggage, Campus2home, Shipkar, Avaan "
    "Excess, Om Packers and Movers, Moving Force India",
) + tuple(INDIA_STATES) + tuple(INDIA_CITIES)

DEFINITIONS: dict[str, str] = {
    "lug_india": (
        "The India Luggage Delivery Service Market comprises paid logistics "
        "services that collect, transport, and deliver travelers' personal "
        "luggage, suitcases, excess baggage, sports equipment, student "
        "belongings, and business travel baggage between specified "
        "locations within India or between India and international "
        "destinations. The market covers luggage delivery through road, "
        "air, rail, and multimodal transportation, including doorstep "
        "pickup, airport transfers, railway station transfers, and "
        "customized delivery arrangements. "
        "INCLUDED: paid door-to-door luggage pickup and delivery services; "
        "airport-to-home, airport-to-hotel, and hotel-to-airport luggage "
        "transfers; railway station-related luggage collection and "
        "delivery; domestic and international luggage shipping involving "
        "India; student, corporate, sports, and excess baggage delivery; "
        "road, air, rail, and multimodal luggage transportation; standard "
        "and customized luggage delivery services; online and offline "
        "luggage delivery bookings. "
        "EXCLUDED: regular airline checked baggage included in passenger "
        "tickets; standard airline excess baggage fees without a separate "
        "delivery service; personal luggage carried directly by passengers "
        "without a separate delivery service; international luggage "
        "deliveries with no Indian pickup or delivery point; commercial "
        "freight, bulk industrial cargo, and regular e-commerce parcel "
        "deliveries; passenger transportation, airline tickets, and railway "
        "tickets; standalone luggage storage, packaging, and insurance "
        "services; household relocation services covering furniture, "
        "appliances, and complete household moves. "
        "COMPANY TYPE IN THIS LIST - LUGGAGE DELIVERY SERVICE PROVIDER: a "
        "company headquartered in India that performs paid luggage pickup "
        "and delivery for travellers - dedicated luggage and excess-baggage "
        "shippers, plus couriers, logistics companies and packers & movers "
        "that offer a distinct luggage / baggage delivery service. Mandatory "
        "companies: Fly my Luggage, Campus2home, Shipkar, Avaan Excess, Blue "
        "Dart, Delhivery, DTDC, Om Packers and Movers, Moving Force India "
        "Packers & Movers."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; the whole
# criteria block must stay under ~1,000 chars, so this is tight).
VERIFY_RULES: dict[str, str] = {
    "lug_india": (
        "KEEP India firms doing PAID luggage/baggage pickup+delivery: "
        "luggage shippers, airport/rail/hotel transfers, student or excess "
        "baggage, and couriers or packers & movers with a luggage service "
        "(Blue Dart, DTDC, Delhivery).\n"
        "DROP airlines, ticketing, parcel-only couriers, freight-only, "
        "storage-only, household-shifting-only movers."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for).
EVIDENCE: dict[str, str] = {
    "lug_india": (
        r"luggage|baggage|suitcase|\bbags?\b|excess|student|courier|"
        r"packers|movers|shifting|relocat|shipping|ship |delivery|pickup|"
        r"pick-up|airport|railway|logistic|parcel|express"
    ),
}


def by_slug_lug(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_LUG:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown lug market slug: {slug}")
