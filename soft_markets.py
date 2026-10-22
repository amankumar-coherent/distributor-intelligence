"""Global Soft Serve Ice Cream Machines Market -- ONE market, worldwide,
DISTRIBUTORS / IMPORTERS / FOODSERVICE EQUIPMENT DEALERS (user, 2026-10-21).

Inputs: user's definition + inclusion/exclusion table, and
"Report_Sample_Global_Soft Serve Ice Cream Machines Market_2026-2033_CMI.pptx"
(Downloads). Its company profile separates Soft-Serve Machine Manufacturers /
OEMs (Taylor, Carpigiani, Nissei, Electro Freeze, Stoelting, Spaceman,
Donper, ICETRO, Cool Snow, Oceanpower, PASMO, Wellice), Distributors /
Importers / Foodservice Equipment Dealers (Clark Associates, TriMark, Edward
Don, Singer, Wasserstrom, KaTom, Johnson-Lancaster, Boelter, BL Korea,
Icestar, Jubang Bank, KNS Korea, Semi Corporation, THE-ICE) and Refrigeration
& Component Suppliers (Danfoss, Copeland, BITZER, Secop, Nidec, Tecumseh...).
The DISTRIBUTOR / IMPORTER / DEALER group is collected; the user stressed the
exclusions must never come through, so they are stated in the scope AND in
verify_distributors_ai_mode's rule for this market.

Scope rides in ONE Google URL (budget 7,550 / hard 7,800); measure with
`python run_soft_market.py --slug soft_distributor --region "United Arab
Emirates" --show-prompt`.
"""
from __future__ import annotations

MARKETS_SOFT: tuple[tuple[str, str, str, str], ...] = (
    (
        "soft_distributor",
        "Global Soft Serve Ice Cream Machines Market",
        "PRODUCT SCOPE (required): COMMERCIAL SOFT-SERVE ICE CREAM MACHINES - "
        "gravity-fed and pump-fed, single, twist and multi-flavor, countertop "
        "and floor-standing machines for soft serve, frozen yogurt and frozen "
        "custard, sold to QSRs, dessert shops, cafes and hotels. Brands "
        "include Taylor, Carpigiani, Electro Freeze, Stoelting, Spaceman, "
        "Nissei, Donper, ICETRO. IMPORTERS, DISTRIBUTORS and FOODSERVICE "
        "EQUIPMENT DEALERS of these machines. EXCLUDE the machine makers, "
        "component and refrigeration suppliers, hard ice cream, batch freezer "
        "and home machine sellers, beverage dispensers, mix suppliers and "
        "repair-only firms. Name the soft-serve brands it sells.",
        "distributor",
    ),
)

SIBLINGS_SOFT: dict[str, str] = {}

_REPORT_COUNTRIES = (
    "United States", "Canada", "Germany", "United Kingdom", "France", "Italy",
    "Spain", "Russia", "China", "Japan", "South Korea", "India", "Australia",
    "Singapore", "Malaysia", "Indonesia", "Thailand", "Vietnam", "Philippines",
    "United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Oman",
    "Bahrain", "South Africa", "Brazil", "Mexico", "Argentina",
)
_MORE = (
    "Netherlands", "Belgium", "Switzerland", "Poland", "Turkey", "Ireland",
    "Sweden", "Denmark", "Austria", "Portugal", "Czech Republic", "Greece",
    "Romania", "Israel", "Egypt", "Morocco", "Nigeria", "Kenya", "Taiwan",
    "Hong Kong", "New Zealand", "Chile", "Colombia", "Peru", "Pakistan",
    "Bangladesh", "Sri Lanka",
)
SOFT_REGIONS: tuple[str, ...] = _REPORT_COUNTRIES + _MORE

DEFINITIONS: dict[str, str] = {
    "soft_distributor": (
        "The Global Soft Serve Ice Cream Machines Market comprises the "
        "manufacturing and commercial sale of specialized refrigeration-based "
        "dispensing equipment designed to produce, freeze, aerate, maintain, "
        "and dispense soft-serve ice cream and compatible frozen desserts, "
        "including frozen yogurt, sorbet, non-dairy frozen desserts, and "
        "frozen custard. The market covers gravity-fed and pump or "
        "pressure-fed machines across single-flavor, dual-flavor, twist, and "
        "multi-flavor configurations, including countertop, floor-standing, "
        "and mobile or portable commercial equipment. These machines are "
        "primarily purchased by quick-service restaurants, ice cream and "
        "dessert outlets, cafes, convenience stores, hotels, entertainment "
        "venues, and other commercial foodservice establishments. Market "
        "measurement boundary: the sales value of complete soft-serve "
        "machines, including factory-integrated components and standard "
        "equipment supplied with the machines. Manufacturers, distributors, "
        "and equipment dealers form the commercial supply ecosystem, while "
        "refrigeration and component suppliers represent the upstream supply "
        "chain. "
        "INCLUDED: commercial soft-serve and compatible frozen-dessert "
        "dispensing machines; gravity-fed and pump/pressure-fed; single, dual, "
        "twist/triple and multi-flavor; countertop, floor-standing and "
        "commercial mobile/portable; air- and water-cooled; low, medium and "
        "high capacity (liters per hour); soft serve, frozen yogurt, sorbet, "
        "non-dairy, frozen custard and specialty frozen desserts; QSRs, dessert "
        "cafes, convenience stores, ice cream chains, independent dessert "
        "shops, full-service restaurants, hotels, resorts, entertainment "
        "venues, institutional foodservice; direct manufacturer sales, "
        "importers, distributors, foodservice equipment dealers, franchise "
        "procurement, online equipment sales, institutional procurement; new, "
        "replacement and upgrade purchases and rental/leasing supply; "
        "components integrated into complete machines. Geography: North "
        "America, Europe, Asia Pacific, Middle East and Africa, Latin America. "
        "EXCLUDED: hard ice cream production machines, batch freezers and "
        "domestic ice cream makers; refrigeration equipment without "
        "soft-serve freezing and dispensing; standalone beverage dispensers; "
        "domestic countertop appliances and noncommercial recreational "
        "equipment; standalone chillers, freezers and cooling systems; "
        "industrial bulk ice cream lines and continuous industrial freezers; "
        "hard-packed ice cream, gelato from separate batch equipment and "
        "conventional frozen beverages; residential consumers, retail sales of "
        "packaged frozen desserts and industrial frozen dessert plants; "
        "ingredients, consumables and unrelated foodservice equipment; "
        "standalone maintenance contracts, repair services and unrelated "
        "rental fees; independent aftermarket and standalone refrigeration "
        "component sales; double counting of manufacturer and distributor "
        "resale. "
        "COMPANY TYPE IN THIS LIST - SOFT-SERVE MACHINE DISTRIBUTOR / "
        "IMPORTER / FOODSERVICE EQUIPMENT DEALER: a company that buys "
        "commercial soft-serve machines from the manufacturers (e.g. Taylor, "
        "Carpigiani, Electro Freeze, Stoelting, Spaceman, Nissei, Donper, "
        "ICETRO, PASMO) and resells, rents or leases them to commercial "
        "foodservice customers - importers, authorized distributors and "
        "foodservice equipment dealers (e.g. TriMark, Edward Don, Singer, "
        "Wasserstrom, KaTom, Semi Corporation, THE-ICE, Icestar). EXCLUDED "
        "companies: soft-serve machine manufacturers / OEMs and their own "
        "sales subsidiaries; refrigeration and component suppliers (e.g. "
        "Danfoss, Copeland, BITZER, Secop, Tecumseh); sellers of hard ice "
        "cream, batch freezer, gelato, industrial or domestic ice cream "
        "machines only; beverage, slush and frozen-drink dispenser sellers; "
        "standalone refrigeration sellers; soft-serve mix, ingredient and "
        "consumable suppliers; repair, maintenance and service-only firms; "
        "consumer retail stores."
    ),
}


def by_slug_soft(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_SOFT:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown soft market slug: {slug}")

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified company whose profile never matches this goes to
# flagged_for_review.xlsx instead of the main Excel. Tested on the delivered
# list 2026-10-21 (all genuine rows match).
EVIDENCE: dict[str, str] = {
    "soft_distributor": r"soft[- ]?serve|frozen yog|soft ice|taylor|carpigiani|electro ?freeze|stoelting|spaceman|nissei|donper|icetro|pasmo|wellice|oceanpower|gel ?matic|frozen dessert machine|frozen custard|ice cream machine|мягкого мороженого|фризер",
}
