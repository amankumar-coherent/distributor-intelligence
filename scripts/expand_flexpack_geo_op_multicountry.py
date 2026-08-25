# -*- coding: utf-8 -*-
"""
Evidence-only multi-country expansions for Continent / Geography + Operational Presence.
Only companies web-verified in this pass. Never invent countries/cities.
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "geo_op_multicountry_expand_audit.json"

COUNTRY_REGION = {
    "United States": "NA", "Canada": "NA", "Mexico": "NA", "Puerto Rico": "NA",
    "Panama": "NA",
    "Brazil": "LATAM", "Chile": "LATAM", "Colombia": "LATAM", "Peru": "LATAM",
    "United Kingdom": "EU", "Ireland": "EU", "France": "EU", "Germany": "EU",
    "Spain": "EU", "Portugal": "EU", "Italy": "EU", "Netherlands": "EU",
    "Belgium": "EU", "Luxembourg": "EU", "Switzerland": "EU", "Austria": "EU",
    "Sweden": "EU", "Norway": "EU", "Denmark": "EU", "Finland": "EU", "Poland": "EU",
    "Czechia": "EU", "Romania": "EU", "Hungary": "EU", "Greece": "EU",
    "Lithuania": "EU", "Latvia": "EU", "Estonia": "EU", "Slovakia": "EU",
    "Bulgaria": "EU", "Bosnia and Herzegovina": "EU", "Croatia": "EU",
    "Serbia": "EU", "Slovenia": "EU", "Turkey": "EU",
    "South Africa": "AF", "Namibia": "AF",
    "Japan": "APAC", "China": "APAC", "Hong Kong": "APAC", "Taiwan": "APAC",
    "South Korea": "APAC", "India": "APAC", "Singapore": "APAC", "Malaysia": "APAC",
    "Indonesia": "APAC", "Thailand": "APAC",
    "Australia": "OC", "New Zealand": "OC",
}
REGION_ORDER = ["NA", "LATAM", "EU", "AF", "ME", "APAC", "OC"]

# company -> {country: [cities...]}  empty list = country-only physical
GEO: dict[str, dict[str, list[str]]] = {
    "MonotaRO": {
        "Japan": ["Osaka [HQ]", "Tokyo", "Nagoya", "Kasama", "Inagawa"],
        "South Korea": ["Seoul"],
        "Indonesia": ["Jakarta"],
        "India": ["New Delhi", "Chennai"],
        "China": ["Shanghai"],
    },
    "Askul": {
        "Japan": [
            "Tokyo [HQ]", "Ageo / Kanto DC", "Edogawa / Tokyo DC", "Yokohama DC",
            "Sendai DC", "Nagoya DC", "Osaka / Kansai DC", "Fukuoka DC",
        ],
        "China": ["Taicang Jiangsu"],
    },
    "Richards Packaging": {
        "Canada": [
            "Mississauga ON [HQ]", "Rocky View County / Calgary AB", "Lachine QC",
            "Surrey BC", "Winnipeg MB", "Dartmouth NS", "Edmonton AB",
        ],
        "United States": [
            "Memphis TN", "Portland OR", "Roseville / Sacramento CA",
            "Santa Fe Springs CA", "Auburn / Seattle WA",
        ],
    },
    "Ratioform": {
        "Germany": ["Pliening [HQ]"],
        "Austria": ["Vienna", "Salzburg"],
        "Switzerland": ["Steinhausen / Zurich"],
        "Italy": ["Tribiano / Milan"],
    },
    "Igepa Group": {
        "Germany": ["Hamburg [HQ]"],
        "Austria": [],
        "Belgium": ["Aalter"],
        "Luxembourg": ["Niederanven"],
        "Netherlands": ["Tiel"],
        "Czechia": ["Odolena Voda"],
        "Poland": ["Krakow"],
        "Hungary": ["Nyiregyhaza"],
        "Sweden": ["Molnlycke"],
        "Norway": ["Drobak"],
        "Finland": ["Helsinki"],
        "Denmark": ["Roskilde"],
        "Switzerland": [],
        "Lithuania": [],
        "France": [],
        "Bosnia and Herzegovina": [],
        "Croatia": [],
        "Serbia": [],
        "Slovenia": [],
    },
    "Macfarlane Packaging": {
        "United Kingdom": ["Glasgow [HQ]"],
        "Ireland": [],
        "Netherlands": ["Hoofddorp"],
        "Germany": [],
        "Poland": ["Stargard"],
        "Italy": ["Palomba / Rome"],
    },
    "Abco Kovex": {
        "Ireland": ["Swords / Dublin [HQ]"],
        "United Kingdom": ["Burton Latimer / East Midlands", "Dudley"],
        "Germany": [],  # company states Distribution Centre in Germany; city NPD
    },
    "Gerdmans": {
        "Sweden": ["Markaryd [HQ]"],
        "Norway": ["Sandvika"],
        "Denmark": ["Niva"],
        "Finland": ["Espoo"],
    },
    "Atlantic Packaging Products": {
        "Canada": [
            "Scarborough ON [HQ]", "Brampton ON", "Mississauga ON", "Whitby ON",
            "Ingersoll ON", "Brockville ON", "Sudbury ON",
        ],
        "United States": [
            "Carol Stream IL", "Wood Dale IL", "Mansfield OH", "York PA",
        ],
    },
    "Hisco": {
        "United States": ["Houston TX [HQ]"],
        "Canada": ["Burlington ON"],
        "Mexico": [],  # HiscoMex multi-site; city list not fully verified here
    },
    "All American Containers": {
        "United States": [
            "Miami FL [HQ]", "Tampa FL", "Atlanta GA", "Somerset NJ",
            "Pittsburgh PA", "Chicago IL", "Houston TX", "Kalama WA",
            "Simi Valley CA", "Windsor CA", "Santa Rosa CA", "San Leandro CA",
        ],
        "Puerto Rico": ["Catano"],
        "Mexico": [],
        "Panama": [],
        "Taiwan": [],
    },
    "Crawford Packaging": {
        "Canada": [
            "London ON [HQ]", "Brampton ON", "Cambridge ON", "Essex ON",
            "Mississauga ON", "Winnipeg MB", "Calgary AB",
        ],
        "United States": ["Laredo TX", "Valdosta GA"],
        "Mexico": [],
    },
    "Papyrus": {
        "Sweden": ["Molndal [HQ]"],
        "Netherlands": ["Wijchen"],
        "Poland": ["Warsaw"],
        "Lithuania": ["Kaunas"],
        "Hungary": ["Budapest"],
        "Latvia": ["Riga"],
        "Estonia": ["Tallinn"],
        "Belgium": ["Groot-Bijgaarden"],
        "Romania": ["Bucharest"],
        "Switzerland": ["Thalwil"],
        "Czechia": ["Prague"],
        "Denmark": ["Vejle"],
        "Norway": ["Oslo"],
        "Finland": ["Vantaa"],
    },
    "Viking Direct": {
        "Ireland": ["Dublin [HQ]"],
        "United Kingdom": ["Leicester"],
        "Netherlands": ["Venlo"],
        "Germany": [],
        "Austria": [],
        "Switzerland": [],
        "Belgium": [],
        "Luxembourg": [],
    },
    "Seow Khim Polythene": {
        "Singapore": ["Pasir Ris / Loyang [HQ]"],
        "Malaysia": [],
        "United States": ["Caledonia MI"],
    },
    "Brangs + Heinrich": {
        "Germany": [
            "Solingen [HQ]", "Hamburg", "Hannover", "Berlin", "Dresden",
            "Stuttgart", "Wiesbaden", "Augsburg",
        ],
        "Netherlands": ["Utrecht"],
        "Belgium": ["Dison"],
        "China": [],
    },
    "Blackwoods": {
        "Australia": ["Macquarie Park NSW [HQ]"],
    },
    "Spicers Paper": {
        "Canada": [
            "Vaughan ON [HQ]", "Calgary AB", "Dartmouth NS", "Richmond BC",
            "Winnipeg MB",
        ],
    },
    "Novvia Group": {
        "United States": ["Franklin TN [HQ]"],
        "Canada": ["Mississauga ON"],
        "Netherlands": ["Tilburg"],
        "China": ["Qingdao"],
        "Singapore": ["Jurong"],
    },
    "Papier-Mettler": {
        "Germany": ["Morbach [HQ]"],
        "Belgium": ["Heusden-Zolder"],
        "Chile": ["Santiago"],
        "Denmark": ["Nyborg"],
        "United Kingdom": ["Stevenage"],
        "France": ["Boulogne-Billancourt"],
        "Italy": ["Vimodrone"],
        "Austria": ["Wolkersdorf"],
        "Poland": ["Wroclaw"],
        "Sweden": ["Kvidinge"],
        "Switzerland": ["Rotkreuz"],
        "Hungary": ["Tata"],
        "Spain": ["Cudillero"],
        "Czechia": ["Prague"],
        "United States": ["Raynham MA"],
        "Australia": ["Notting Hill VIC"],
    },
    "Tantraco": {
        "Singapore": ["Ang Mo Kio [HQ]"],
        "Thailand": ["Bangkok area"],
    },
    "Waltons": {
        "South Africa": [],
        "Namibia": [],
    },
    "Packaging Direct B.V.": {
        "Netherlands": ["Apeldoorn [HQ]"],
    },
    "Furodet Emballages": {
        "France": ["Sennecey-le-Grand [HQ]", "Ormes"],
    },
    "PRODINGER Verpackung": {
        "Germany": [
            "Coburg [HQ]", "Erfurt", "Herbolzheim", "Nurnberg",
            "Henstedt-Ulzburg", "Rodental",
        ],
    },
}

OP: dict[str, list[str]] = {
    "MonotaRO": ["Japan", "South Korea", "Indonesia", "India"],
    "Askul": ["Japan"],
    "Richards Packaging": ["Canada", "United States"],
    "Ratioform": ["Germany", "Austria", "Switzerland", "Italy"],
    "Igepa Group": [
        "Germany", "Austria", "Belgium", "Luxembourg", "Netherlands", "Czechia",
        "Poland", "Hungary", "Sweden", "Norway", "Finland", "Denmark",
        "Switzerland", "Lithuania", "France", "Bosnia and Herzegovina", "Croatia",
        "Serbia", "Slovenia",
    ],
    "Macfarlane Packaging": [
        "United Kingdom", "Ireland", "Netherlands", "Germany", "Poland", "Italy",
    ],
    "Abco Kovex": ["Ireland", "United Kingdom", "Germany"],
    "Gerdmans": ["Sweden", "Norway", "Denmark", "Finland"],
    "Atlantic Packaging Products": ["Canada", "United States"],
    "Hisco": ["United States", "Canada", "Mexico"],
    "All American Containers": [
        "United States", "Puerto Rico", "Mexico", "Panama", "Taiwan",
    ],
    "Crawford Packaging": ["Canada", "United States", "Mexico"],
    "Papyrus": [
        "Sweden", "Netherlands", "Poland", "Lithuania", "Hungary", "Latvia",
        "Estonia", "Belgium", "Romania", "Switzerland", "Czechia", "Denmark",
        "Norway", "Finland",
    ],
    "Viking Direct": [
        "Ireland", "United Kingdom", "Netherlands", "Germany", "Austria",
        "Switzerland", "Belgium", "Luxembourg",
    ],
    "Seow Khim Polythene": ["Singapore", "Malaysia", "United States"],
    "Brangs + Heinrich": ["Germany", "Netherlands", "Belgium", "China"],
    "Blackwoods": ["Australia", "New Zealand"],
    "Spicers Paper": ["Canada"],
    "Novvia Group": [
        "United States", "Canada", "Netherlands", "China", "Singapore",
    ],
    "Papier-Mettler": [
        "Germany", "Belgium", "Chile", "Denmark", "United Kingdom", "France",
        "Italy", "Austria", "Poland", "Sweden", "Switzerland", "Hungary",
        "Spain", "Czechia", "United States", "Australia",
    ],
    "Tantraco": ["Singapore", "Thailand"],
    "Waltons": ["South Africa", "Namibia"],
    "Packaging Direct B.V.": [
        "Netherlands", "Belgium", "Austria", "Croatia", "Czechia", "France",
        "Hungary", "Italy", "Poland", "Slovakia", "Slovenia", "United Kingdom",
        "Bulgaria", "Denmark", "Estonia", "Finland", "Greece", "Ireland",
        "Latvia", "Lithuania", "Portugal", "Romania", "Spain", "Sweden",
    ],
    "Furodet Emballages": [
        "France", "Germany", "Belgium", "Luxembourg", "Netherlands", "Austria",
    ],
    "PRODINGER Verpackung": ["Germany"],
}


def fmt_geo(presence: dict[str, list[str]]) -> str:
    by_reg: dict[str, list[str]] = {}
    for country, cities in presence.items():
        reg = COUNTRY_REGION.get(country, "APAC")
        if cities:
            by_reg.setdefault(reg, []).append(f"{country}: {', '.join(cities)}")
        else:
            by_reg.setdefault(reg, []).append(country)
    parts = []
    for reg in REGION_ORDER:
        if reg in by_reg:
            parts.append(f"{reg}({'; '.join(by_reg[reg])})")
    for reg, items in by_reg.items():
        if reg not in REGION_ORDER:
            parts.append(f"{reg}({'; '.join(items)})")
    return "; ".join(parts)


def fmt_op(countries: list[str]) -> str:
    return "; ".join(countries)


def main() -> None:
    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ci = headers.index("Company")
    gi = headers.index("Continent / Geography")
    oi = headers.index("Operational Presence")

    updated = []
    missing = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, ci + 1).value
        if not name or name not in GEO:
            continue
        geo = fmt_geo(GEO[name])
        op = fmt_op(OP[name])
        old_g = ws.cell(row, gi + 1).value
        old_o = ws.cell(row, oi + 1).value
        ws.cell(row, gi + 1).value = geo
        ws.cell(row, oi + 1).value = op
        updated.append({
            "company": name,
            "old_geo": old_g,
            "new_geo": geo,
            "old_op": old_o,
            "new_op": op,
            "op_countries": len(OP[name]),
        })

    for name in GEO:
        if not any(ws.cell(r, ci + 1).value == name for r in range(2, ws.max_row + 1)):
            missing.append(name)

    wb.save(FP)
    audit = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "updated": len(updated),
        "missing_in_sheet": missing,
        "rows": updated,
        "note": "Evidence-only multi-country expand; cities only when verified",
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"updated={len(updated)} missing={missing}")
    print(f"saved {FP}")
    print(f"audit {AUDIT}")
    for u in updated:
        print(f"  {u['company']}: OP {u['op_countries']} countries")


if __name__ == "__main__":
    main()
