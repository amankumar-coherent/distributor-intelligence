# -*- coding: utf-8 -*-
"""
VERIFY + FILL Continent / Geography and Operational Presence.

DEFINITIONS (user):
1) Continent / Geography = countries with PHYSICAL ops
   (HQ, offices, warehouses, plants, stores). Write ALL country names
   where physically present. Cities included when evidenced.
   Format: NA(United States, Canada, Mexico)  OR with cities:
           NA(United States: City1 [HQ], City2; Canada: City3)
2) Operational Presence = countries of COMMERCIAL sell/serve reach.
   Write ALL country names. Format: United States; Canada; Mexico

NEVER invent countries or cities.
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
AUDIT = FP.parent / "geo_op_countries_verify_audit.json"

COUNTRY_REGION = {
    "United States": "NA", "Canada": "NA", "Mexico": "NA", "Puerto Rico": "NA",
    "Panama": "NA", "Guam": "NA", "U.S. Virgin Islands": "NA",
    "Brazil": "LATAM", "Argentina": "LATAM", "Chile": "LATAM", "Colombia": "LATAM",
    "Peru": "LATAM", "Bolivia": "LATAM", "Uruguay": "LATAM",
    "United Kingdom": "EU", "Ireland": "EU", "France": "EU", "Germany": "EU",
    "Spain": "EU", "Portugal": "EU", "Italy": "EU", "Netherlands": "EU",
    "Belgium": "EU", "Luxembourg": "EU", "Switzerland": "EU", "Austria": "EU",
    "Sweden": "EU", "Norway": "EU", "Denmark": "EU", "Finland": "EU", "Poland": "EU",
    "Czechia": "EU", "Romania": "EU", "Hungary": "EU", "Greece": "EU",
    "Lithuania": "EU", "Latvia": "EU", "Estonia": "EU", "Slovakia": "EU",
    "Bulgaria": "EU", "Turkey": "EU",
    "South Africa": "AF", "Tunisia": "AF", "Angola": "AF",
    "Israel": "ME",
    "Japan": "APAC", "China": "APAC", "Hong Kong": "APAC", "Taiwan": "APAC",
    "South Korea": "APAC", "India": "APAC", "Singapore": "APAC", "Malaysia": "APAC",
    "Thailand": "APAC",
    "Australia": "OC", "New Zealand": "OC",
}
REGION_ORDER = ["NA", "LATAM", "EU", "AF", "ME", "APAC", "OC"]

# Physical: country -> cities (optional). ALL countries listed.
PHYSICAL: dict[str, dict[str, list[str]]] = {
    "Uline": {
        "United States": [
            "Pleasant Prairie WI [HQ]", "Braselton GA", "Plainfield CT", "Kenosha WI",
            "Etna OH", "DFW Airport TX", "Ontario CA", "Naples FL", "Hudson WI",
            "Allentown PA", "Reno NV", "Lacey WA",
        ],
        "Canada": ["Milton ON", "St. Albert AB"],
        "Mexico": ["Mexicali BC", "Apodaca NL"],
    },
    "Associated Bag": {
        "United States": ["Milwaukee WI [HQ]", "New Kingstown PA", "Dallas TX", "Sparks NV"],
    },
    "Hoffmann Verpackung": {
        "Germany": [
            "Kaarst [HQ]", "Wetter/Ruhr", "Schkeuditz", "Mittenwalde",
            "Bad Camberg", "Erding", "Ahrensburg",
        ],
        "Switzerland": ["Nesselnbach"],
    },
    "RAJA": {
        "France": ["Tremblay-en-France / Roissy [HQ]", "Clichy", "Survilliers", "Neuville-en-Ferrain", "Villeneuve-Loubet"],
        "United Kingdom": ["Brogborough", "Harlow", "Leicester"],
        "Germany": ["Ettlingen", "Duisburg", "Grossostheim"],
        "Austria": ["Vienna", "Grodig", "Pregarten"],
        "Belgium": ["Tongeren", "Mouscron", "Antwerp"],
        "Netherlands": ["Breda", "Venlo"],
        "Spain": ["Ripollet", "Galdakao"],
        "Italy": ["Castel San Giovanni", "Castelletto Cervo"],
        "Switzerland": ["Pratteln", "Baar", "Aarau"],
        "Poland": ["Warsaw"],
        "Czechia": ["Nupaky / Ricany"],
        "Denmark": ["Soborg"],
        "Sweden": ["Landvetter"],
        "Norway": ["Vestby"],
        "Portugal": ["Cacem"],
        "Slovakia": ["Bratislava"],
        "Hungary": ["Budapest"],
        "Ireland": ["Dublin"],
        "Luxembourg": ["Differdange"],
    },
    "RAJAPACK UK": {"United Kingdom": ["Brogborough [HQ]", "Harlow"]},
    "NNZ": {
        "Netherlands": ["Groningen [HQ]"],
        "Germany": ["Lueneburg", "Wesel", "Mutterstadt"],
        "Austria": [], "Switzerland": [], "Denmark": [], "Norway": [], "Finland": [],
        "Poland": [], "Estonia": [], "Latvia": [], "Lithuania": [], "United Kingdom": [],
        "United States": [], "Canada": [], "South Africa": [], "China": [],
    },
    "NNZ GmbH": {"Germany": ["Lueneburg", "Wesel", "Mutterstadt"]},
    "Ahlsell": {
        "Sweden": ["Stockholm [HQ]", "Hallsberg"],
        "Norway": ["Oslo"], "Finland": ["Hyvinkaa"], "Denmark": ["Brondby"],
        "Estonia": ["Tallinn"], "Poland": [], "China": [],
    },
    "Antalis": {
        "France": ["Boulogne-Billancourt [HQ]"],
        "Austria": [], "Belgium": [], "Bolivia": [], "Bulgaria": [], "Canada": [],
        "Chile": [], "Czechia": [], "Denmark": [], "Estonia": [], "Finland": [],
        "Germany": [], "Hungary": [], "Ireland": [], "Italy": [], "Latvia": [],
        "Lithuania": [], "Mexico": [], "Netherlands": [], "Norway": [], "Peru": [],
        "Poland": [], "Portugal": [], "Romania": [], "Slovakia": [], "Spain": [],
        "Sweden": [], "Switzerland": [], "Turkey": [], "United Kingdom": [],
    },
    "Lyreco": {
        "France": ["Marly [HQ]"],
        "Austria": [], "Belgium": [], "Czechia": [], "Denmark": [], "Finland": [],
        "Germany": [], "Hungary": [], "Ireland": [], "Italy": [], "Luxembourg": [],
        "Netherlands": [], "Norway": [], "Poland": [], "Portugal": [], "Slovakia": [],
        "Spain": [], "Sweden": [], "Switzerland": [], "United Kingdom": [],
        "Hong Kong": [], "South Korea": [], "Malaysia": [], "Singapore": [], "Thailand": [],
    },
    "Bunzl": {
        "United Kingdom": ["London [HQ]"],
        "Ireland": [],
        "United States": ["Creve Coeur / St. Louis MO"],
        "Canada": ["Burlington ON"],
        "Mexico": [], "Puerto Rico": [],
        "Australia": ["Mount Waverley VIC"],
        "New Zealand": [],
        "Austria": [], "Belgium": [], "Brazil": [], "Chile": [], "China": [],
        "Colombia": [], "Czechia": [], "Denmark": [], "France": [], "Germany": [],
        "Hong Kong": [], "Hungary": [], "Israel": [], "Italy": [], "Netherlands": [],
        "Peru": [], "Poland": [], "Romania": [], "Singapore": [], "Slovakia": [],
        "Spain": [], "Switzerland": [], "Turkey": [],
    },
    "Veritiv": {"United States": ["Atlanta GA [HQ]"], "Mexico": []},
    "Imperial Dade": {"United States": ["Jersey City NJ [HQ]"], "Canada": []},
    "Global Industrial": {"United States": ["Port Washington NY [HQ]"]},
    "Global Industrial Canada": {"Canada": ["Markham ON [HQ]"]},
    "Fastenal": {
        "United States": ["Winona MN [HQ]"],
        "Canada": [], "Mexico": [], "Puerto Rico": [], "Panama": [], "China": [],
    },
    "Manutan": {
        "France": ["Gonesse [HQ]", "Niort", "Veauche", "Strasbourg"],
        "United Kingdom": [], "Belgium": [], "Netherlands": [], "Germany": [],
        "Italy": [], "Spain": [], "Portugal": [], "Denmark": [], "Sweden": [],
        "Norway": [], "Finland": [], "Czechia": [], "Slovakia": [], "Hungary": [],
        "Poland": [], "Switzerland": [],
    },
    "Brangs + Heinrich": {
        "Germany": ["Solingen [HQ]"], "Netherlands": ["Utrecht"], "Belgium": ["Dison"],
    },
    "Verpacken Online": {"Germany": ["Solingen"]},
    "Inapa": {
        "Portugal": ["Sintra [HQ]"], "Germany": ["Hamburg"], "France": ["Dourdan"],
        "Spain": [], "Belgium": [], "Luxembourg": [], "Turkey": [], "Angola": [],
    },
    "Papier Union": {"Germany": ["Hamburg"]},
    "Staples": {"United States": ["Framingham MA [HQ]"]},  # US retail entity; Canada is separate co
    "Quill": {"United States": ["Lincolnshire IL [HQ]"]},
    "The Home Depot": {
        "United States": ["Atlanta GA [HQ]"],
        "Canada": [],
        "Mexico": [],
        "Puerto Rico": [],
    },
    "Officeworks": {"Australia": ["Chadstone VIC [HQ]"]},
    "Blackwoods": {"Australia": ["Macquarie Park NSW [HQ]"]},
    "Winc": {"Australia": ["Mascot NSW [HQ]"]},
    "Ball & Doggett": {"Australia": ["Dandenong South VIC [HQ]"]},
    "Spicers": {"Australia": ["Dandenong South VIC [HQ]"]},
    "Spicers Paper": {"Canada": ["Vaughan ON [HQ]"]},
    "Cospak": {"Australia": ["Minto NSW [HQ]"]},
    "Signet": {"Australia": ["Wakerley QLD [HQ]"]},
    "Allied Packaging": {"Australia": ["Botany NSW [HQ]"]},
    "Pluspak": {
        "Australia": ["Scoresby VIC [HQ]", "Milperra NSW", "Crestmead QLD", "Canning Vale WA"],
    },
    "Ozpack": {"Australia": ["Silverwater NSW [HQ]"]},
    "National Packaging Services": {
        "United States": ["Secaucus NJ [HQ]", "Rochester NY"],
    },
    "Service Box & Tape": {
        "United States": [
            "Atlanta GA - Donnelly Ave [HQ]",
            "Atlanta GA - Liddell Drive",
            "City of South Fulton GA - Welcome All Road",
        ],
    },
    "Skybox Packaging": {"United States": ["Mansfield OH [HQ]"]},
    "Ace Paper & Packaging": {"United States": ["Los Angeles CA [HQ]"]},
    "All American Packaging": {"United States": ["Forest Hills NY [HQ]"]},
    "Askul": {"Japan": ["Tokyo [HQ]"]},
    "MonotaRO": {"Japan": ["Osaka [HQ]"]},
    "Macfarlane Packaging": {"United Kingdom": ["Glasgow [HQ]"]},
    "National Flexible": {"United Kingdom": ["Birkenshaw / Bradford [HQ]"]},
    "Abco Kovex": {"Ireland": ["Swords / Dublin [HQ]"]},
    "The Packaging Centre": {"Ireland": ["Dublin [HQ]"], "United Kingdom": ["Saintfield"]},
    "McLernon Packaging": {"Ireland": ["Dublin [HQ]"]},
    "Viking Direct": {"Ireland": ["Dublin [HQ]"]},
    "Mondoffice": {"France": ["Survilliers [HQ]"]},
    "Kalamazoo": {"Spain": ["Galdakao [HQ]"]},
    "JPG": {"France": ["Survilliers [HQ]"]},
    "Syspex": {"Singapore": ["Joo Koon [HQ]"], "Malaysia": []},
    "Seow Khim Polythene": {"Singapore": ["Pasir Ris / Loyang [HQ]"]},
    "Tantraco": {"Singapore": ["Ang Mo Kio [HQ]"]},
    "FlexPRO": {"Bulgaria": ["Sofia [HQ]"]},
    "Flexi-Films": {"United Kingdom": ["Annesley / Nottingham [HQ]"]},
    "Empaques del Norte": {"Mexico": ["San Nicolas de los Garza [HQ]"]},
    "Transpak": {"Germany": ["Solms [HQ]"]},
    "ESWE Versandpack": {"Germany": ["Sternenfels [HQ]"]},
    "Packaging2Buy": {"United Kingdom": ["Stevenage"]},
    "Packaging Express": {"United Kingdom": ["Hayes"]},
    "Swiftpak": {"United Kingdom": ["Theale / Reading [HQ]"]},
    "Polythene UK": {"United Kingdom": ["Witney [HQ]"]},
    "UK Packaging Supplies": {"United Kingdom": ["Tottenham / London [HQ]"]},
    "Unicorr Packaging": {"United States": ["North Haven CT [HQ]"]},
    "Uboxes": {"United States": ["Miramar FL [HQ]"]},
    "The Bottle Crew": {"United States": ["Farmington Hills MI [HQ]"]},
    "The Boxery": {"United States": ["Brooklyn NY [HQ]"]},
    "Stamar Packaging": {"United States": ["Elgin IL [HQ]"]},
    "Schermerhorn Bros.": {"United States": ["Lombard IL [HQ]"]},
    "Preferred Plastics & Packaging": {"United States": ["Lyndhurst NJ [HQ]"]},
    "Plastic Bag Source": {"United States": ["Monsey NY [HQ]"]},
    "Packaging Partners": {"United States": ["Franklin WI [HQ]"]},
    "Packaging Price": {"United States": ["Warrenville IL [HQ]"]},
    "Great Pacific Packaging": {"United States": ["Fife WA [HQ]"]},
    "Innovative Packaging Group": {"United States": ["Dallas TX [HQ]"]},
    "Imperial Supplies": {"United States": ["Green Bay WI [HQ]"]},
    "R.J. Schinner Co.": {"United States": ["Menomonee Falls WI [HQ]"]},
    "Gertex Solutions": {"Canada": ["Etobicoke ON [HQ]"]},
    "Poly Pak": {"Canada": ["Vancouver BC [HQ]"]},
    "Super Poly": {"Canada": ["Edmonton AB [HQ]"]},
    "Unipac Packaging Products": {"Canada": ["Edmonton AB [HQ]"]},
    "PakFactory": {"Canada": ["Markham ON [HQ]"]},
    "Atlantic Packaging Products": {"Canada": ["Scarborough ON [HQ]"]},
    "Richards Packaging": {"Canada": ["Mississauga ON [HQ]"]},
    "Crawford Packaging": {"Canada": ["London ON [HQ]"]},
    "Polyair": {"Canada": ["Etobicoke ON"], "United States": ["Sarasota FL"]},
    "Furodet Emballages": {"France": ["Sennecey-le-Grand [HQ]", "Ormes"]},
    "Gerdmans": {"Sweden": ["Markaryd [HQ]"]},
    "Ratioform": {"Germany": ["Pliening [HQ]"]},
    "Papyrus": {"Sweden": ["Molndal [HQ]"]},
    "Igepa Group": {"Germany": ["Hamburg [HQ]"]},
    "Karton.eu": {"Germany": ["Spremberg [HQ]"]},
    "Network Packaging": {"United Kingdom": ["Wolverhampton [HQ]"]},
    "GWP Group": {"United Kingdom": ["Cricklade [HQ]"]},
    "GM Packaging": {"United Kingdom": ["Murton / Seaham [HQ]"]},
    "Industrial Packaging": {"United Kingdom": ["Southport [HQ]"]},
    "Lesters Packaging": {"United Kingdom": ["Burntwood [HQ]"]},
    "Multiplastics": {"United Kingdom": ["Northwich [HQ]"]},
    "Complete Packaging Solutions": {"United Kingdom": ["Long Crendon [HQ]"]},
    "Diamond Packaging": {"United Kingdom": ["Canvey Island [HQ]"]},
    "Able Packaging": {"United Kingdom": ["Leeds [HQ]"]},
    "Pack Supplies": {"United Kingdom": ["Thatcham"]},
    "Packaging World": {"United Kingdom": ["Swansea [HQ]"]},
    "Seal Packaging": {"United Kingdom": ["Luton"]},
    "Transpack": {"United Kingdom": ["Totton / Southampton"]},
    "Packaging House": {"New Zealand": ["Wiri / Auckland [HQ]"]},
    "Packaging Products": {"New Zealand": ["Lower Hutt [HQ]"]},
    "Premier Packaging": {"New Zealand": ["Mount Wellington / Auckland [HQ]"]},
    "Moglix": {"India": ["Noida [HQ]"]},
    "IndustryBuying": {"India": ["New Delhi [HQ]"]},
    "India Pack": {"India": ["Pune [HQ]"]},
    "Allied Distributors": {"South Africa": []},
    "Office Depot": {"Mexico": []},
    "SupplyOne": {"United States": ["Newtown Square PA [HQ]"]},
    "Packaging Corporation of America": {"United States": ["Lake Forest IL [HQ]"]},
    "Shorr Packaging": {"United States": ["Aurora IL [HQ]"]},
    "WebstaurantStore": {"United States": ["Lititz PA [HQ]"]},
    "Aramsco": {"United States": ["Radnor PA [HQ]"]},
    "W.B. Mason": {"United States": ["Brockton MA [HQ]"]},
    "Hisco": {"United States": ["Houston TX [HQ]"]},
    "Kaufman Container": {"United States": ["Cleveland OH [HQ]"]},
    "O.Berk": {"United States": ["Union NJ [HQ]"]},
    "SKS Bottle & Packaging": {"United States": ["Saratoga Springs NY [HQ]"]},
    "HD Supply": {"United States": ["Atlanta GA [HQ]"]},
    "BradyPLUS": {"United States": ["Las Vegas NV [HQ]"]},
    "Zoro Tools": {"United States": ["Buffalo Grove IL [HQ]"]},
    "Acme Paper & Supply": {"United States": ["Savage MD [HQ]"]},
    "All American Containers": {"United States": ["Miami FL [HQ]"]},
    "McMaster-Carr": {"United States": ["Elmhurst IL [HQ]"]},
}

OP: dict[str, list[str]] = {
    "Uline": ["United States", "Canada", "Mexico"],
    "Associated Bag": ["United States"],
    "Hoffmann Verpackung": ["Germany", "Switzerland"],
    "RAJA": [
        "Austria", "Belgium", "Czechia", "Denmark", "France", "Germany", "Hungary",
        "Ireland", "Italy", "Luxembourg", "Netherlands", "Norway", "Poland", "Portugal",
        "Slovakia", "Spain", "Sweden", "Switzerland", "United Kingdom",
    ],
    "RAJAPACK UK": ["United Kingdom"],
    "NNZ": [
        "Netherlands", "Germany", "Austria", "Switzerland", "Denmark", "Norway",
        "Finland", "Poland", "Estonia", "Latvia", "Lithuania", "United Kingdom",
        "United States", "Canada", "South Africa", "China",
    ],
    "NNZ GmbH": ["Germany"],
    "Ahlsell": ["Sweden", "Norway", "Finland", "Denmark", "Estonia", "Poland"],
    "Antalis": [
        "Austria", "Belgium", "Bolivia", "Bulgaria", "Canada", "Chile", "Czechia",
        "Denmark", "Estonia", "Finland", "France", "Germany", "Hungary", "Ireland",
        "Italy", "Latvia", "Lithuania", "Mexico", "Netherlands", "Norway", "Peru",
        "Poland", "Portugal", "Romania", "Slovakia", "Spain", "Sweden", "Switzerland",
        "Turkey", "United Kingdom",
    ],
    "Lyreco": [
        "Austria", "Belgium", "Czechia", "Denmark", "France", "Finland", "Germany",
        "Hungary", "Ireland", "Italy", "Luxembourg", "Netherlands", "Norway", "Poland",
        "Portugal", "Slovakia", "Spain", "Sweden", "Switzerland", "United Kingdom",
        "Hong Kong", "South Korea", "Malaysia", "Singapore", "Thailand",
        "United States", "Canada", "Mexico", "Brazil", "Argentina",
        "Latvia", "Lithuania", "Estonia", "Romania", "Turkey",
        "Taiwan", "China", "India", "Australia", "New Zealand",
    ],
    "Bunzl": [
        "United Kingdom", "Ireland", "United States", "Canada", "Mexico", "Puerto Rico",
        "Australia", "New Zealand", "Austria", "Belgium", "Brazil", "Chile", "China",
        "Colombia", "Czechia", "Denmark", "France", "Germany", "Hong Kong", "Hungary",
        "Israel", "Italy", "Netherlands", "Peru", "Poland", "Romania", "Singapore",
        "Slovakia", "Spain", "Switzerland", "Turkey",
    ],
    "Veritiv": ["United States", "Mexico"],
    "Imperial Dade": ["United States", "Canada"],
    "Global Industrial": ["United States"],
    "Global Industrial Canada": ["Canada"],
    "Fastenal": ["United States", "Canada", "Mexico", "Puerto Rico", "Panama"],
    "Manutan": [
        "France", "United Kingdom", "Belgium", "Netherlands", "Germany", "Italy",
        "Spain", "Portugal", "Denmark", "Sweden", "Norway", "Finland", "Czechia",
        "Slovakia", "Hungary", "Poland", "Switzerland",
    ],
    "Brangs + Heinrich": ["Germany", "Netherlands", "Belgium"],
    "Inapa": ["Portugal", "Germany", "France", "Spain", "Belgium", "Luxembourg", "Turkey", "Angola"],
    "The Packaging Centre": ["Ireland", "United Kingdom"],
    "The Home Depot": ["United States", "Canada", "Mexico", "Puerto Rico"],
    "Staples": ["United States"],
    "Officeworks": ["Australia"],
    "Pluspak": ["Australia"],
    "Syspex": ["Singapore", "Malaysia"],
    "Polyair": ["Canada", "United States"],
    "Allied Packaging": ["Australia"],
    "Blackwoods": ["Australia"],
    "Winc": ["Australia"],
    "Cospak": ["Australia"],
    "Ozpack": ["Australia"],
    "National Packaging Services": ["United States"],
    "Service Box & Tape": ["United States"],
    "Skybox Packaging": ["United States"],
    "Ace Paper & Packaging": ["United States"],
    "All American Packaging": ["United States", "Canada"],  # sales coverage NE + Canada
}


def canon(token: str) -> str | None:
    t = (token or "").strip()
    if not t:
        return None
    low = t.lower().rstrip(".")
    alias = {
        "usa": "United States", "us": "United States", "u.s.": "United States",
        "uk": "United Kingdom", "u.k.": "United Kingdom",
        "czech republic": "Czechia", "czechia": "Czechia",
        "holland": "Netherlands", "the netherlands": "Netherlands",
        "korea": "South Korea", "hong-kong": "Hong Kong",
    }
    if low in alias:
        return alias[low]
    specials = {
        "united states": "United States", "united kingdom": "United Kingdom",
        "south africa": "South Africa", "south korea": "South Korea",
        "new zealand": "New Zealand", "hong kong": "Hong Kong",
        "puerto rico": "Puerto Rico", "u.s. virgin islands": "U.S. Virgin Islands",
    }
    if low in specials:
        return specials[low]
    if t in COUNTRY_REGION:
        return t
    titled = t.title() if t.islower() or t.isupper() else t
    return titled if titled in COUNTRY_REGION else None


US_STATE_RE = re.compile(
    r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|"
    r"MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)\b"
)
CA_PROV_RE = re.compile(
    r"\b(Ontario|Quebec|Alberta|British Columbia|Manitoba|Saskatchewan|ON|QC|AB|BC|MB|SK)\b",
    re.I,
)


def parse_hq(hq: str) -> tuple[str | None, str | None]:
    s = (hq or "").strip()
    if not s or "not publicly" in s.lower():
        return None, None
    s_clean = re.sub(r"\([^)]*\)", " ", s)
    low = s_clean.lower()

    # Explicit country phrases first (avoid 'wales' matching New South Wales)
    if re.search(r"\baustralia\b", low):
        country = "Australia"
    elif re.search(r"\bnew zealand\b", low):
        country = "New Zealand"
    elif re.search(r"\bcanada\b", low):
        country = "Canada"
    elif re.search(r"\b(united states|usa|u\.s\.a\.|u\.s\.)\b", low):
        country = "United States"
    elif re.search(r"\b(united kingdom|\buk\b)\b", low):
        country = "United Kingdom"
    elif re.search(r"\b(england|scotland)\b", low) and "new south wales" not in low:
        country = "United Kingdom"
    else:
        country = None

    city = None
    if "," in s_clean:
        parts = [p.strip() for p in s_clean.split(",") if p.strip()]
        if not country:
            for part in reversed(parts):
                c = canon(re.split(r"[;/]", part)[0].strip())
                if c:
                    country = c
                    break
        city_parts = []
        for part in parts:
            part0 = re.split(r"[;/]", part)[0].strip()
            if country and canon(part0) == country:
                break
            if canon(part0):
                continue
            city_parts.append(part)
        city = ", ".join(city_parts).strip(" ,") if city_parts else None

    if not country:
        country = canon(s_clean)
    if not country and US_STATE_RE.search(s_clean):
        country = "United States"
        # Prefer "City, ST" pattern
        m = re.search(r"([A-Za-z .'-]+),\s*([A-Z]{2})\b(?:\s+\d{5})?", s_clean)
        if m:
            city = f"{m.group(1).strip()}, {m.group(2)}"
        elif not city:
            city = s_clean.strip()
    if not country and CA_PROV_RE.search(s_clean):
        country = "Canada"
    if country and city:
        return country, f"{city} [HQ]" if "[HQ]" not in city else city
    return country, f"{city} [HQ]" if city else None


def format_geo(physical: dict[str, list[str]]) -> str:
    """All country names; cities appended when known."""
    if not physical:
        return "Not publicly disclosed"
    buckets: dict[str, list[str]] = {r: [] for r in REGION_ORDER}
    for country, cities in physical.items():
        cc = canon(country)
        if not cc or cc not in COUNTRY_REGION:
            continue
        cities = [c for c in (cities or []) if c]
        if cities:
            buckets[COUNTRY_REGION[cc]].append(f"{cc}: {', '.join(cities)}")
        else:
            buckets[COUNTRY_REGION[cc]].append(cc)
    parts = [f"{r}({'; '.join(buckets[r])})" for r in REGION_ORDER if buckets[r]]
    return "; ".join(parts) if parts else "Not publicly disclosed"


def format_op(countries: list[str]) -> str:
    seen, out = set(), []
    for c in countries:
        cc = canon(c)
        if cc and cc in COUNTRY_REGION and cc not in seen:
            seen.add(cc)
            out.append(cc)
    return "; ".join(out) if out else "Not publicly disclosed"


def main() -> None:
    wb = load_workbook(FP)
    ws = wb.active
    assert ws.cell(1, 5).value == "Continent / Geography"
    assert ws.cell(1, 6).value == "Operational Presence"

    audit = {"updated_at": datetime.now().isoformat(timespec="seconds"), "definition": {
        "Continent / Geography": "ALL countries with physical ops (HQ/offices/warehouses/plants/stores); cities when evidenced",
        "Operational Presence": "ALL countries of commercial sell/serve reach",
    }, "rows": []}

    n_verified = n_hq = n_weak = 0
    for row in ws.iter_rows(min_row=2):
        name = str(row[0].value or "").strip()
        hq = str(row[3].value or "").strip()
        if not name:
            continue
        old_g, old_o = str(row[4].value or ""), str(row[5].value or "")
        hq_country, hq_city = parse_hq(hq)

        if name in PHYSICAL:
            phys = {k: list(v) for k, v in PHYSICAL[name].items()}
            if hq_country and hq_country in COUNTRY_REGION:
                if hq_country not in phys:
                    phys[hq_country] = [hq_city] if hq_city else []
                elif hq_city and not any("[HQ]" in x for x in phys[hq_country]):
                    phys[hq_country] = [hq_city] + phys[hq_country]
            source = "verified"
            n_verified += 1
        elif hq_country:
            phys = {hq_country: [hq_city] if hq_city else []}
            source = "hq"
            n_hq += 1
        else:
            phys = {}
            source = "none"
            n_weak += 1

        op_list = list(OP[name]) if name in OP else list(phys.keys())
        new_g, new_o = format_geo(phys), format_op(op_list)
        row[4].value = new_g
        row[5].value = new_o
        audit["rows"].append({
            "company": name, "source": source, "hq": hq,
            "old_geo": old_g, "new_geo": new_g, "old_op": old_o, "new_op": new_o,
        })

    wb.save(FP)
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")

    # validate no AU->UK bug
    bad = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        hq, g = str(row[3] or ""), str(row[4] or "")
        if "australia" in hq.lower() and "United Kingdom" in g and "Australia" not in g:
            bad += 1
    print(f"saved {FP}")
    print(f"verified={n_verified} hq_baseline={n_hq} weak={n_weak} au_uk_bug={bad}")
    print(f"audit {AUDIT}")


if __name__ == "__main__":
    main()
