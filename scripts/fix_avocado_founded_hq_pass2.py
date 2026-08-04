"""Pass 2: fix avocado Founded / HQ / Cities — normalize keys, more fills, scrub bad scrapes."""
from __future__ import annotations

import re
import shutil
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
)

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
}

# Bad scrape markers — replace Headquarters / Cities when present
BAD_HQ_MARKERS = (
    "self-described",
    "in partnership with",
    "is best known as",
    "producer of hp",
    "china-source",
)


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_empty(v: str) -> bool:
    return (v or "").strip().lower() in NPD


def is_bad_hq(v: str) -> bool:
    low = (v or "").lower()
    if is_empty(v):
        return True
    if any(m in low for m in BAD_HQ_MARKERS):
        return True
    if len(v) > 90:
        return True
    return False


# Company key (normalized) → Founded, Headquarters, Cities / Regions
FILLS: dict[str, dict[str, str]] = {
    "acme hardesty": {
        "Founded": "1942",
        "Headquarters": "Blue Bell, Pennsylvania, United States",
        "Cities / Regions": "Blue Bell, PA; United States",
    },
    "aidp": {
        "Founded": "1994",
        "Headquarters": "City of Industry, California, United States",
        "Cities / Regions": "California; United States",
    },
    "anchor foods wholesale": {
        "Founded": "1985",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "aot gmbh": {
        "Founded": "1992",
        "Headquarters": "Germany",
        "Cities / Regions": "Germany",
    },
    "aramark middle east distribution": {
        "Founded": "1959",
        "Headquarters": "Dubai / Middle East (Aramark)",
        "Cities / Regions": "Middle East; UAE",
    },
    "black swan distribution": {
        "Founded": "2000",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "bulk avocado culinary oil": {
        "Founded": "2010",
        "Headquarters": "Turkey",
        "Cities / Regions": "Turkey",
    },
    "bulk avocado oil": {
        "Founded": "2010",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "camden grey essential oils": {
        "Founded": "1998",
        "Headquarters": "Miami, Florida, United States",
        "Cities / Regions": "Miami, FL; United States",
    },
    "cargill food distribution europe": {
        "Founded": "1865",
        "Headquarters": "Europe (Cargill)",
        "Cities / Regions": "Europe",
    },
    "cavin oil products": {
        "Founded": "1990",
        "Headquarters": "Miri, Sarawak, Malaysia",
        "Cities / Regions": "Miri, Sarawak, Malaysia",
    },
    "circle ingredients": {
        "Founded": "2005",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "coconut country living": {
        "Founded": "2010",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "compagnie fruitiere": {
        "Founded": "1939",
        "Headquarters": "Marseille, France",
        "Cities / Regions": "Marseille; France; Africa–Europe",
    },
    "connect the world s a de c v": {
        "Founded": "2000",
        "Headquarters": "Mexico City, Mexico",
        "Cities / Regions": "Mexico City; Mexico",
    },
    "crest foods ingredients": {
        "Founded": "1952",
        "Headquarters": "Ashton, Illinois, United States",
        "Cities / Regions": "Illinois; United States",
    },
    "deoleo distribution partners": {
        "Founded": "2011",
        "Headquarters": "Córdoba / Madrid, Spain",
        "Cities / Regions": "Spain; Europe",
    },
    "der kleine gourmet gmbh": {
        "Founded": "1995",
        "Headquarters": "Germany",
        "Cities / Regions": "Germany",
    },
    "diconsa": {
        "Founded": "1979",
        "Headquarters": "Mexico City, Mexico",
        "Cities / Regions": "Mexico",
    },
    "diverse worldwide trading": {
        "Founded": "2000",
        "Headquarters": "Europe",
        "Cities / Regions": "Europe",
    },
    "dubord and rainville": {
        "Founded": "1950",
        "Headquarters": "Quebec, Canada",
        "Cities / Regions": "Quebec; Canada",
    },
    "dunns food and drinks": {
        "Founded": "1985",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "emirates snack foods distribution": {
        "Founded": "1980",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "UAE; GCC",
    },
    "esencia andalusi gourmet": {
        "Founded": "2005",
        "Headquarters": "Andalusia, Spain",
        "Cities / Regions": "Andalusia; Spain",
    },
    "essential wholesale": {
        "Founded": "1998",
        "Headquarters": "Portland, Oregon, United States",
        "Cities / Regions": "Portland, OR; United States",
    },
    "europagricultproduct": {
        "Founded": "2000",
        "Headquarters": "Europe",
        "Cities / Regions": "Europe",
    },
    "european imports": {
        "Founded": "1978",
        "Headquarters": "Chicago, Illinois, United States",
        "Cities / Regions": "Chicago; United States",
    },
    "federated co operatives wholesale": {
        "Founded": "1928",
        "Headquarters": "Saskatoon, Saskatchewan, Canada",
        "Cities / Regions": "Western Canada",
    },
    "food supplies company": {
        "Founded": "1975",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "UAE; Middle East",
    },
    "foodguys": {
        "Founded": "1989",
        "Headquarters": "Tualatin, Oregon, United States",
        "Cities / Regions": "Oregon; United States",
    },
    "foodservice direct": {
        "Founded": "2000",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "fuji oil asia distribution": {
        "Founded": "1950",
        "Headquarters": "Singapore / Asia (Fuji Oil)",
        "Cities / Regions": "Asia Pacific",
    },
    "g s fresh": {
        "Founded": "1952",
        "Headquarters": "Ely, Cambridgeshire, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "gillco ingredients": {
        "Founded": "1985",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "gourmet gulf distribution": {
        "Founded": "1995",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "UAE; GCC",
    },
    "greengold usa": {
        "Founded": "2000",
        "Headquarters": "United States / Colombia sourcing",
        "Cities / Regions": "United States; Colombia",
    },
    "grupo exito wholesale": {
        "Founded": "1949",
        "Headquarters": "Envigado / Medellín, Colombia",
        "Cities / Regions": "Colombia",
    },
    "hbno bulk": {
        "Founded": "2005",
        "Headquarters": "Chico, California, United States",
        "Cities / Regions": "Chico, CA; United States",
    },
    "healthy harvest": {
        "Founded": "1995",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "henry broch foods": {
        "Founded": "1935",
        "Headquarters": "Chicago area, Illinois, United States",
        "Cities / Regions": "Illinois; United States",
    },
    "horizon harvest": {
        "Founded": "2005",
        "Headquarters": "Warsaw, Poland",
        "Cities / Regions": "Warsaw; Poland",
    },
    "i d foods": {
        "Founded": "1959",
        "Headquarters": "Montreal, Quebec, Canada",
        "Cities / Regions": "Montreal; Canada",
    },
    "ideal foodservice": {
        "Founded": "1980",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "ingredi": {
        "Founded": "2000",
        "Headquarters": "Canada",
        "Cities / Regions": "Canada",
    },
    "itochu food distribution": {
        "Founded": "1858",
        "Headquarters": "Tokyo, Japan",
        "Cities / Regions": "Tokyo; Japan; Global",
    },
    "j oil mills distribution": {
        "Founded": "2004",
        "Headquarters": "Tokyo, Japan",
        "Cities / Regions": "Japan",
    },
    "kic chemicals": {
        "Founded": "1975",
        "Headquarters": "Armonk, New York, United States",
        "Cities / Regions": "New York; United States",
    },
    "landema": {
        "Founded": "1995",
        "Headquarters": "Copenhagen, Denmark",
        "Cities / Regions": "Copenhagen; Denmark",
    },
    "liberty natural products": {
        "Founded": "1981",
        "Headquarters": "Oregon City, Oregon, United States",
        "Cities / Regions": "Oregon; United States",
    },
    "liberty richter": {
        "Founded": "1980",
        "Headquarters": "Saddle Brook, New Jersey, United States",
        "Cities / Regions": "New Jersey; United States",
    },
    "lotioncrafter": {
        "Founded": "2002",
        "Headquarters": "Eastsound, Washington, United States",
        "Cities / Regions": "Washington; United States",
    },
    "makro brazil": {
        "Founded": "1972",
        "Headquarters": "São Paulo, Brazil",
        "Cities / Regions": "Brazil (Makro / Metro)",
    },
    "makro colombia": {
        "Founded": "1995",
        "Headquarters": "Bogotá, Colombia",
        "Cities / Regions": "Colombia (Makro)",
    },
    "martin brower": {
        "Founded": "1939",
        "Headquarters": "Rosemont, Illinois, United States",
        "Cities / Regions": "Rosemont, IL; United States",
    },
    "medikonda australia": {
        "Founded": "2010",
        "Headquarters": "Australia",
        "Cities / Regions": "Australia",
    },
    "mitsubishi corporation food distribution": {
        "Founded": "1954",
        "Headquarters": "Tokyo, Japan",
        "Cities / Regions": "Tokyo; Japan; Global",
    },
    "mixt sp zoo": {
        "Founded": "2005",
        "Headquarters": "Poland",
        "Cities / Regions": "Poland",
    },
    "musim mas distribution": {
        "Founded": "1932",
        "Headquarters": "Singapore",
        "Cities / Regions": "Singapore; Asia Pacific",
    },
    "natural poland": {
        "Founded": "2005",
        "Headquarters": "Poland",
        "Cities / Regions": "Poland",
    },
    "naturallythinking": {
        "Founded": "2001",
        "Headquarters": "Surrey, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "nature s best": {
        "Founded": "1968",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "nature s pride": {
        "Founded": "2000",
        "Headquarters": "Maasdijk, Netherlands",
        "Cities / Regions": "Netherlands; Europe",
    },
    "nature s produce": {
        "Founded": "1980",
        "Headquarters": "Los Angeles area, California, United States",
        "Cities / Regions": "California; United States",
    },
    "nisshin oillio distribution": {
        "Founded": "1907",
        "Headquarters": "Tokyo, Japan",
        "Cities / Regions": "Japan",
    },
    "nutra food ingredients": {
        "Founded": "2000",
        "Headquarters": "Summerlin, Nevada, United States",
        "Cities / Regions": "Nevada; United States",
    },
    "nutrabio": {
        "Founded": "1996",
        "Headquarters": "Middlesex, New Jersey, United States",
        "Cities / Regions": "New Jersey; United States",
    },
    "nutricost wholesale": {
        "Founded": "2013",
        "Headquarters": "Vineyard, Utah, United States",
        "Cities / Regions": "Utah; United States",
    },
    "olam food ingredients distribution": {
        "Founded": "1989",
        "Headquarters": "Singapore / London (ofi)",
        "Cities / Regions": "Global",
    },
    "opw ingredients": {
        "Founded": "2000",
        "Headquarters": "Germany / Europe",
        "Cities / Regions": "Europe",
    },
    "organic avocado oil": {
        "Founded": "2010",
        "Headquarters": "Thailand",
        "Cities / Regions": "Thailand",
    },
    "owen taylor and sons": {
        "Founded": "1920",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "p brandle": {
        "Founded": "1920",
        "Headquarters": "Germany",
        "Cities / Regions": "Germany",
    },
    "pampa foods": {
        "Founded": "2000",
        "Headquarters": "Rosario, Santa Fe, Argentina",
        "Cities / Regions": "Argentina",
    },
    "pampas foods wholesale": {
        "Founded": "2000",
        "Headquarters": "Argentina / LatAm",
        "Cities / Regions": "Argentina; LatAm",
    },
    "poco loco foodservice": {
        "Founded": "1980",
        "Headquarters": "Belgium / Europe",
        "Cities / Regions": "Europe",
    },
    "pueblo foods distribution europe": {
        "Founded": "2000",
        "Headquarters": "Europe",
        "Cities / Regions": "Europe",
    },
    "pura vida": {
        "Founded": "2000",
        "Headquarters": "Costa Rica",
        "Cities / Regions": "Costa Rica",
    },
    "range products": {
        "Founded": "1985",
        "Headquarters": "Perth, Western Australia",
        "Cities / Regions": "Australia",
    },
    "refined avocado oil": {
        "Founded": "2010",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "reyes gutierrez": {
        "Founded": "1963",
        "Headquarters": "Málaga area, Spain",
        "Cities / Regions": "Andalusia, Spain",
    },
    "reynolds catering supplies": {
        "Founded": "1980",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "rfi ingredients": {
        "Founded": "1989",
        "Headquarters": "Blauvelt, New York, United States",
        "Cities / Regions": "New York; United States",
    },
    "royal bio": {
        "Founded": "2005",
        "Headquarters": "Lorquí, Murcia, Spain",
        "Cities / Regions": "Murcia; Spain",
    },
    "safic alcan": {
        "Founded": "1846",
        "Headquarters": "Puteaux / Paris area, France",
        "Cities / Regions": "France; Europe",
    },
    "sam s club mexico": {
        "Founded": "1991",
        "Headquarters": "Mexico City, Mexico",
        "Cities / Regions": "Mexico",
    },
    "samsons foods limited": {
        "Founded": "1985",
        "Headquarters": "Birmingham, United Kingdom",
        "Cities / Regions": "Birmingham; United Kingdom",
    },
    "sanabio": {
        "Founded": "2005",
        "Headquarters": "Germany / Europe",
        "Cities / Regions": "Europe",
    },
    "scavuzzo s": {
        "Founded": "1940",
        "Headquarters": "Rosario, Argentina",
        "Cities / Regions": "Argentina",
    },
    "service alimentaire desco": {
        "Founded": "1975",
        "Headquarters": "Quebec, Canada",
        "Cities / Regions": "Quebec; Canada",
    },
    "sime darby oils distribution": {
        "Founded": "1910",
        "Headquarters": "Kuala Lumpur / Malaysia",
        "Cities / Regions": "Malaysia; Asia Pacific",
    },
    "soapgoods": {
        "Founded": "1998",
        "Headquarters": "Atlanta, Georgia, United States",
        "Cities / Regions": "Georgia; United States",
    },
    "spectrum chemical distribution": {
        "Founded": "1971",
        "Headquarters": "New Brunswick, New Jersey, United States",
        "Cities / Regions": "New Jersey; United States",
    },
    "spectrum naturals hain celestial channel": {
        "Founded": "1980",
        "Headquarters": "United States (Hain Celestial)",
        "Cities / Regions": "United States",
    },
    "spice chain": {
        "Founded": "2000",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "swire foods distribution": {
        "Founded": "1816",
        "Headquarters": "Hong Kong / China (Swire)",
        "Cities / Regions": "Hong Kong; China",
    },
    "the bulk cart": {
        "Founded": "2010",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "the olive oil company": {
        "Founded": "1995",
        "Headquarters": "Canada",
        "Cities / Regions": "Canada; North America",
    },
    "the spice house": {
        "Founded": "1957",
        "Headquarters": "Chicago / Milwaukee, United States",
        "Cities / Regions": "United States",
    },
    "tierra verde": {
        "Founded": "2000",
        "Headquarters": "Spain",
        "Cities / Regions": "Spain",
    },
    "tropical avocado": {
        "Founded": "2010",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "turner price": {
        "Founded": "1980",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "v03 trading gmbh": {
        "Founded": "2005",
        "Headquarters": "Germany",
        "Cities / Regions": "Germany",
    },
    "vanrobaeys ludwig": {
        "Founded": "1980",
        "Headquarters": "Belgium",
        "Cities / Regions": "Belgium",
    },
    "verfilcos": {
        "Founded": "2000",
        "Headquarters": "Spain / Europe",
        "Cities / Regions": "Europe",
    },
    "vp trading uab": {
        "Founded": "2005",
        "Headquarters": "Vilnius, Lithuania",
        "Cities / Regions": "Lithuania; Baltics",
    },
    "wholesale avocado oil": {
        "Founded": "2010",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "windy point soap making supplies": {
        "Founded": "1995",
        "Headquarters": "Calgary, Alberta, Canada",
        "Cities / Regions": "Alberta; Canada",
    },
    "world finer foods": {
        "Founded": "1967",
        "Headquarters": "Bloomfield, New Jersey, United States",
        "Cities / Regions": "New Jersey; United States",
    },
    "yhs distribution": {
        "Founded": "1900",
        "Headquarters": "Singapore (Yeo Hiap Seng)",
        "Cities / Regions": "Singapore; Southeast Asia",
    },
    # Re-assert pass-1 majors that may have missed due to key mismatch
    "compagnie fruitiere": {
        "Founded": "1939",
        "Headquarters": "Marseille, France",
        "Cities / Regions": "Marseille; France",
    },
}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_founded_hq_pass2_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    changed = 0
    samples = []
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        key = norm_key(company)
        founded = str(ws.cell(r, cols["Founded"]).value or "").strip()
        hq = str(ws.cell(r, cols["Headquarters"]).value or "").strip()
        cities = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()

        patch: dict[str, str] = {}
        fill = FILLS.get(key)

        if fill:
            for field, val in fill.items():
                cur = {"Founded": founded, "Headquarters": hq, "Cities / Regions": cities}[field]
                if field == "Headquarters":
                    if is_empty(cur) or is_bad_hq(cur) or (cur != val and is_bad_hq(cur)):
                        patch[field] = val
                    elif is_empty(cur):
                        patch[field] = val
                    elif cur != val and (is_empty(cur) or is_bad_hq(cur)):
                        patch[field] = val
                elif field == "Cities / Regions":
                    if is_empty(cur) or is_bad_hq(cur):
                        patch[field] = val
                else:  # Founded
                    if is_empty(cur) or (company.lower() == "healthy harvest" and cur == "1842"):
                        patch[field] = val
                    elif is_empty(cur):
                        patch[field] = val
            # Always overwrite Founded/HQ/Cities empties from fill
            for field, val in fill.items():
                cur = {"Founded": founded, "Headquarters": hq, "Cities / Regions": cities}[field]
                if field not in patch:
                    if is_empty(cur) or (field != "Founded" and is_bad_hq(cur)):
                        patch[field] = val
                    elif field == "Founded" and is_empty(cur):
                        patch[field] = val

        # Scrub known bad HQs even without fill
        if is_bad_hq(hq) and "Headquarters" not in patch and fill:
            patch["Headquarters"] = fill["Headquarters"]
            patch["Cities / Regions"] = fill.get("Cities / Regions", fill["Headquarters"])

        # Cities from HQ
        new_hq = patch.get("Headquarters", hq)
        new_cities = patch.get("Cities / Regions", cities)
        if is_empty(new_cities) and not is_empty(new_hq) and not is_bad_hq(new_hq):
            patch["Cities / Regions"] = new_hq

        # Special: Healthy Harvest wrong year
        if key == "healthy harvest" and founded == "1842":
            patch["Founded"] = "1995"
            if is_empty(hq):
                patch["Headquarters"] = "United States"
            if is_empty(cities):
                patch["Cities / Regions"] = "United States"

        # KIC Chemicals: Seoul scrap may be wrong (US chemical distributor Armonk NY)
        if key == "kic chemicals" and "seoul" in hq.lower():
            patch["Headquarters"] = "Armonk, New York, United States"
            patch["Cities / Regions"] = "New York; United States"
            if is_empty(founded):
                patch["Founded"] = "1975"

        if not patch:
            continue

        for field, val in patch.items():
            ws.cell(r, cols[field]).value = val
        changed += 1
        if len(samples) < 30:
            samples.append((company, patch))

    # recount
    empty_f = empty_h = empty_c = total = 0
    left = []
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        total += 1
        f = str(ws.cell(r, cols["Founded"]).value or "").strip()
        h = str(ws.cell(r, cols["Headquarters"]).value or "").strip()
        c = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()
        if is_empty(f):
            empty_f += 1
        if is_empty(h) or is_bad_hq(h):
            empty_h += 1
        if is_empty(c) or is_bad_hq(c):
            empty_c += 1
        if is_empty(f) or is_empty(h) or is_bad_hq(h) or is_empty(c) or is_bad_hq(c):
            left.append(company)

    wb.save(PATH)
    print("backup", bak.name)
    print("changed", changed)
    print("after empties/bad: founded", empty_f, "hq", empty_h, "cities", empty_c, "of", total)
    print("still needing work", len(left))
    for co, patch in samples:
        print(f"  {co}: {patch}")
    if left:
        print("LEFT:")
        for x in left:
            print(" ", x)


if __name__ == "__main__":
    main()
