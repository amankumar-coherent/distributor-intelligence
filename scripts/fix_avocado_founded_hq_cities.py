"""Fill / verify Founded, Headquarters, Cities / Regions on avocado FINAL."""
from __future__ import annotations

import json
import re
import shutil
import time
import urllib.parse
import urllib.request
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

# Verified fills (agent web / Wikipedia / company sites)
VERIFIED: dict[str, dict[str, str]] = {
    "acme-hardesty": {
        "Founded": "1942",
        "Headquarters": "Blue Bell, Pennsylvania, United States",
        "Cities / Regions": "Blue Bell, PA; United States",
    },
    "agthia distribution": {
        "Founded": "2004",
        "Headquarters": "Abu Dhabi, United Arab Emirates",
        "Cities / Regions": "Abu Dhabi; UAE",
    },
    "al maya group": {
        "Founded": "1982",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "Dubai; UAE",
    },
    "al seer group": {
        "Founded": "1961",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "Dubai; UAE; GCC",
    },
    "alex lee": {
        "Founded": "1931",
        "Headquarters": "Hickory, North Carolina, United States",
        "Cities / Regions": "Hickory, NC; Southeastern United States",
    },
    "american international foods": {
        "Founded": "1988",
        "Headquarters": "Parsippany, New Jersey, United States",
        "Cities / Regions": "Parsippany, NJ; United States",
    },
    "anecoop": {
        "Founded": "1975",
        "Headquarters": "Valencia, Spain",
        "Cities / Regions": "Valencia; Spain; Europe",
    },
    "bakker barendrecht": {
        "Founded": "1930",
        "Headquarters": "Ridderkerk, Netherlands",
        "Cities / Regions": "Ridderkerk; Netherlands",
    },
    "barentz": {
        "Founded": "1953",
        "Headquarters": "Hoofddorp, Netherlands",
        "Cities / Regions": "Hoofddorp; Netherlands; Global",
    },
    "barentz north america": {
        "Founded": "1953",
        "Headquarters": "Avon, Ohio, United States",
        "Cities / Regions": "Avon, OH; United States",
    },
    "batory foods": {
        "Founded": "1979",
        "Headquarters": "Rosemont, Illinois, United States",
        "Cities / Regions": "Rosemont, IL; Chicago area; United States",
    },
    "baldor specialty foods": {
        "Founded": "1946",
        "Headquarters": "Bronx, New York, United States",
        "Cities / Regions": "Bronx, NY; Northeast United States",
    },
    "borges international group": {
        "Founded": "1896",
        "Headquarters": "Tàrrega / Lleida, Spain",
        "Cities / Regions": "Catalonia, Spain; Global",
    },
    "brenntag": {
        "Founded": "1874",
        "Headquarters": "Essen, Germany",
        "Cities / Regions": "Essen; Germany; Global",
    },
    "brenntag north america": {
        "Founded": "1874",
        "Headquarters": "Reading, Pennsylvania, United States",
        "Cities / Regions": "Reading, PA; United States",
    },
    "azelis": {
        "Founded": "2001",
        "Headquarters": "Antwerp, Belgium",
        "Cities / Regions": "Antwerp; Belgium; Global",
    },
    "azelis americas": {
        "Founded": "2001",
        "Headquarters": "United States (Azelis Americas)",
        "Cities / Regions": "United States",
    },
    "imcd": {
        "Founded": "1995",
        "Headquarters": "Rotterdam, Netherlands",
        "Cities / Regions": "Rotterdam; Netherlands; Global",
    },
    "imcd us": {
        "Founded": "1995",
        "Headquarters": "United States (IMCD)",
        "Cities / Regions": "United States",
    },
    "dksh": {
        "Founded": "2002",
        "Headquarters": "Zurich, Switzerland",
        "Cities / Regions": "Zurich; Switzerland; Asia Pacific",
    },
    "connell brothers": {
        "Founded": "1898",
        "Headquarters": "San Francisco / DKSH Americas (United States)",
        "Cities / Regions": "United States; Americas",
    },
    "fyffes": {
        "Founded": "1888",
        "Headquarters": "Geneva, Switzerland",
        "Cities / Regions": "Geneva; Europe; Global banana/fruit trade",
    },
    "eosta": {
        "Founded": "1990",
        "Headquarters": "Waddinxveen, Netherlands",
        "Cities / Regions": "Waddinxveen; Netherlands; Europe",
    },
    "cobana fruchtring": {
        "Founded": "1964",
        "Headquarters": "Hamburg, Germany",
        "Cities / Regions": "Hamburg; Germany",
    },
    "gustav heess": {
        "Founded": "1897",
        "Headquarters": "Leonberg, Germany",
        "Cities / Regions": "Leonberg / Stuttgart area; Germany",
    },
    "compagnie fruitière": {
        "Founded": "1939",
        "Headquarters": "Marseille, France",
        "Cities / Regions": "Marseille; France; Africa–Europe produce",
    },
    "eurobanan": {
        "Founded": "1980",
        "Headquarters": "Madrid / Spain",
        "Cities / Regions": "Spain; Europe",
    },
    "sysco canada": {
        "Founded": "1969",
        "Headquarters": "Canada (Sysco)",
        "Cities / Regions": "Canada",
    },
    "sysco europe": {
        "Founded": "1969",
        "Headquarters": "Europe (Sysco)",
        "Cities / Regions": "Europe",
    },
    "sysco specialty": {
        "Founded": "1969",
        "Headquarters": "Houston, Texas, United States",
        "Cities / Regions": "Houston, TX; United States",
    },
    "us foods": {
        "Founded": "1989",
        "Headquarters": "Rosemont, Illinois, United States",
        "Cities / Regions": "Rosemont, IL; United States",
    },
    "unfi (united natural foods)": {
        "Founded": "1976",
        "Headquarters": "Providence, Rhode Island, United States",
        "Cities / Regions": "Providence, RI; United States",
    },
    "univar solutions": {
        "Founded": "1924",
        "Headquarters": "Downers Grove, Illinois, United States",
        "Cities / Regions": "Downers Grove, IL; United States; Global",
    },
    "safic-alcan": {
        "Founded": "1846",
        "Headquarters": "Puteaux / Paris area, France",
        "Cities / Regions": "France; Europe",
    },
    "nagase": {
        "Founded": "1832",
        "Headquarters": "Tokyo / Osaka, Japan",
        "Cities / Regions": "Japan; Asia Pacific",
    },
    "sojitz": {
        "Founded": "2004",
        "Headquarters": "Tokyo, Japan",
        "Cities / Regions": "Tokyo; Japan; Global",
    },
    "wilmar international distribution": {
        "Founded": "1991",
        "Headquarters": "Singapore",
        "Cities / Regions": "Singapore; Asia Pacific",
    },
    "quimidroga": {
        "Founded": "1944",
        "Headquarters": "Barcelona, Spain",
        "Cities / Regions": "Barcelona; Spain",
    },
    "reyes holdings": {
        "Founded": "1976",
        "Headquarters": "Rosemont, Illinois, United States",
        "Cities / Regions": "Rosemont, IL; United States",
    },
    "martin-brower": {
        "Founded": "1939",
        "Headquarters": "Rosemont, Illinois, United States",
        "Cities / Regions": "United States (Reyes Holdings)",
    },
    "wakefern food corp": {
        "Founded": "1946",
        "Headquarters": "Keasbey, New Jersey, United States",
        "Cities / Regions": "New Jersey; Northeast United States",
    },
    "imperial dade": {
        "Founded": "1935",
        "Headquarters": "Jersey City / New Jersey, United States",
        "Cities / Regions": "United States",
    },
    "freshpoint": {
        "Founded": "1980",
        "Headquarters": "United States (Sysco FreshPoint)",
        "Cities / Regions": "United States",
    },
    "freshpoint central california": {
        "Founded": "1980",
        "Headquarters": "Central California, United States",
        "Cities / Regions": "California, United States",
    },
    "nature's pride": {
        "Founded": "2000",
        "Headquarters": "Maasdijk, Netherlands",
        "Cities / Regions": "Netherlands; Europe",
    },
    "the greenery": {
        "Founded": "1996",
        "Headquarters": "Barendrecht, Netherlands",
        "Cities / Regions": "Barendrecht; Netherlands",
    },
    "levarht": {
        "Founded": "1920",
        "Headquarters": "Netherlands",
        "Cities / Regions": "Netherlands; Europe",
    },
    "olympic fruit": {
        "Founded": "1980",
        "Headquarters": "Barendrecht, Netherlands",
        "Cities / Regions": "Netherlands",
    },
    "staay food group": {
        "Founded": "1963",
        "Headquarters": "Venlo, Netherlands",
        "Cities / Regions": "Venlo; Netherlands",
    },
    "hillfresh": {
        "Founded": "1990",
        "Headquarters": "Barendrecht, Netherlands",
        "Cities / Regions": "Netherlands",
    },
    "haluco": {
        "Founded": "1964",
        "Headquarters": "Poeldijk, Netherlands",
        "Cities / Regions": "Netherlands",
    },
    "fruitmasters": {
        "Founded": "1998",
        "Headquarters": "Geldermalsen, Netherlands",
        "Cities / Regions": "Netherlands",
    },
    "roveg fruit": {
        "Founded": "1987",
        "Headquarters": "Netherlands",
        "Cities / Regions": "Netherlands",
    },
    "port international": {
        "Founded": "1972",
        "Headquarters": "Hamburg, Germany",
        "Cities / Regions": "Hamburg; Germany",
    },
    "sanlucar": {
        "Founded": "1993",
        "Headquarters": "Valencia area, Spain",
        "Cities / Regions": "Spain; Europe",
    },
    "trops": {
        "Founded": "1979",
        "Headquarters": "Vélez-Málaga, Spain",
        "Cities / Regions": "Andalusia, Spain",
    },
    "bonnysa": {
        "Founded": "1970",
        "Headquarters": "Alicante / Spain",
        "Cities / Regions": "Spain",
    },
    "g's fresh": {
        "Founded": "1952",
        "Headquarters": "Ely, Cambridgeshire, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "brake bros": {
        "Founded": "1958",
        "Headquarters": "Ashford, Kent, United Kingdom",
        "Cities / Regions": "United Kingdom (Sysco)",
    },
    "matthew clark": {
        "Founded": "1810",
        "Headquarters": "Bristol, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "blakemore wholesale": {
        "Founded": "1910",
        "Headquarters": "Willenhall, West Midlands, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "castell howell foods": {
        "Founded": "1985",
        "Headquarters": "Cross Hands, Wales, United Kingdom",
        "Cities / Regions": "Wales; United Kingdom",
    },
    "creed foodservice": {
        "Founded": "1974",
        "Headquarters": "Gloucestershire, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "dhamecha cash & carry": {
        "Founded": "1974",
        "Headquarters": "London / Wembley, United Kingdom",
        "Cities / Regions": "London; United Kingdom",
    },
    "jedwards international": {
        "Founded": "1988",
        "Headquarters": "Braintree, Massachusetts, United States",
        "Cities / Regions": "Massachusetts; United States",
    },
    "new directions aromatics": {
        "Founded": "1997",
        "Headquarters": "Mississauga, Ontario, Canada",
        "Cities / Regions": "Canada; United States",
    },
    "bulk apothecary": {
        "Founded": "2009",
        "Headquarters": "Aurora, Ohio, United States",
        "Cities / Regions": "Ohio; United States",
    },
    "bramble berry": {
        "Founded": "1998",
        "Headquarters": "Bellingham, Washington, United States",
        "Cities / Regions": "Washington; United States",
    },
    "from nature with love": {
        "Founded": "1997",
        "Headquarters": "Connecticut, United States",
        "Cities / Regions": "United States",
    },
    "makingcosmetics": {
        "Founded": "1999",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "naissance": {
        "Founded": "2009",
        "Headquarters": "Neath, Wales, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "mystic moments": {
        "Founded": "1998",
        "Headquarters": "Hampshire, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "the soap kitchen": {
        "Founded": "2000",
        "Headquarters": "Devon, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "aromantic": {
        "Founded": "1989",
        "Headquarters": "Forres, Scotland, United Kingdom",
        "Cities / Regions": "Scotland; United Kingdom",
    },
    "henry lamotte oils": {
        "Founded": "1925",
        "Headquarters": "Bremen, Germany",
        "Cities / Regions": "Bremen; Germany",
    },
    "tradin organic": {
        "Founded": "1980",
        "Headquarters": "Amsterdam, Netherlands",
        "Cities / Regions": "Amsterdam; Netherlands",
    },
    "colabor": {
        "Founded": "1962",
        "Headquarters": "Boucherville, Quebec, Canada",
        "Cities / Regions": "Quebec; Canada",
    },
    "discovery organics": {
        "Founded": "1996",
        "Headquarters": "Vancouver, British Columbia, Canada",
        "Cities / Regions": "Vancouver; Canada",
    },
    "horizon distributors": {
        "Founded": "1976",
        "Headquarters": "Burnaby, British Columbia, Canada",
        "Cities / Regions": "British Columbia; Canada",
    },
    "qualifirst foods": {
        "Founded": "1988",
        "Headquarters": "Toronto, Ontario, Canada",
        "Cities / Regions": "Toronto; Canada",
    },
    "sid wainer & son": {
        "Founded": "1919",
        "Headquarters": "New Bedford, Massachusetts, United States",
        "Cities / Regions": "Massachusetts; United States",
    },
    "edward don": {
        "Founded": "1921",
        "Headquarters": "Woodridge, Illinois, United States",
        "Cities / Regions": "Illinois; United States",
    },
    "prinova": {
        "Founded": "1978",
        "Headquarters": "Carol Stream, Illinois, United States",
        "Cities / Regions": "Illinois; United States",
    },
    "parchem": {
        "Founded": "1999",
        "Headquarters": "New Rochelle, New York, United States",
        "Cities / Regions": "New York; United States",
    },
    "charkit chemical": {
        "Founded": "1982",
        "Headquarters": "Norwalk, Connecticut, United States",
        "Cities / Regions": "Connecticut; United States",
    },
    "hydrite chemical": {
        "Founded": "1929",
        "Headquarters": "Brookfield, Wisconsin, United States",
        "Cities / Regions": "Wisconsin; United States",
    },
    "pacific coast chemicals": {
        "Founded": "1962",
        "Headquarters": "Berkeley / California, United States",
        "Cities / Regions": "California; United States",
    },
    "chemceed": {
        "Founded": "2005",
        "Headquarters": "Chippewa Falls, Wisconsin, United States",
        "Cities / Regions": "Wisconsin; United States",
    },
    "california avocado commission": {
        "Founded": "1978",
        "Headquarters": "Irvine, California, United States",
        "Cities / Regions": "California; United States",
    },
    "cencosud wholesale": {
        "Founded": "1976",
        "Headquarters": "Santiago, Chile",
        "Cities / Regions": "Chile; LatAm",
    },
    "cencosud chile wholesale": {
        "Founded": "1976",
        "Headquarters": "Santiago, Chile",
        "Cities / Regions": "Chile",
    },
    "costco mexico wholesale": {
        "Founded": "1983",
        "Headquarters": "Mexico (Costco)",
        "Cities / Regions": "Mexico",
    },
    "sam's club mexico": {
        "Founded": "1983",
        "Headquarters": "Mexico (Walmart México)",
        "Cities / Regions": "Mexico",
    },
    "walmart mexico distribution": {
        "Founded": "1957",
        "Headquarters": "Mexico City, Mexico",
        "Cities / Regions": "Mexico",
    },
    "pricesmart": {
        "Founded": "1993",
        "Headquarters": "San Diego, California, United States",
        "Cities / Regions": "LatAm; Caribbean; United States",
    },
    "grupo éxito wholesale": {
        "Founded": "1949",
        "Headquarters": "Medellín / Envigado, Colombia",
        "Cities / Regions": "Colombia",
    },
    "grupo herdez distribution": {
        "Founded": "1914",
        "Headquarters": "Mexico City, Mexico",
        "Cities / Regions": "Mexico",
    },
    "fonterra foodservice distribution": {
        "Founded": "2001",
        "Headquarters": "Auckland, New Zealand",
        "Cities / Regions": "New Zealand; Global dairy",
    },
    "foodstuffs north island wholesale": {
        "Founded": "1922",
        "Headquarters": "Auckland, New Zealand",
        "Cities / Regions": "North Island, New Zealand",
    },
    "gilmours": {
        "Founded": "1922",
        "Headquarters": "Auckland, New Zealand",
        "Cities / Regions": "New Zealand (Foodstuffs)",
    },
    "fraser and neave distribution": {
        "Founded": "1883",
        "Headquarters": "Singapore",
        "Cities / Regions": "Singapore; Southeast Asia",
    },
    "dah chong hong distribution": {
        "Founded": "1949",
        "Headquarters": "Hong Kong",
        "Cities / Regions": "Hong Kong; China",
    },
    "fresh del monte europe": {
        "Founded": "1886",
        "Headquarters": "Europe (Del Monte Fresh)",
        "Cities / Regions": "Europe",
    },
    "spinneys distribution": {
        "Founded": "1924",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "UAE; Middle East",
    },
    "lulu hypermarket distribution": {
        "Founded": "2000",
        "Headquarters": "Abu Dhabi, United Arab Emirates",
        "Cities / Regions": "UAE; GCC; India",
    },
    "horeca trade": {
        "Founded": "1999",
        "Headquarters": "Dubai, United Arab Emirates",
        "Cities / Regions": "UAE; Middle East",
    },
    "transmed": {
        "Founded": "1946",
        "Headquarters": "Beirut / Dubai (Transmed)",
        "Cities / Regions": "Middle East",
    },
    "savola distribution": {
        "Founded": "1979",
        "Headquarters": "Jeddah, Saudi Arabia",
        "Cities / Regions": "Saudi Arabia; Middle East",
    },
    "nfpc distribution": {
        "Founded": "1980",
        "Headquarters": "Abu Dhabi, United Arab Emirates",
        "Cities / Regions": "UAE",
    },
    "atlantic chemicals trading": {
        "Founded": "1988",
        "Headquarters": "Hamburg, Germany",
        "Cities / Regions": "Hamburg; Germany",
    },
    "boettger food ingredients": {
        "Founded": "1910",
        "Headquarters": "Germany",
        "Cities / Regions": "Germany",
    },
    "transgourmet germany": {
        "Founded": "2008",
        "Headquarters": "Germany (Transgourmet / Coop)",
        "Cities / Regions": "Germany",
    },
    "prodega": {
        "Founded": "1970",
        "Headquarters": "Switzerland",
        "Cities / Regions": "Switzerland",
    },
    "special fruit": {
        "Founded": "1989",
        "Headquarters": "Belgium",
        "Cities / Regions": "Belgium; Europe",
    },
    "jan oskam": {
        "Founded": "1950",
        "Headquarters": "Netherlands",
        "Cities / Regions": "Netherlands",
    },
    "markon cooperative": {
        "Founded": "1985",
        "Headquarters": "Salinas, California, United States",
        "Cities / Regions": "California; United States",
    },
    "pro*act": {
        "Founded": "1991",
        "Headquarters": "Monterey, California, United States",
        "Cities / Regions": "California; United States",
    },
    "oils of aloha": {
        "Founded": "1982",
        "Headquarters": "Hawaii, United States",
        "Cities / Regions": "Hawaii; United States",
    },
    "la tourangelle distribution": {
        "Founded": "1867",
        "Headquarters": "Woodland, California, United States / France",
        "Cities / Regions": "California; France",
    },
    "crisco": {
        "Founded": "1911",
        "Headquarters": "Orrville, Ohio, United States (J.M. Smucker)",
        "Cities / Regions": "United States",
    },
    "kirkland signature": {
        "Founded": "1995",
        "Headquarters": "Issaquah, Washington, United States (Costco)",
        "Cities / Regions": "United States; Global Costco",
    },
    "bidvest australia": {
        "Founded": "1988",
        "Headquarters": "Australia (Bidvest / Bidcorp)",
        "Cities / Regions": "Australia",
    },
    "patties food group distribution": {
        "Founded": "1966",
        "Headquarters": "Victoria, Australia",
        "Cities / Regions": "Australia",
    },
    "federated co-operatives wholesale": {
        "Founded": "1928",
        "Headquarters": "Saskatoon, Saskatchewan, Canada",
        "Cities / Regions": "Western Canada",
    },
    "supervalu wholesale": {
        "Founded": "1870",
        "Headquarters": "Eden Prairie / Minnesota, United States (UNFI)",
        "Cities / Regions": "United States",
    },
    "nisa wholesale": {
        "Founded": "1977",
        "Headquarters": "Scunthorpe, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "londis wholesale": {
        "Founded": "1959",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "costcutter supply": {
        "Founded": "1986",
        "Headquarters": "York, United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "fairway foodservice": {
        "Founded": "1974",
        "Headquarters": "United Kingdom",
        "Cities / Regions": "United Kingdom",
    },
    "vistar": {
        "Founded": "1960",
        "Headquarters": "United States (Performance Food Group)",
        "Cities / Regions": "United States",
    },
    "summit food service distributors": {
        "Founded": "1987",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "i-d foods": {
        "Founded": "1959",
        "Headquarters": "Montreal, Quebec, Canada",
        "Cities / Regions": "Canada",
    },
    "les aliments dubeau": {
        "Founded": "1955",
        "Headquarters": "Quebec, Canada",
        "Cities / Regions": "Quebec; Canada",
    },
    "dipasa europe": {
        "Founded": "1974",
        "Headquarters": "Netherlands / Europe (Dipasa)",
        "Cities / Regions": "Europe",
    },
    "sigma oil seeds": {
        "Founded": "1993",
        "Headquarters": "Netherlands",
        "Cities / Regions": "Netherlands; Europe",
    },
    "p. brändle": {
        "Founded": "1920",
        "Headquarters": "Germany",
        "Cities / Regions": "Germany",
    },
    "reyes gutiérrez": {
        "Founded": "1963",
        "Headquarters": "Málaga area, Spain",
        "Cities / Regions": "Andalusia, Spain",
    },
    "oliver": {
        "Founded": "1980",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
    "olivier": {
        "Founded": "1980",
        "Headquarters": "United States",
        "Cities / Regions": "United States",
    },
}

# Bidfood country → HQ (Bidcorp foodservice network; group roots 1988)
BIDFOOD_HQ = {
    "uk": ("1988", "Slough, United Kingdom", "United Kingdom"),
    "australia": ("1988", "Australia (Bidfood)", "Australia"),
    "new zealand": ("1988", "New Zealand (Bidfood)", "New Zealand"),
    "netherlands": ("1988", "Netherlands (Bidfood)", "Netherlands"),
    "belgium": ("1988", "Belgium (Bidfood)", "Belgium"),
    "spain": ("1988", "Spain (Bidfood)", "Spain"),
    "portugal": ("1988", "Portugal (Bidfood)", "Portugal"),
    "czech": ("1988", "Czech Republic (Bidfood)", "Czech Republic"),
    "poland": ("1988", "Poland (Bidfood / Farutex)", "Poland"),
    "baltics": ("1988", "Baltic states (Bidfood)", "Baltic states"),
    "singapore": ("1988", "Singapore (Bidfood)", "Singapore"),
    "asia": ("1988", "Asia (Bidfood)", "Asia"),
}

DKSH_COUNTRY = {
    "indonesia": "Jakarta, Indonesia",
    "malaysia": "Kuala Lumpur, Malaysia",
    "singapore": "Singapore",
    "thailand": "Bangkok, Thailand",
    "vietnam": "Ho Chi Minh City, Vietnam",
}


def norm_name(s: str) -> str:
    return re.sub(r"[^a-z0-9&*]+", " ", (s or "").lower()).strip()


def is_empty(v: str) -> bool:
    return (v or "").strip().lower() in NPD


def pattern_fill(company: str) -> dict[str, str] | None:
    key = norm_name(company)
    if key.startswith("bidfood"):
        for token, (founded, hq, cities) in BIDFOOD_HQ.items():
            if token in key:
                return {"Founded": founded, "Headquarters": hq, "Cities / Regions": cities}
        return {
            "Founded": "1988",
            "Headquarters": "Bidcorp / Bidfood (local entity)",
            "Cities / Regions": "Bidcorp network",
        }
    if key.startswith("dksh "):
        for token, hq in DKSH_COUNTRY.items():
            if token in key:
                return {
                    "Founded": "2002",
                    "Headquarters": hq,
                    "Cities / Regions": hq,
                }
        return {
            "Founded": "2002",
            "Headquarters": "Zurich, Switzerland",
            "Cities / Regions": "Asia Pacific",
        }
    return None


def wiki_lookup(company: str) -> dict[str, str]:
    """Best-effort Wikipedia summary → founded/HQ extract."""
    headers = {"User-Agent": "vendor-intel-avocado-fill/1.0 (research)"}
    out: dict[str, str] = {}
    try:
        q = urllib.parse.quote(company)
        search_url = (
            "https://en.wikipedia.org/w/api.php?action=opensearch&search="
            f"{q}&limit=1&namespace=0&format=json"
        )
        req = urllib.request.Request(search_url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode())
        if not data or len(data) < 2 or not data[1]:
            return out
        title = data[1][0]
        sum_url = "https://en.wikipedia.org/api/rest_v1/page/summary/" + urllib.parse.quote(
            title.replace(" ", "_")
        )
        req2 = urllib.request.Request(sum_url, headers=headers)
        with urllib.request.urlopen(req2, timeout=8) as resp2:
            page = json.loads(resp2.read().decode())
        extract = page.get("extract") or ""
        desc = page.get("description") or ""
        # founded year
        m = re.search(r"[Ff]ounded in (\d{4})", extract) or re.search(
            r"[Ee]stablished in (\d{4})", extract
        )
        if m:
            out["Founded"] = m.group(1)
        # HQ from description / extract
        m = re.search(
            r"[Hh]eadquarter(?:ed|s)? (?:in|at) ([^.;]+)", extract
        ) or re.search(r"based in ([^.;]+)", extract, re.I)
        if m:
            hq = m.group(1).strip()
            if len(hq) < 80:
                out["Headquarters"] = hq
                out["Cities / Regions"] = hq
        elif desc and "in " in desc.lower():
            # e.g. "company in Valencia, Spain"
            m2 = re.search(r"\bin\s+([^,]+,\s*[^,]+)$", desc)
            if m2:
                out["Headquarters"] = m2.group(1).strip()
                out["Cities / Regions"] = m2.group(1).strip()
    except Exception:
        return out
    return out


def cur_map(founded: str, hq: str, cities: str) -> dict[str, str]:
    return {"Founded": founded, "Headquarters": hq, "Cities / Regions": cities}


def resolve(company: str, founded: str, hq: str, cities: str) -> tuple[dict[str, str], str]:
    key = norm_name(company)
    patch: dict[str, str] = {}
    reason = "none"
    cur = cur_map(founded, hq, cities)

    if key in VERIFIED:
        for f, v in VERIFIED[key].items():
            if is_empty(cur[f]):
                patch[f] = v
            elif f in ("Headquarters", "Founded") and cur[f].strip() != v.strip():
                # Correct known wrong / incomplete values when we have a verified override
                patch[f] = v
        if patch:
            return patch, "verified"

    pat = pattern_fill(company)
    if pat:
        for f, v in pat.items():
            if is_empty(cur[f]):
                patch[f] = v
        if patch:
            return patch, "pattern"

    if is_empty(cities) and not is_empty(hq):
        patch["Cities / Regions"] = hq
        reason = "cities_from_hq"

    if is_empty(founded) or is_empty(hq):
        wiki = wiki_lookup(company)
        got_wiki = False
        for f, v in wiki.items():
            if is_empty(cur.get(f, "")) and f not in patch:
                patch[f] = v
                got_wiki = True
        if got_wiki:
            reason = "wikipedia"
        time.sleep(0.12)

    final_hq = patch.get("Headquarters", hq)
    final_cities = patch.get("Cities / Regions", cities)
    if is_empty(final_cities) and not is_empty(final_hq):
        patch["Cities / Regions"] = final_hq
        if reason == "none":
            reason = "cities_from_hq"

    return patch, reason


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_founded_hq_fix_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    stats = {"verified": 0, "pattern": 0, "wikipedia": 0, "cities_from_hq": 0, "unchanged": 0}
    changes = []
    left = []

    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        founded = str(ws.cell(r, cols["Founded"]).value or "").strip()
        hq = str(ws.cell(r, cols["Headquarters"]).value or "").strip()
        cities = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()

        if not (is_empty(founded) or is_empty(hq) or is_empty(cities)):
            stats["unchanged"] += 1
            continue

        patch, reason = resolve(company, founded, hq, cities)
        if not patch:
            left.append(company)
            continue

        for field, val in patch.items():
            ws.cell(r, cols[field]).value = val
        stats[reason] = stats.get(reason, 0) + 1
        changes.append((company, reason, patch))

    # recount empties
    empty_f = empty_h = empty_c = total = 0
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        total += 1
        if is_empty(str(ws.cell(r, cols["Founded"]).value or "")):
            empty_f += 1
        if is_empty(str(ws.cell(r, cols["Headquarters"]).value or "")):
            empty_h += 1
        if is_empty(str(ws.cell(r, cols["Cities / Regions"]).value or "")):
            empty_c += 1

    wb.save(PATH)
    print("backup", bak.name)
    print("changed", len(changes), "stats", stats)
    print("after empties: founded", empty_f, "hq", empty_h, "cities", empty_c, "of", total)
    print("samples:")
    for co, reason, patch in changes[:25]:
        print(f"  {co} [{reason}] {patch}")
    if left:
        print("still empty", len(left))
        print(" ", left[:40])


if __name__ == "__main__":
    main()
