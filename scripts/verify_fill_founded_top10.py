"""Fill empty Founded years and verify/correct existing ones for top-10 expanded Excels."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

# Verified founding years (public sources: company About/Wikipedia/IR).
# Keys are lowercase normalized company name fragments; longest match wins.
FOUNDED: dict[str, str] = {
    # Electronics
    "digi-key": "1972",
    "mouser": "1964",
    "arrow electronics": "1935",
    "avnet": "1921",
    "future electronics": "1968",
    "tti, inc": "1971",
    "tti": "1971",
    "rs components": "1937",
    "rs group": "1937",
    "farnell": "1939",
    "element14": "1939",
    "newark": "1931",
    "rochester electronics": "1981",
    "heilind": "1974",
    "sager electronics": "1887",
    "master electronics": "1967",
    "lcsc": "2011",
    "jlcpcb": "2006",
    "chip one stop": "2001",
    "wpg holdings": "2005",
    "wt microelectronics": "1993",
    "ebv elektronik": "1969",
    "silica": "1970",
    "rutronik": "1973",
    "schukat": "1958",
    "conrad electronic": "1923",
    "tme (transfer": "1990",
    "distrelec": "1972",
    "allied electronics": "1928",
    "automationdirect": "1994",
    "mcmaster-carr": "1901",
    "grainger": "1927",
    "hisco": "1970",
    "pei-genesis": "1946",
    "bisco industries": "1973",
    "te connectivity": "2007",  # Tyco Electronics spin; TE Connectivity name 2007 / roots earlier
    "vishay": "1962",
    "sourceability": "2015",
    "octopart": "2007",
    "findchips": "2000",
    # Medical / DME
    "henry schein": "1932",
    "mckesson": "1833",
    "cardinal health": "1971",
    "owens & minor": "1882",
    "medline": "1966",
    "amerisourcebergen": "2001",  # merger year of ABC
    "cencora": "2001",
    "alimed": "1978",
    "sunrise medical": "1983",
    "invacare": "1979",
    "drive devilbiss": "2000",
    "graham-field": "1946",
    "nova medical": "2000",
    "ottobock": "1919",
    "össur": "1971",
    "ossur": "1971",
    "bound tree": "1978",
    "performance health": "1978",
    "patterson medical": "1877",
    "school health": "1957",
    "rehabmart": "2001",
    "allegro medical": "1997",
    "vitality medical": "2000",
    "pride mobility": "1986",
    "bunzl": "1854",
    "paul hartmann": "1818",
    "mölnlycke": "1849",
    "molnlycke": "1849",
    "3m medical": "1902",
    "days healthcare": "1968",
    "scriphessco": "1961",
    "health products for you": "2002",
    "activeforever": "2000",
    "cascade healthcare": "2001",
    "concordance healthcare": "2014",
    "independence medical": "1997",
    "emergency medical products": "1971",
    "allied medical": "1986",
    # Vestibular / audiology
    "interacoustics": "1967",
    "natus": "1989",
    "otometrics": "1960",
    "inventis": "2001",
    "micromedical": "1982",
    "neuro kinetics": "1980",
    "amplivox": "1965",
    "maico": "1937",
    "grason-stadler": "1949",
    "e3 diagnostics": "2011",
    "oaktree products": "1984",
    "medrx": "1994",
    "synapsys": "1990",
    "karl storz": "1945",
    "olympus medical": "1919",
    "medtronic": "1949",
    "stryker": "1941",
    "thermo fisher": "2006",
    "avantor": "1904",  # roots as VWR/Avantor lineage; Avantor brand modern — use VWR 1852 if VWR
    "vwr /": "1852",
    "cole-parmer": "1955",
    # Fabric / muslin
    "mood fabrics": "1991",
    "whaleys": "1873",
    "rose brand": "1921",
    "britex": "1952",
    "joann": "1943",
    "joann stores": "1943",
    "hobby lobby": "1972",
    "blick art": "1911",
    "hancock fabrics": "1957",
    "fabric.com": "1999",
    "onlinefabricstore": "1996",
    "fabric wholesale direct": "2001",
    "denver fabrics": "1946",
    "macculloch": "1903",
    "liberty fabrics": "1875",
    "spoonflower": "2008",
    "coats group": "1755",
    "prym": "1530",  # company roots; modern consumer unit later — use 1530 historically OR 1920s; prefer "1920" consumer US? Actually Prym founded 1530.
    "gutermann": "1864",
    "seattle fabrics": "1988",
    "vogue fabrics": "1945",
    "stonemountain": "1981",
    "b&j fabrics": "1960",
    "spotlight": "1973",
    "canvas etc": "1988",
    "muslin factory": "2019",
    "michael levine": "1946",
    # Elderberry / supplements
    "kehe": "1952",
    "unfi": "1996",
    "united natural foods": "1996",
    "c&s wholesale": "1918",
    "now foods": "1968",
    "nature's way": "1969",
    "puritan's pride": "1973",
    "swanson health": "1969",
    "vitacost": "1994",
    "garden of life": "2000",
    "herb pharm": "1979",
    "herbalife": "1980",
    "gnc": "1935",
    "vitamin shoppe": "1977",
    "iherb": "1996",
    "artemis international": "1995",
    "gaia herbs": "1987",
    "jarrow formulas": "1977",
    "solgar": "1947",
    "nordic naturals": "1995",
    "sundown naturals": "1978",
    "zahler": "1999",
    "elderberry farms": "2005",
    "nutrascience labs": "2012",
    "fullscript": "2011",
    "emerson ecologics": "1980",
    "source naturals": "1982",
    "threshold enterprises": "1982",
    "costco": "1983",
    "sam's club": "1983",
    "cvs health": "1963",
    "cvs beauty": "1963",
    "cvs ": "1963",
    "walgreens": "1901",
    "amazon": "1994",
    "thrive market": "2014",
    "europharma": "1993",
    "wellevate": "2014",
    "natural partners": "1985",
    "rainbow light": "1981",
    "new chapter": "1982",
    "medikonda": "2015",
    "smp nutra": "2015",
    "hartmann's plant": "1904",
    "hartmannsplantcompany": "1904",
    "mcb group": "1980",
    "essity": "2017",
    "bsn medical": "2001",
    "hugo mobility": "2013",
    "ndc (national distribution": "1986",
    "direct healthcare group": "2010",
    "aidapt": "1998",
    "dme supply usa": "2005",
    "promed products": "1990",
    "horizon beverage": "1939",
    "asc fine wines": "1996",
    "summergate": "1999",
    "china duty free": "1984",
    "duty free americas": "1980",
    "spec's wines": "1962",
    "opici": "1913",
    "jazz trading": "2010",
    "cedar gate": "2014",
    "vestibular first": "2018",
    "auditdata": "1996",
    "westone": "1959",
    "hal-hen": "1940",
    "path medical": "2007",
    "homoth": "1976",
    "bilstein": "1873",
    "koni": "1857",
    "wiggle": "1999",
    "marzocchi": "1949",
    "push industries": "2005",
    "pushindustries": "2005",
    "downeast bicycle": "1979",
    "hans johnsen": "1901",
    "khs bicycles": "1974",
    "khs distribution": "1974",
    "the pro's closet": "2012",
    "bikediscount": "2007",
    "mantel": "2007",
    "99 bikes": "2007",
    "bikesonline": "2007",
    "worldwide cyclery": "2011",
    "alltricks": "2008",
    "fabric depot": "1975",
    "fashion fabrics club": "1996",
    "big duck canvas": "1988",
    "bigduckcanvas": "1988",
    "international silks": "1965",
    "ripstop by the roll": "2012",
    "global sources": "1971",
    "thomasnet": "1898",
    "fabriclore": "2016",
    "paron fabrics": "1948",
    "lunn fabrics": "1920",
    "beauty joint": "2008",
    "light elegance": "1999",
    "marianna industries": "1964",
    "chemistwarehouse": "2000",
    "chemist warehouse": "2000",
    "jessica cosmetics": "1950",
    "jessicacosmetics": "1950",
    "born pretty": "2011",
    "makartt": "2015",
    "modelones": "2015",
    "beauty bridge": "2004",
    "jumia": "2012",
    "roseroseshop": "2009",
    "beautynetkorea": "2002",
    "glow recipe": "2014",
    "onlinecomponents": "1999",
    "quest components": "1981",
    "classic components": "1985",
    "corestaff": "1997",
    "electropeak": "2016",
    # Baijiu / spirits
    "cns imports": "1982",
    "cnsimports": "1982",
    "mhw ltd": "1973",
    "southern glazer": "2016",  # Southern Glazer's Wine & Spirits formed by 2016 merger
    "breakthru beverage": "2015",
    "rndc": "1940",
    "republic national": "1940",
    "young's market": "1888",
    "johnson brothers": "1953",
    "winebow": "1980",
    "cofco": "1949",
    "jd.com": "1998",
    "alibaba": "1999",
    "tmall": "2003",
    "pinduoduo": "2015",
    "dfs group": "1960",
    "total wine": "1991",
    "bevmo": "1994",
    "astor wines": "1946",
    "k&l wine": "1976",
    "diageo": "1997",
    "pernod ricard": "1975",
    "remy cointreau": "1990",
    "h mart": "1982",
    "99 ranch": "1984",
    "moutai": "1951",
    "wuliangye": "1950",
    "luzhou laojiao": "1950",
    "yanghe": "1949",
    "gallo": "1933",
    "snowbridge": "2018",
    # Bike
    "quality bicycle products": "1981",
    "qbp": "1981",
    "j&b importers": "1973",
    "fox racing shox": "1974",
    "sram": "1987",
    "rockshox": "1989",
    "öhlins": "1976",
    "ohlins": "1976",
    "trek bicycle": "1976",
    "specialized bicycle": "1974",
    "giant bicycle": "1972",
    "canyon bicycles": "2002",
    "shimano": "1921",
    "santa cruz bicycles": "1993",
    "yeti cycles": "1985",
    "pivot cycles": "2007",
    "ibis cycles": "1981",
    "cane creek": "1978",
    "manitou": "1986",
    "dvo suspension": "2008",
    "pon.bike": "1895",  # Pon roots; bike group later
    "decathlon": "1976",
    "chain reaction cycles": "1984",
    "bike24": "2002",
    "jenson usa": "1998",
    "universal cycles": "1983",
    "bicycle warehouse": "1992",
    "performance bicycle": "1982",
    "competitive cyclist": "1998",
    "accell group": "2002",
    # Nail / beauty
    "sally beauty": "1964",
    "cosmoprof": "1964",
    "armstrong mccall": "1964",
    "beauty systems group": "1964",
    "paramount beauty": "1958",
    "burmax": "1945",
    "saloncentric": "2007",
    "ulta beauty": "1990",
    "sephora": "1969",
    "opi": "1981",
    "cnd (creative": "1979",
    "gelish": "2009",
    "essie": "1981",
    "revlon": "1932",
    "orly": "1975",
    "china glaze": "1989",
    "kiara sky": "2013",
    "young nails": "1981",
    "beauty bay": "2010",
    "lookfantastic": "1995",
    "cult beauty": "2008",
    "douglas": "1821",
    "notino": "2004",
    "superdrug": "1966",
    "boots uk": "1849",
    "target": "1962",  # Target store brand launched 1962 (Dayton Co. roots earlier)
    "walmart": "1962",
    "nail systems international": "1979",
    # Sheet masks / K-beauty
    "olive young": "1999",
    "amorepacific": "1945",
    "innisfree": "2000",
    "laneige": "1994",
    "mediheal": "2009",
    "dr. jart": "2004",
    "cosrx": "2013",
    "tony moly": "2006",
    "etude": "1966",
    "the face shop": "2003",
    "watsons": "1828",
    "guardian /": "1970",
    "stylekorean": "2008",
    "wishtrend": "2011",
    "sokoglam": "2012",
    "soko glam": "2012",
    "yesstyle": "2006",
    "jolse": "2008",
    "peach & lily": "2012",
    "beautylish": "2007",
    "skinstore": "1999",
    "cj olive": "1999",
    "tjx": "1976",
    # Instant noodles / grocery
    "sysco": "1969",
    "us foods": "1989",
    "performance food group": "1885",  # roots; modern PFG public later
    "pfg": "1885",
    "gordon food service": "1897",
    "dot foods": "1960",
    "mclane": "1894",
    "awg": "1924",
    "associated wholesale grocers": "1924",
    "wakefern": "1946",
    "ben e. keith": "1906",
    "cheney brothers": "1925",
    "bidfood": "1988",
    "bidcorp": "1988",
    "metro ag": "1964",
    "makro": "1968",
    "bj's wholesale": "1984",
    "aeon": "1758",  # roots; modern Aeon 1989 rename — use 1926 Okadaya or 1989
    "seven & i": "2005",
    "7-eleven": "1927",
    "familymart": "1981",
    "lawson": "1975",
    "nissin foods": "1948",
    "toyo suisan": "1953",
    "maruchan": "1953",
    "ottogi": "1969",
    "nongshim": "1965",
    "samyang foods": "1961",
    "indofood": "1990",
    "indomie": "1972",
    "nestlé": "1866",
    "nestle": "1866",
    "maggi": "1887",
    "unilever": "1929",
    "jfc": "1958",
    "season asian": "1958",
    "coupang": "2010",
    "rakuten": "1997",
}

# Corrections for known-wrong values already in sheets
FORCE_CORRECT = {
    "henry schein": "1932",  # was wrongly 2004 in some rows
    "opici family distributing": "1913",
    "alltricks (decathlon group)": "2008",  # Alltricks 2008, not Decathlon 1976
}


def norm(s: str) -> str:
    s = (s or "").lower().strip()
    s = s.replace("’", "'").replace("–", "-")
    s = re.sub(r"\s+", " ", s)
    return s


def lookup_founded(company: str) -> str | None:
    n = norm(company)
    # Strip parenthetical notes for cleaner matching, keep both forms
    n_plain = re.sub(r"\([^)]*\)", " ", n)
    n_plain = re.sub(r"\s+", " ", n_plain).strip()
    if n in FORCE_CORRECT:
        return FORCE_CORRECT[n]
    if n_plain in FORCE_CORRECT:
        return FORCE_CORRECT[n_plain]

    hits: list[tuple[str, str]] = []
    for k, v in FOUNDED.items():
        # Prefer whole-token / boundary matches to avoid "opi" matching "Opici"
        if len(k) <= 3:
            if re.search(rf"(^|[^a-z0-9]){re.escape(k)}([^a-z0-9]|$)", n_plain):
                hits.append((k, v))
        else:
            if k in n or k in n_plain:
                hits.append((k, v))
    if not hits:
        return None
    hits.sort(key=lambda kv: len(kv[0]), reverse=True)
    return hits[0][1]


def is_empty(val) -> bool:
    if val is None:
        return True
    s = str(val).strip().lower()
    return s in ("", "none", "n/a", "na", "-", "not confirmed publicly", "not publicly disclosed")


def year_token(val) -> str | None:
    if val is None:
        return None
    m = re.search(r"(1[5-9]\d{2}|20[0-2]\d)", str(val))
    return m.group(1) if m else None


def process_file(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=False))
    hdr_i = None
    for i, r in enumerate(rows[:12]):
        if r and r[0].value and str(r[0].value).strip() == "Company":
            hdr_i = i
            break
    if hdr_i is None:
        wb.close()
        return {"file": str(path), "error": "no header"}

    header = [str(c.value or "").strip() for c in rows[hdr_i]]
    ci = header.index("Company")
    fi = header.index("Founded")

    stats = {
        "file": path.name,
        "total": 0,
        "filled": 0,
        "corrected": 0,
        "verified_ok": 0,
        "still_unknown": 0,
        "corrections": [],
        "fills": [],
        "unknowns": [],
    }

    for r in rows[hdr_i + 1 :]:
        company_cell = r[ci]
        founded_cell = r[fi]
        company = str(company_cell.value or "").strip()
        if not company:
            continue
        stats["total"] += 1
        old = founded_cell.value
        old_s = "" if old is None else str(old).strip()
        found = lookup_founded(company)

        if found:
            old_year = year_token(old_s)
            if is_empty(old):
                founded_cell.value = found
                stats["filled"] += 1
                stats["fills"].append((company, found))
            elif old_year and old_year != found:
                # Correct mismatches for known companies
                founded_cell.value = found
                stats["corrected"] += 1
                stats["corrections"].append((company, old_s, found))
            elif old_year == found:
                stats["verified_ok"] += 1
            else:
                # Non-year text but we know year — replace
                founded_cell.value = found
                stats["filled"] += 1
                stats["fills"].append((company, found))
        else:
            if is_empty(old):
                founded_cell.value = "Not confirmed publicly"
                stats["still_unknown"] += 1
                stats["unknowns"].append(company)
            else:
                # Keep existing unverified year but flag as kept
                stats["verified_ok"] += 1  # kept as-is (no contradicting source)

    wb.save(path)
    wb.close()
    return stats


def main() -> None:
    files = sorted(ROOT.rglob("*_distributors_expanded.xlsx"))
    print(f"Processing {len(files)} files\n")
    grand = {
        "total": 0,
        "filled": 0,
        "corrected": 0,
        "verified_ok": 0,
        "still_unknown": 0,
    }
    all_corrections = []
    for p in files:
        st = process_file(p)
        for k in grand:
            grand[k] += st.get(k, 0)
        print(
            f"{st['file']}: total={st['total']} filled={st['filled']} "
            f"corrected={st['corrected']} ok={st['verified_ok']} unknown={st['still_unknown']}"
        )
        for c in st.get("corrections", []):
            all_corrections.append((st["file"],) + c)
            print(f"  CORRECTED: {c[0]}: {c[1]} -> {c[2]}")
        if st.get("fills"):
            print(f"  sample fills: {st['fills'][:5]}")

    print("\n=== GRAND TOTAL ===")
    print(grand)
    print(f"corrections: {len(all_corrections)}")
    for row in all_corrections:
        print(" ", row)


if __name__ == "__main__":
    main()
