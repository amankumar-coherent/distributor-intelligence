"""Fill blank HQ/Cities and normalize formats across top-10 expanded Excels.

Target formats:
  Headquarters: Street (optional), City, State/Region, Country
  Cities / Regions: City, State/Region, Country; coverage areas…
"""
from __future__ import annotations

import importlib.util
import re
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

spec = importlib.util.spec_from_file_location(
    "fillgeo", r"D:\vendor-intel\scripts\fill_founded_hq_cities_top10.py"
)
fillgeo = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fillgeo)
GEO: dict = fillgeo.GEO  # key -> (founded|None, hq, cities)

# Explicit overrides for empty / restored / known weak rows
OVERRIDE: dict[str, tuple[str, str]] = {
    "Shiwandistillery": (
        "Foshan, Guangdong, China",
        "Foshan, Guangdong, China; China rice wine and spirits export",
    ),
    "McLane Company": (
        "2121 Westgate Plaza, Temple, TX 76504, USA",
        "Temple, Texas, USA; USA nationwide grocery distribution",
    ),
    "Performance Food Group": (
        "12500 West Creek Parkway, Richmond, VA 23238, USA",
        "Richmond, Virginia, USA; USA foodservice",
    ),
    "Performance Food Group (PFG)": (
        "12500 West Creek Parkway, Richmond, VA 23238, USA",
        "Richmond, Virginia, USA; USA foodservice",
    ),
    "Bilstein": (
        "Ennepetal, North Rhine-Westphalia, Germany",
        "Ennepetal, Germany; Europe; global automotive/suspension",
    ),
    "Koni": (
        "Oud-Beijerland, South Holland, Netherlands",
        "Oud-Beijerland, Netherlands; Europe; global",
    ),
    "Sewing Machine Plus": (
        "713 Center Drive, San Marcos, CA 92069, USA",
        "San Marcos, California, USA; USA online and retail",
    ),
    "Elder Farms": (
        "Mount Vernon, Missouri, USA",
        "Mount Vernon, Missouri, USA; Ozarks region",
    ),
    "Elderberry Farms": (
        "Payson, Utah, USA",
        "Payson, Utah, USA",
    ),
    "C&S Wholesale Grocers": (
        "7 Corporate Drive, Keene, NH 03431, USA",
        "Keene, New Hampshire, USA; USA grocery wholesale",
    ),
    "Doctor Althea Global": (
        "Seoul, South Korea",
        "Seoul, South Korea; online international",
    ),
    "Bharucha Associates": (
        "Mumbai, Maharashtra, India",
        "Mumbai, Maharashtra, India; India medical distribution",
    ),
    "DM Distributors": (
        "10E Arctic Circle, Biddeford, ME 04005, USA",
        "Biddeford, Maine, USA; USA upholstery and textile supply",
    ),
    "Doray Kumas": (
        "Istanbul, Turkey",
        "Istanbul, Turkey; Turkey textile wholesale",
    ),
    "Anand Agencies (Chennai)": (
        "No. S2/61, Karpagam Nagar 2nd St, Selaiyur, Chennai, Tamil Nadu 600073, India",
        "Chennai, Tamil Nadu, India",
    ),
    "Wholesale Nail Supply": (
        "USA (wholesalenailsupply.com)",
        "USA online professional nail wholesale",
    ),
    "Gymedicare": (
        "Sihui Nanjiang Industrial Park, Zhaoqing, Guangdong, China",
        "Zhaoqing, Guangdong, China",
    ),
    "Hartmannsplantcompany": (
        "USA",
        "USA nursery and plant wholesale",
    ),
    "Pon.Bike distribution (Cannondale, Santa Cruz parent ops)": (
        "Nijmegen, Gelderland, Netherlands",
        "Nijmegen, Netherlands; global premium bike brands",
    ),
    "Sager Electronics": (
        "Middleborough, Massachusetts, USA",
        "Middleborough, Massachusetts, USA; Americas electronics distribution",
    ),
    "Owens & Minor": (
        "9120 Lockwood Boulevard, Mechanicsville, VA 23116, USA",
        "Mechanicsville, Virginia, USA; USA; international",
    ),
    "McKesson Medical-Surgical": (
        "6555 State Highway 161, Irving, TX 75039, USA",
        "Irving, Texas, USA; USA; Canada",
    ),
    "McKesson": (
        "6555 State Highway 161, Irving, TX 75039, USA",
        "Irving, Texas, USA; Richmond, Virginia, USA; Columbus, Ohio, USA; Canada",
    ),
    "McKesson (Consumer Health)": (
        "6555 State Highway 161, Irving, TX 75039, USA",
        "Irving, Texas, USA; USA; Canada",
    ),
    "Young's Market Company": (
        "Orange, California, USA",
        "Orange, California, USA; Western USA spirits wholesale",
    ),
    "Bunzl Healthcare / related MRO medical": (
        "London, United Kingdom",
        "London, United Kingdom; global healthcare distribution",
    ),
    "Paul Hartmann AG distribution": (
        "Heidenheim, Baden-Württemberg, Germany",
        "Heidenheim, Germany; Europe; global",
    ),
    "Mölnlycke Health Care distribution": (
        "Gothenburg, Västra Götaland, Sweden",
        "Gothenburg, Sweden; global hospital channel",
    ),
    "Douglas (Europe)": (
        "Düsseldorf, North Rhine-Westphalia, Germany",
        "Düsseldorf, Germany; Europe beauty retail",
    ),
    "Douglas Beauty Europe": (
        "Düsseldorf, North Rhine-Westphalia, Germany",
        "Düsseldorf, Germany; Europe beauty retail",
    ),
    "Watsons / AS Watson beauty retail": (
        "Hong Kong",
        "Hong Kong; Asia; Europe (AS Watson)",
    ),
    "Boots UK beauty": (
        "Nottingham, England, United Kingdom",
        "Nottingham, United Kingdom; UK; Ireland",
    ),
    "Gordon Food Service": (
        "Grand Rapids, Michigan, USA",
        "Grand Rapids, Michigan, USA; North America foodservice",
    ),
    "Aeon Group wholesale/retail": (
        "Chiba, Japan",
        "Chiba, Japan; Asia retail",
    ),
    "Nestlé Maggi distribution": (
        "Vevey, Vaud, Switzerland",
        "Vevey, Switzerland; global Maggi brand wholesale",
    ),
    "Coats Group Trade Distribution": (
        "Uxbridge, England, United Kingdom",
        "Uxbridge, United Kingdom; global",
    ),
    "Gutermann Distribution Partners": (
        "Gutach-Breisgau, Baden-Württemberg, Germany",
        "Gutach, Germany; global via distributors",
    ),
    "Prym Consumer USA / Haberdashery Dist.": (
        "Stolberg, North Rhine-Westphalia, Germany",
        "Stolberg, Germany; Europe; North America",
    ),
    "Liberty Fabrics (Wholesale)": (
        "London, England, United Kingdom",
        "London, United Kingdom; international wholesale",
    ),
    "Whaleys (Bradford) Ltd": (
        "Bradford, West Yorkshire, United Kingdom",
        "Bradford, United Kingdom; UK; Europe",
    ),
    "ROC-LON Fabrics (Rockland Industries)": (
        "Baltimore, Maryland, USA",
        "Baltimore, Maryland, USA; USA textile wholesale",
    ),
}

US_STATES = {
    "AL": "Alabama",
    "AK": "Alaska",
    "AZ": "Arizona",
    "AR": "Arkansas",
    "CA": "California",
    "CO": "Colorado",
    "CT": "Connecticut",
    "DE": "Delaware",
    "FL": "Florida",
    "GA": "Georgia",
    "HI": "Hawaii",
    "ID": "Idaho",
    "IL": "Illinois",
    "IN": "Indiana",
    "IA": "Iowa",
    "KS": "Kansas",
    "KY": "Kentucky",
    "LA": "Louisiana",
    "ME": "Maine",
    "MD": "Maryland",
    "MA": "Massachusetts",
    "MI": "Michigan",
    "MN": "Minnesota",
    "MS": "Mississippi",
    "MO": "Missouri",
    "MT": "Montana",
    "NE": "Nebraska",
    "NV": "Nevada",
    "NH": "New Hampshire",
    "NJ": "New Jersey",
    "NM": "New Mexico",
    "NY": "New York",
    "NC": "North Carolina",
    "ND": "North Dakota",
    "OH": "Ohio",
    "OK": "Oklahoma",
    "OR": "Oregon",
    "PA": "Pennsylvania",
    "RI": "Rhode Island",
    "SC": "South Carolina",
    "SD": "South Dakota",
    "TN": "Tennessee",
    "TX": "Texas",
    "UT": "Utah",
    "VT": "Vermont",
    "VA": "Virginia",
    "WA": "Washington",
    "WV": "West Virginia",
    "WI": "Wisconsin",
    "WY": "Wyoming",
    "DC": "District of Columbia",
}


def is_weak(v: object) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    if not s:
        return True
    low = s.lower()
    if low in {
        "n/a",
        "na",
        "none",
        "-",
        "—",
        "null",
        "unknown",
        "not confirmed publicly",
        "not specified",
        "global",
    }:
        return True
    if low.startswith("not listed") or low.startswith("n/a"):
        return True
    return False


def geo_lookup(company: str) -> tuple[str, str] | None:
    low = company.lower()
    # longest key match
    best = None
    best_len = 0
    for k, trip in GEO.items():
        if k in low and len(k) > best_len:
            hq = trip[1] or ""
            cities = trip[2] or ""
            if hq or cities:
                best = (hq, cities)
                best_len = len(k)
    return best


def normalize_hq(hq: str) -> str:
    s = str(hq or "").strip()
    if not s:
        return s
    # collapse whitespace
    s = re.sub(r"\s+", " ", s)
    # slash city pairs -> comma/semicolon style for HQ primary
    s = s.replace(" / ", ", ")
    # UK token
    s = re.sub(r"\bU\.K\.\b", "United Kingdom", s)
    s = re.sub(r",\s*UK$", ", United Kingdom", s)
    s = re.sub(r",\s*UK\b", ", United Kingdom", s)
    # United States -> USA
    s = re.sub(r"\bUnited States of America\b", "USA", s)
    s = re.sub(r"\bUnited States\b", "USA", s)
    # "City, ST" or "City, ST ZIP" without USA
    m = re.search(r",\s*([A-Z]{2})(?:\s+\d{5}(?:-\d{4})?)?$", s)
    if m and "USA" not in s.upper() and m.group(1) in US_STATES:
        s = s + ", USA"
    # "Payson, Utah" without country
    m2 = re.search(
        r",\s*("
        + "|".join(re.escape(x) for x in US_STATES.values())
        + r")$",
        s,
        re.I,
    )
    if m2 and "USA" not in s.upper() and "United States" not in s:
        s = s + ", USA"
    # expand trailing ", ST," before USA already handled; also "Keene, NH"
    m3 = re.match(r"^([^,]+),\s*([A-Z]{2})$", s)
    if m3 and m3.group(2) in US_STATES:
        s = f"{m3.group(1)}, {US_STATES[m3.group(2)]}, USA"
    # clean trailing junk emdash notes lightly — keep parenthetical HQ notes
    s = s.replace("—", " - ")
    return s.strip(" ,;")


def derive_cities_from_hq(hq: str) -> str:
    """Build a Cities/Regions string from HQ when cities is weak/vague."""
    s = normalize_hq(hq)
    if not s:
        return ""
    # strip street numbers start if long street address
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) >= 3 and re.match(r"^\d", parts[0]):
        # Street, City, State, Country
        city_parts = parts[1:]
        return ", ".join(city_parts)
    if len(parts) >= 2:
        return ", ".join(parts[-3:] if len(parts) > 3 else parts)
    return s


def normalize_cities(cities: str, hq: str) -> str:
    s = str(cities or "").strip()
    if is_weak(s) or s.upper() in {"USA", "UK", "GLOBAL", "ONLINE", "EUROPE"}:
        derived = derive_cities_from_hq(hq)
        if derived:
            s = derived
        elif s.upper() == "UK":
            s = "United Kingdom"
        elif s.upper() == "USA":
            s = "USA"
    s = re.sub(r"\s+", " ", s)
    s = s.replace(" / ", "; ")
    s = re.sub(r"\bU\.K\.\b", "United Kingdom", s)
    # trailing bare UK token in segments
    segs = [x.strip() for x in re.split(r";", s) if x.strip()]
    fixed = []
    for seg in segs:
        seg = re.sub(r"^UK$", "United Kingdom", seg)
        seg = re.sub(r",\s*UK$", ", United Kingdom", seg)
        seg = re.sub(r"\bUnited States\b", "USA", seg)
        # City, ST -> expand + USA
        m = re.match(r"^([^,]+),\s*([A-Z]{2})$", seg)
        if m and m.group(2) in US_STATES:
            seg = f"{m.group(1)}, {US_STATES[m.group(2)]}, USA"
        elif re.search(r",\s*([A-Z]{2})(?:\s+\d{5})?$", seg):
            mm = re.search(r",\s*([A-Z]{2})(?:\s+\d{5}(?:-\d{4})?)?$", seg)
            if mm and mm.group(1) in US_STATES and "USA" not in seg.upper():
                seg = seg + ", USA"
        fixed.append(seg)
    return "; ".join(fixed)


def resolve_for_company(company: str, hq: str, cities: str) -> tuple[str, str, str]:
    """Return (hq, cities, action_note)."""
    note = ""
    if company in OVERRIDE:
        oh, oc = OVERRIDE[company]
        return oh, oc, "override"

    # GEO map upgrade when current is weak or country-only / slashy
    lookup = geo_lookup(company)
    weak_now = is_weak(hq) or is_weak(cities)
    country_only = str(hq).strip().upper() in {
        "USA",
        "UK",
        "UNITED KINGDOM",
        "INDIA",
        "CHINA",
        "GERMANY",
        "FRANCE",
        "JAPAN",
        "AUSTRALIA",
        "CANADA",
        "SWEDEN",
        "NETHERLANDS",
        "TURKEY",
        "HONG KONG",
        "SWITZERLAND",
        "BELGIUM",
        "ITALY",
        "SPAIN",
        "EUROPE",
    }
    if lookup and (weak_now or country_only or " / " in str(hq)):
        lh, lc = lookup
        if lh and (weak_now or country_only or len(lh) > len(str(hq or ""))):
            hq = lh
            note = "geo_map"
        if lc and (is_weak(cities) or cities.upper() in {"USA", "UK", "GLOBAL"} or " / " in str(cities)):
            cities = lc
            note = note or "geo_map"

    hq_n = normalize_hq(hq)
    cities_n = normalize_cities(cities, hq_n)
    if hq_n != str(hq or "").strip() or cities_n != str(cities or "").strip():
        note = note or "normalized"
    return hq_n, cities_n, note


def process(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=False))
    hdr_i = next(
        i
        for i, r in enumerate(rows[:12])
        if r and r[0].value and str(r[0].value).strip() == "Company"
    )
    header = [str(c.value or "").strip() for c in rows[hdr_i]]
    ci = header.index("Company")
    hi = header.index("Headquarters")
    ri = header.index("Cities / Regions")
    si = header.index("Data Sources") if "Data Sources" in header else None

    stats = {"filled": 0, "normalized": 0, "unchanged": 0}
    examples = []

    for r in rows[hdr_i + 1 :]:
        company = str(r[ci].value or "").strip()
        if not company:
            continue
        old_hq = str(r[hi].value or "").strip()
        old_ci = str(r[ri].value or "").strip()
        new_hq, new_ci, note = resolve_for_company(company, old_hq, old_ci)
        if not new_hq and not new_ci:
            stats["unchanged"] += 1
            continue
        changed = False
        if new_hq and new_hq != old_hq:
            r[hi].value = new_hq
            changed = True
        if new_ci and new_ci != old_ci:
            r[ri].value = new_ci
            changed = True
        if changed:
            if is_weak(old_hq) or is_weak(old_ci) or note == "override":
                stats["filled"] += 1
            else:
                stats["normalized"] += 1
            if len(examples) < 8:
                examples.append(f"{company}: HQ '{old_hq[:40]}' -> '{new_hq[:40]}'")
            if si is not None and note:
                prev = str(r[si].value or "")
                tag = f"Geo {note}"
                if tag not in prev:
                    r[si].value = (prev + "; " if prev else "") + tag
        else:
            stats["unchanged"] += 1

    if "Run Summary" in wb.sheetnames:
        ws2 = wb["Run Summary"]
        ws2["A12"] = (
            f"HQ/Cities pass: filled={stats['filled']} normalized={stats['normalized']}"
        )

    wb.save(path)
    wb.close()
    stats["examples"] = examples
    return stats


def main() -> None:
    total_f = total_n = 0
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        st = process(path)
        total_f += st["filled"]
        total_n += st["normalized"]
        print(
            f"{path.parent.name}: filled={st['filled']} normalized={st['normalized']} unchanged={st['unchanged']}"
        )
        for e in st["examples"][:4]:
            print(f"  {e}")

    # audit remaining weak
    print("\nRemaining weak HQ/Cities:")
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        header = [str(h or "").strip() for h in data[hdr_i]]
        hi, ri = header.index("Headquarters"), header.index("Cities / Regions")
        weak = []
        for r in data[hdr_i + 1 :]:
            if not r or not r[0]:
                continue
            if is_weak(r[hi]) or is_weak(r[ri]):
                weak.append(str(r[0]))
        print(f"  {path.parent.name}: {len(weak)} weak" + (f" -> {weak}" if weak else ""))
        wb.close()
    print(f"\nTOTAL filled={total_f} normalized={total_n}")


if __name__ == "__main__":
    main()
