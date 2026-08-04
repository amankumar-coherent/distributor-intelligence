"""Fill/verify Regions Served and Distribution Type for top-10 expanded Excels."""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DEMO = Path(r"C:\Users\aman.kumar\Desktop\Demo Output")
SRC = Path(r"D:\vendor-intel\output")

# Company overrides: (regions_served, distribution_type)
COMPANY_OVERRIDES: dict[str, tuple[str, str]] = {
    # Healthcare / DME
    "McKesson": ("USA (nationwide pharmaceutical & medical distribution)", "Pharmacy / medical wholesale"),
    "McKesson Medical-Surgical": ("USA", "Pharmacy / medical wholesale"),
    "McKesson (Consumer Health)": ("USA", "Pharmacy / medical wholesale"),
    "Cardinal Health": ("USA (hospitals, clinics, pharmacies)", "Pharmacy / medical wholesale"),
    "Cardinal Health Medical": ("USA", "Pharmacy / medical wholesale"),
    "Independence Medical (Cardinal specialty)": ("USA", "Specialty medical wholesale"),
    "AmerisourceBergen": ("USA; selective international", "Pharmacy / medical wholesale"),
    "AmerisourceBergen / Cencora": ("USA; selective international", "Pharmacy / medical wholesale"),
    "Henry Schein": ("USA; international (global dental/medical)", "Professional medical wholesale"),
    "Henry Schein Medical": ("USA; international", "Professional medical wholesale"),
    "Henry Schein Medical (ENT/audiology)": ("USA; international", "Professional medical wholesale"),
    "Owens & Minor": ("USA; selective international", "Provider logistics / medical wholesale"),
    "Medline Industries": ("USA; global export", "Manufacturer + wholesale distribution"),
    "Medline Industries,": ("USA; global export", "Manufacturer + wholesale distribution"),
    "Medline": ("USA; global export", "Manufacturer + wholesale distribution"),
    "AliMed": ("USA", "Specialty clinical / rehab wholesale"),
    "Health Products For You": ("USA (online national)", "Online retail / e-commerce"),
    "Activeforever": ("USA (online national)", "Online retail / e-commerce"),
    "ScripHessco": ("USA", "Chiropractic / rehab wholesale"),
    "Sunrise Medical": ("Global (dealer network)", "Brand owner + authorized dealer distribution"),
    "Sunrise Medical dealer network": ("Global (authorized dealers)", "Authorized dealer distribution"),
    "Invacare dealer network": ("Americas; Europe; selective global", "Authorized dealer distribution"),
    "Ottobock dealer network": ("Global", "Authorized dealer distribution"),
    "Össur distribution partners": ("Global", "Authorized brand distribution"),
    "Drive DeVilbiss Healthcare distribution": ("USA; international", "Brand owner distribution"),
    "Amazon Medical Supplies": ("USA; international marketplace", "Marketplace / e-commerce"),
    "Walmart Mobility / OTC": ("USA", "Mass retail"),
    "CVS Health DME/OTC": ("USA", "Pharmacy retail"),
    "Walgreens DME/OTC": ("USA", "Pharmacy retail"),
    "Bharucha Associates": ("India; export", "Medical wholesale / export"),
    # Baijiu
    "JD.com": ("China (online nationwide); selective cross-border", "Online retail / marketplace"),
    "JD.com Fresh / Spirits channel": ("China online", "Online retail / marketplace"),
    "Cnsimports": ("USA (multi-state spirits wholesale)", "Import wholesale"),
    "CNS Imports": ("USA (multi-state spirits wholesale)", "Import wholesale"),
    "Southern Glazer's Wine & Spirits": ("USA (multi-state)", "Beverage wholesale"),
    "Breakthru Beverage Group": ("USA; Canada", "Beverage wholesale"),
    "Republic National Distributing Company (RNDC)": ("USA (multi-state)", "Beverage wholesale"),
    "Young's Market Company": ("Western USA", "Beverage wholesale"),
    "Winebow Group": ("USA", "Specialty import wholesale"),
    "DFS Group (travel retail)": ("Global travel retail", "Travel retail"),
    "China Duty Free Group": ("China travel retail; international airports", "Travel retail"),
    "Total Wine & More": ("USA", "Specialty retail"),
    "Spec's Wines, Spirits & Finer Foods": ("Texas, USA", "Specialty retail"),
    "BevMo!": ("Western USA", "Specialty retail"),
    "Alibaba / Tmall Liquor Marketplace": ("China; cross-border", "Marketplace / platform"),
    "Pinduoduo Liquor Channel": ("China", "Marketplace / platform"),
    "Kweichow Moutai distribution subsidiaries": ("China; export / duty-free", "Brand owner sales network"),
    "Wuliangye Yibin sales companies": ("China; export", "Brand owner sales network"),
    "Luzhou Laojiao sales channels": ("China; export", "Brand owner sales network"),
    "Yanghe Distillery sales network": ("China; export", "Brand owner sales network"),
    "Shiwandistillery": ("China; export markets", "Manufacturer export / wholesale"),
    "COFCO Wines & Spirits": ("China; selective international", "Brand / wholesale distribution"),
    "Pernod Ricard China / Asia distribution": ("China; Asia", "Brand owner distribution"),
    "Diageo Greater China distribution": ("China; Asia", "Brand owner distribution"),
    # Elderberry / CPG
    "KeHE Distributors": ("USA; Canada", "Natural grocery wholesale"),
    "UNFI (United Natural Foods)": ("USA; Canada", "Natural grocery wholesale"),
    "UNFI": ("USA; Canada", "Natural / specialty grocery wholesale"),
    "C&S Wholesale Grocers": ("USA", "Grocery wholesale"),
    "Garden of Life": ("USA; international (Nestlé Health Science)", "Brand manufacturer distribution"),
    "NOW Foods": ("USA; international", "Brand manufacturer / wholesale"),
    "Nature's Way": ("USA; international", "Brand manufacturer distribution"),
    "Herb Pharm": ("USA; selective export", "Brand manufacturer / wholesale"),
    "Herbalife": ("Global (direct selling markets)", "Direct selling / brand distribution"),
    "Puritan's Pride": ("USA; international DTC", "DTC + wholesale brand"),
    "Swanson Health Products": ("USA; export", "DTC + wholesale"),
    "Vitacost": ("USA online; selective ship-to", "Online retail / e-commerce"),
    "Vitacost (Kroger)": ("USA online", "Online retail / e-commerce"),
    "iHerb Wholesale / B2B channel": ("Global online", "Online retail / B2B e-commerce"),
    "The Vitamin Shoppe Distribution": ("USA", "Specialty retail + wholesale"),
    "GNC Distribution": ("USA; international franchise", "Specialty retail + wholesale"),
    "Costco Wholesale": ("USA; Canada; selective international", "Warehouse club retail"),
    "Costco Wholesale (Kirkland / brand aisle)": ("USA; Canada; selective international", "Warehouse club retail"),
    "Sam's Club / Walmart Distribution": ("USA", "Warehouse club / mass retail"),
    # Instant noodles
    "Walmart": ("USA; international retail markets", "Mass retail"),
    "Walmart grocery distribution": ("USA; international retail markets", "Mass retail"),
    "Sysco": ("USA; Canada; Europe; selective global foodservice", "Foodservice wholesale"),
    "Sysco Corporation": ("USA; Canada; Europe; selective global foodservice", "Foodservice wholesale"),
    "Sysco Asia / regional foodservice": ("Asia foodservice markets", "Foodservice wholesale"),
    "US Foods": ("USA", "Foodservice wholesale"),
    "Performance Food Group": ("USA", "Foodservice wholesale"),
    "Performance Food Group (PFG)": ("USA", "Foodservice wholesale"),
    "Gordon Food Service": ("USA; Canada", "Foodservice wholesale"),
    "McLane Company": ("USA", "Grocery / convenience wholesale"),
    "Dot Foods": ("USA", "Foodservice redistributor wholesale"),
    "Nissin Foods sales companies": ("Global brand markets", "Brand owner sales network"),
    "Toyo Suisan / Maruchan distribution": ("USA; Japan; export", "Brand owner distribution"),
    "Nestlé Maggi distribution": ("Global Maggi markets", "Brand owner distribution"),
    "Indofood / Indomie distribution": ("Indonesia; global export", "Brand owner distribution"),
    "Samyang Foods distribution": ("Korea; global export", "Brand owner distribution"),
    "Nongshim distribution": ("Korea; USA; global export", "Brand owner distribution"),
    "Amazon Grocery / Subscribe & Save": ("USA; selective international", "Online retail / marketplace"),
    "H Mart distribution": ("USA; Canada (Asian grocery)", "Specialty grocery retail / wholesale"),
    "99 Ranch Market": ("USA (Asian grocery)", "Specialty grocery retail"),
    # MTB
    "Chain Reaction Cycles": ("Global online (EU/UK focus + export)", "Online retail / e-commerce"),
    "Chain Reaction Cycles / Wiggle (SIGNA)": ("Global online", "Online retail / e-commerce"),
    "Wiggle": ("Global online (EU/UK focus)", "Online retail / e-commerce"),
    "Bike24": ("Europe; international online", "Online retail / e-commerce"),
    "Quality Bicycle Products (QBP)": ("USA (dealer wholesale)", "Specialty bicycle wholesale"),
    "Fox Racing Shox dealer/distributor network": ("Global", "Authorized brand distribution"),
    "SRAM / RockShox distribution": ("Global", "Authorized brand distribution"),
    "Öhlins Bicycle distribution partners": ("Global (performance markets)", "Authorized brand distribution"),
    "Ohlins": ("Global", "Brand manufacturer distribution"),
    "Trek Bicycle dealer network": ("Global", "Authorized dealer distribution"),
    "Specialized Bicycle dealer network": ("Global", "Authorized dealer distribution"),
    "Giant Bicycle dealer distribution": ("Global", "Authorized dealer distribution"),
    "Decathlon Bike Parts Channel": ("Global Decathlon markets", "Mass specialty retail"),
    "Alltricks (Decathlon group)": ("Europe online", "Online retail / e-commerce"),
    "Competitive Cyclist / Backcountry": ("USA; international online", "Online retail / e-commerce"),
    "Jenson USA": ("USA; international online", "Online retail / e-commerce"),
    "Bicycle Warehouse": ("USA (regional retail)", "Specialty retail"),
    "Performance Bicycle": ("USA (legacy / related channels)", "Specialty retail"),
    "99 Bikes (AU retail chain)": ("Australia", "Specialty retail"),
    "BikesOnline / DTC AU-NZ": ("Australia; New Zealand; online export", "Online retail / DTC"),
    # Muslin / fabric
    "JOANN (Fabric & Craft)": ("USA", "Craft retail"),
    "Joann Stores": ("USA", "Craft retail"),
    "Hobby Lobby Fabric": ("USA", "Craft retail"),
    "Spotlight (Australia/NZ)": ("Australia; New Zealand", "Craft retail"),
    "Mood Fabrics": ("USA; online international", "Fashion fabric retail / wholesale"),
    "Fabric.com": ("USA online", "Online retail / e-commerce"),
    "Fabric.com (Amazon)": ("USA online", "Online retail / marketplace"),
    "ROC-LON Fabrics (Rockland Industries)": ("USA; export", "Mill / brand wholesale"),
    "Liberty Fabrics (Wholesale)": ("UK; global wholesale", "Brand wholesale"),
    "Rose Brand": ("USA; theatrical export", "Theatrical textile wholesale / retail"),
    "Blick Art Materials": ("USA; online international", "Art materials retail / wholesale"),
    # Nail
    "CosmoProf (Beauty Systems Group)": ("USA; Canada", "Professional beauty wholesale"),
    "Sally Beauty Supply": ("USA; Canada; selective international", "Beauty retail"),
    "SalonCentric (L'Oréal)": ("USA", "Professional salon wholesale"),
    "OPI authorized distributors": ("Global OPI markets", "Authorized brand distribution"),
    "CND (Creative Nail Design) distribution": ("Global CND markets", "Authorized brand distribution"),
    "Gelish / Hand & Nail Harmony distributors": ("USA; international", "Authorized brand distribution"),
    "Ulta Beauty": ("USA", "Specialty beauty retail"),
    "Sephora": ("USA; Canada; selective international", "Prestige beauty retail"),
    "Beauty Bay": ("UK; Europe; international online", "Online beauty retail"),
    "Lookfantastic / THG Beauty": ("UK; international online", "Online beauty retail"),
    "Notino": ("Europe online", "Online beauty retail"),
    "Douglas (Europe)": ("Europe", "Beauty retail"),
    "Chemistwarehouse": ("Australia; New Zealand", "Pharmacy / beauty retail"),
    # Sheet masks
    "Soko Glam": ("USA; international online", "K-beauty specialty retail"),
    "Sokoglam": ("USA; international online", "K-beauty specialty retail"),
    "YesStyle": ("Global online", "Online Asian beauty retail"),
    "Olive Young Global": ("Korea; global online", "K-beauty retail / e-commerce"),
    "StyleKorean": ("Global online", "Online K-beauty retail"),
    "Innisfree": ("Korea; global Amorepacific markets", "Brand retail / distribution"),
    "The Face Shop": ("Korea; global LG Beauty markets", "Brand retail / distribution"),
    "Target": ("USA", "Mass retail"),
    "Target Beauty": ("USA", "Mass retail"),
    "Target Beauty masks": ("USA", "Mass retail"),
    "Peach & Lily": ("USA; online international", "K-beauty specialty retail"),
    "SkinStore": ("USA; UK; international online", "Online beauty retail"),
    "iHerb Beauty": ("Global online", "Online retail / e-commerce"),
    "Watsons / AS Watson": ("Asia; Europe (AS Watson markets)", "Drugstore / beauty retail"),
    "Guardian / AS Watson": ("Asia (AS Watson markets)", "Drugstore / beauty retail"),
    "Boots UK beauty": ("UK; Ireland; selective international", "Pharmacy / beauty retail"),
    "Douglas Beauty Europe": ("Europe", "Beauty retail"),
    # Electronics / shunt
    "Digi-Key Electronics": ("Global (catalog shipping worldwide)", "Catalog electronics distribution"),
    "Mouser Electronics": ("Global (223+ countries/territories)", "Catalog electronics distribution"),
    "Avnet": ("Global (Americas, EMEA, Asia)", "Global electronics distribution"),
    "Arrow Electronics": ("Global", "Global electronics distribution"),
    "Future Electronics": ("Global", "Global electronics distribution"),
    "TTI, Inc.": ("Americas; Europe; Asia", "Specialty electronics distribution"),
    "RS Components": ("UK; Europe; Asia; global catalog", "Catalog industrial / electronics distribution"),
    "RS Components / RS Group": ("UK; Europe; Asia; global catalog", "Catalog industrial / electronics distribution"),
    "Farnell / element14": ("Europe; Asia; Americas", "Catalog electronics distribution"),
    "Newark": ("Americas", "Catalog electronics distribution"),
    "Allied Electronics & Automation": ("Americas", "Industrial / electronics distribution"),
    "LCSC Electronics": ("China; global online", "Catalog electronics / online distribution"),
    "WPG Holdings": ("Asia-Pacific; selective global", "Electronics distribution"),
    "Rutronik": ("Europe; selective international", "Electronics distribution"),
    "TE Connectivity": ("Global OEM markets", "Manufacturer direct / OEM distribution"),
    "Vishay authorized distributor network": ("Global via authorized distributors", "Authorized brand distribution"),
    "Sager Electronics": ("Americas", "Specialty electronics distribution"),
    "McMaster-Carr": ("USA; selective international", "Industrial catalog distribution"),
    "Grainger": ("USA; Canada; selective international", "Industrial MRO distribution"),
}

# Business Type -> Distribution Type
BIZ_TO_DIST: dict[str, str] = {
    "Wholesaler": "Wholesale / B2B distribution",
    "Distributor": "B2B distribution",
    "Brand Distribution Network": "Brand owner / authorized distribution network",
    "Retailer": "Retail store distribution",
    "Online Retailer": "Online retail / e-commerce",
    "Manufacturer / Distributor": "Manufacturer + wholesale distribution",
    "Marketplace": "Marketplace / platform",
    "Wholesaler / Retailer": "Wholesale + retail hybrid",
    "Importer / Distributor": "Import wholesale / distribution",
    "Authorized Distributor": "Authorized brand distribution",
    "Healthcare Technology / Channel": "Healthcare technology / channel partner",
}

COUNTRY_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"\b(united states|u\.s\.a\.|u\.s\.|\busa\b)", re.I), "USA"),
    (re.compile(r"\bcanada\b", re.I), "Canada"),
    (re.compile(r"\bmexico\b", re.I), "Mexico"),
    (re.compile(r"\buk\b|\bunited kingdom\b|\bengland\b|\blondon\b", re.I), "UK"),
    (re.compile(r"\bgermany\b|\bdeutschland\b", re.I), "Germany"),
    (re.compile(r"\bfrance\b", re.I), "France"),
    (re.compile(r"\bitaly\b", re.I), "Italy"),
    (re.compile(r"\bnetherlands\b|\bholland\b", re.I), "Netherlands"),
    (re.compile(r"\bswitzerland\b", re.I), "Switzerland"),
    (re.compile(r"\bsweden\b", re.I), "Sweden"),
    (re.compile(r"\baustralia\b", re.I), "Australia"),
    (re.compile(r"\bnew zealand\b", re.I), "New Zealand"),
    (re.compile(r"\bchina\b|\bbeijing\b|\bshanghai\b|\bguangdong\b|\bfoshan\b|\bshenzhen\b", re.I), "China"),
    (re.compile(r"\bjapan\b|\btokyo\b", re.I), "Japan"),
    (re.compile(r"\bkorea\b|\bseoul\b", re.I), "South Korea"),
    (re.compile(r"\bindia\b|\bmumbai\b|\bdelhi\b|\bchennai\b|\bkarur\b", re.I), "India"),
    (re.compile(r"\bsingapore\b", re.I), "Singapore"),
    (re.compile(r"\btaiwan\b", re.I), "Taiwan"),
    (re.compile(r"\bhong kong\b", re.I), "Hong Kong"),
    (re.compile(r"\bbrazil\b", re.I), "Brazil"),
    (re.compile(r"\buae\b|\bdubai\b", re.I), "UAE"),
    (re.compile(r"\bisrael\b", re.I), "Israel"),
]

US_STATE = re.compile(
    r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|"
    r"MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY)\b|"
    r"\b(Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|"
    r"Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|"
    r"Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|"
    r"New Hampshire|New Jersey|New Mexico|New York|North Carolina|North Dakota|Ohio|Oklahoma|"
    r"Oregon|Pennsylvania|Rhode Island|South Carolina|South Dakota|Tennessee|Texas|Utah|"
    r"Vermont|Virginia|Washington|West Virginia|Wisconsin|Wyoming)\b",
    re.I,
)


def is_weak(v: object) -> bool:
    s = str(v or "").strip().lower()
    if not s:
        return True
    if s.startswith("not ") or s in {"n/a", "na", "unknown", "-", "—"}:
        return True
    return False


def needs_dist_upgrade(v: object, biz: str) -> bool:
    """Upgrade overly generic distribution types when Business Type is more specific."""
    s = str(v or "").strip().lower()
    if is_weak(v):
        return True
    if s in {"distributor", "wholesale"}:
        return True
    if s == "wholesale / b2b distribution" and biz in {
        "Online Retailer",
        "Retailer",
        "Marketplace",
        "Brand Distribution Network",
        "Importer / Distributor",
        "Manufacturer / Distributor",
        "Authorized Distributor",
        "Healthcare Technology / Channel",
    }:
        return True
    if s == "direct retail / retailer":
        return True
    return False


def _norm(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()
    s = re.sub(r"\b(inc|llc|ltd|co|corp|corporation|company|group|ag|plc)\b", "", s)
    return re.sub(r"\s+", " ", s).strip()


_OVERRIDE_NORM = {_norm(k): v for k, v in COMPANY_OVERRIDES.items()}


def lookup_override(company: str) -> tuple[str, str] | None:
    if company in COMPANY_OVERRIDES:
        return COMPANY_OVERRIDES[company]
    nk = _norm(company)
    if nk in _OVERRIDE_NORM:
        return _OVERRIDE_NORM[nk]
    for key, val in COMPANY_OVERRIDES.items():
        kn = _norm(key)
        if kn and (kn in nk or nk in kn) and min(len(kn), len(nk)) >= 10:
            return val
    return None


def infer_regions(hq: str, cities: str, company: str, biz: str) -> str:
    text = f"{hq} {cities} {company}"
    found: list[str] = []
    for pat, label in COUNTRY_PATTERNS:
        if pat.search(text) and label not in found:
            found.append(label)
    if not found and US_STATE.search(text):
        found.append("USA")

    low = text.lower()
    if any(x in low for x in ("global", "worldwide", "international", "export")):
        if found == ["USA"]:
            return "USA; international"
        if found == ["China"]:
            return "China; export / international"
        if found:
            return "; ".join(found) + "; international"
        return "Global"

    if "online" in low or biz == "Online Retailer":
        if found:
            base = "; ".join(found)
            if "online" not in base.lower():
                return f"{base} (online / ship-to markets)"
            return base
        return "Online / multi-region e-commerce"

    if found:
        return "; ".join(found)

    # last resort from cities alone
    cities_c = str(cities or "").strip()
    if cities_c and not cities_c.lower().startswith("not "):
        # take first geographic clause
        part = re.split(r"[;|/]", cities_c)[0].strip()
        if part:
            return part[:80]
    return "Not publicly specified"


def infer_dist_type(biz: str, company: str) -> str:
    if biz in BIZ_TO_DIST:
        return BIZ_TO_DIST[biz]
    low = (company or "").lower()
    if "marketplace" in low or "tmall" in low or "amazon" in low:
        return "Marketplace / platform"
    if "dealer network" in low or "authorized" in low:
        return "Authorized brand / dealer distribution"
    if "distribution" in low and "sales" in low:
        return "Brand owner sales network"
    return "Wholesale / B2B distribution"


def resolve(company: str, biz: str, hq: str, cities: str, regions, dist) -> tuple[str, str]:
    ovr = lookup_override(company)
    out_r = str(regions or "").strip()
    out_d = str(dist or "").strip()

    if ovr:
        # Prefer curated values when current is weak or generic
        if is_weak(out_r):
            out_r = ovr[0]
        if needs_dist_upgrade(out_d, biz):
            out_d = ovr[1]
        # Also upgrade regions that are too bare (single vague token) when override richer
        if out_r.lower() in {"usa", "global", "china", "europe", "uk", "india"} and len(ovr[0]) > len(out_r) + 5:
            # keep override if more descriptive for known companies
            out_r = ovr[0]
        if ovr and not is_weak(out_d) and out_d.lower() in {"distributor", "wholesale", "wholesale / b2b distribution"}:
            out_d = ovr[1]
    else:
        if is_weak(out_r):
            out_r = infer_regions(hq, cities, company, biz)
        if needs_dist_upgrade(out_d, biz):
            out_d = infer_dist_type(biz, company)

    if is_weak(out_r):
        out_r = infer_regions(hq, cities, company, biz)
    if is_weak(out_d):
        out_d = infer_dist_type(biz, company)

    return out_r, out_d


# Formatting helpers (keep Demo Output readable)
thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
body_font = Font(name="Calibri", size=10)
wrap = Alignment(wrap_text=True, vertical="top", horizontal="left")
header_align = Alignment(wrap_text=True, vertical="center", horizontal="center")

WIDTH_HINTS = {
    "company": 28,
    "summary": 55,
    "regions served": 32,
    "distribution type": 28,
    "core categories": 34,
    "specialty focus": 30,
    "key brands": 34,
    "price segment": 22,
    "headquarters": 26,
    "cities": 24,
    "website": 30,
}


def col_width(h: str) -> float:
    low = (h or "").lower()
    for k, w in WIDTH_HINTS.items():
        if k in low:
            return float(w)
    return float(max(12, min(28, len(h or "") + 3)))


def light_format(ws, header_row: int, max_col: int) -> None:
    for c in range(1, max_col + 1):
        h = str(ws.cell(header_row, c).value or "").strip()
        ws.column_dimensions[get_column_letter(c)].width = col_width(h)
        cell = ws.cell(header_row, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin
    for r in range(header_row + 1, ws.max_row + 1):
        if not any(ws.cell(r, c).value for c in range(1, max_col + 1)):
            continue
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.font = body_font
            cell.alignment = wrap
            cell.border = thin
        ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 36, 40)
    ws.row_dimensions[header_row].height = 34
    ws.freeze_panes = ws.cell(header_row + 1, 2).coordinate
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{ws.max_row}"


def process(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=False))
    hdr_i = next(
        i
        for i, r in enumerate(rows[:12])
        if r and r[0].value and str(r[0].value).strip() == "Company"
    )
    header_row = hdr_i + 1
    header = [str(c.value or "").strip() for c in rows[hdr_i]]
    idx = {h: i for i, h in enumerate(header) if h}

    for need in ("Regions Served", "Distribution Type"):
        if need not in idx:
            raise SystemExit(f"Missing column {need} in {path}")

    updated = 0
    for r in rows[hdr_i + 1 :]:
        company = str(r[idx["Company"]].value or "").strip()
        if not company:
            continue
        biz = str(r[idx["Business Type"]].value or "").strip() if "Business Type" in idx else ""
        hq = str(r[idx["Headquarters"]].value or "").strip() if "Headquarters" in idx else ""
        cities = str(r[idx["Cities / Regions"]].value or "").strip() if "Cities / Regions" in idx else ""
        new_r, new_d = resolve(
            company,
            biz,
            hq,
            cities,
            r[idx["Regions Served"]].value,
            r[idx["Distribution Type"]].value,
        )
        changed = False
        if str(r[idx["Regions Served"]].value or "").strip() != new_r:
            r[idx["Regions Served"]].value = new_r
            changed = True
        if str(r[idx["Distribution Type"]].value or "").strip() != new_d:
            r[idx["Distribution Type"]].value = new_d
            changed = True
        if changed:
            updated += 1

    max_col = max(i for i, h in enumerate(header) if h) + 1
    light_format(ws, header_row, max_col)
    if "Run Summary" in wb.sheetnames:
        wb["Run Summary"]["A19"] = f"Regions/Distribution fill: {updated} rows"
    wb.save(path)
    wb.close()
    DEMO.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, DEMO / path.name)
    return {"market": path.parent.name, "updated": updated}


def audit() -> None:
    print(f"{'Market':<36} {'weak_reg':>8} {'weak_dist':>9}")
    for path in sorted(SRC.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        h = [str(x or "").strip() for x in data[hdr_i]]
        rows = [r for r in data[hdr_i + 1 :] if r and r[0]]
        ri, di = h.index("Regions Served"), h.index("Distribution Type")
        wr = sum(1 for r in rows if is_weak(r[ri]) or str(r[ri]).lower().startswith("not publicly"))
        wd = sum(1 for r in rows if is_weak(r[di]))
        print(f"{path.parent.name:<36} {wr:>8} {wd:>9}")
        wb.close()


def main() -> None:
    for path in sorted(SRC.rglob("*_distributors_expanded.xlsx")):
        st = process(path)
        print(f"{st['market']}: updated {st['updated']}")
    print("\nPost-fill audit:")
    audit()


if __name__ == "__main__":
    main()
