"""Restore companies wrongly dropped when pre-1900 Founded years were treated as empty.
Also re-add a few previously-dropped firms now that Founded was verified.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

spec = importlib.util.spec_from_file_location(
    "expand", r"D:\vendor-intel\scripts\expand_top10_distributors.py"
)
expand = importlib.util.module_from_spec(spec)
spec.loader.exec_module(expand)

spec2 = importlib.util.spec_from_file_location(
    "verify", r"D:\vendor-intel\scripts\verify_fill_founded_top10.py"
)
verify = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(verify)
YEAR_MAP: dict = verify.FOUNDED
HEADERS = expand.HEADERS

# Curated expand lists
EXPAND_LISTS = {
    "muslin_products_market_global": expand.MUSLIN,
    "elderberry_supplements_market_global": expand.ELDERBERRY,
    "baijiu_market_global": expand.BAIJIU,
    "shunt_resistor_market_global": expand.SHUNT,
    "mountain_bike_suspension_market_global": expand.MTB,
    "nail_polish_market_global": expand.NAIL,
    "sheet_face_masks_market_global": expand.MASKS,
    "axillary_crutches_market_global": expand.CRUTCHES,
    "benign_positional_vertigo_market_global": expand.BPV,
    "instant_noodles_market_global": expand.NOODLES,
}

EXTRA_YEARS = {
    "McKesson": "1833",
    "McKesson Medical-Surgical": "1833",
    "McKesson (Consumer Health)": "1833",
    "Owens & Minor": "1882",
    "Bunzl Healthcare / related MRO medical": "1854",
    "Paul Hartmann AG distribution": "1818",
    "Mölnlycke Health Care distribution": "1849",
    "Shiwandistillery": "1830",
    "Young's Market Company": "1888",
    "Gordon Food Service": "1897",
    "McLane Company": "1894",
    "Performance Food Group": "2005",
    "Performance Food Group (PFG)": "2005",
    "Aeon Group wholesale/retail": "1758",
    "Nestlé Maggi distribution": "1866",
    "Bilstein": "1873",
    "Koni": "1857",
    "Pon.Bike distribution (Cannondale, Santa Cruz parent ops)": "2011",
    "ROC-LON Fabrics (Rockland Industries)": "1832",
    "Whaleys (Bradford) Ltd": "1873",
    "Liberty Fabrics (Wholesale)": "1875",
    "Prym Consumer USA / Haberdashery Dist.": "1530",
    "Gutermann Distribution Partners": "1864",
    "Coats Group Trade Distribution": "1755",
    "Douglas (Europe)": "1821",
    "Douglas Beauty Europe": "1821",
    "Watsons / AS Watson beauty retail": "1828",
    "Boots UK beauty": "1849",
    "Sager Electronics": "1954",
    # newly verified fills (re-add)
    "Bharucha Associates": "1994",
    "DM Distributors": "1968",
    "Doray Kumas": "2019",
    "Anand Agencies (Chennai)": "1961",
    "Wholesale Nail Supply": "1991",
}

# Skip restoring directory-style non-distributors
SKIP_RESTORE = {"ThomasNet Industrial Textile Distributors"}

WRONG = {
    "axillary_crutches_market_global": [
        "McKesson",
        "McKesson Medical-Surgical",
        "Owens & Minor",
        "Bunzl Healthcare / related MRO medical",
        "Paul Hartmann AG distribution",
        "Mölnlycke Health Care distribution",
    ],
    "baijiu_market_global": ["Shiwandistillery", "Young's Market Company"],
    "benign_positional_vertigo_market_global": [
        "McKesson",
        "Owens & Minor",
        "McKesson Medical-Surgical",
    ],
    "elderberry_supplements_market_global": ["McKesson (Consumer Health)"],
    "instant_noodles_market_global": [
        "Gordon Food Service",
        "McLane Company",
        "Performance Food Group",
        "Performance Food Group (PFG)",
        "Aeon Group wholesale/retail",
        "Nestlé Maggi distribution",
    ],
    "mountain_bike_suspension_market_global": [
        "Bilstein",
        "Koni",
        "Pon.Bike distribution (Cannondale, Santa Cruz parent ops)",
    ],
    "muslin_products_market_global": [
        "ROC-LON Fabrics (Rockland Industries)",
        "Whaleys (Bradford) Ltd",
        "Liberty Fabrics (Wholesale)",
        "Prym Consumer USA / Haberdashery Dist.",
        "Gutermann Distribution Partners",
        "Coats Group Trade Distribution",
    ],
    "nail_polish_market_global": ["Douglas (Europe)"],
    "sheet_face_masks_market_global": [
        "Watsons / AS Watson beauty retail",
        "Boots UK beauty",
        "Douglas Beauty Europe",
    ],
    "shunt_resistor_market_global": ["Sager Electronics"],
}

# Additional re-adds with verified Founded
READD = {
    "axillary_crutches_market_global": ["Bharucha Associates"],
    "muslin_products_market_global": ["DM Distributors", "Doray Kumas"],
    "sheet_face_masks_market_global": ["Anand Agencies (Chennai)"],
    "nail_polish_market_global": ["Wholesale Nail Supply"],
}

KNOWN_HQ = {
    "McKesson": "Irving, TX, USA",
    "McKesson Medical-Surgical": "USA",
    "McKesson (Consumer Health)": "Irving, TX, USA",
    "Owens & Minor": "USA",
    "ROC-LON Fabrics (Rockland Industries)": "Baltimore, MD, USA (Rockland Industries)",
    "Bharucha Associates": "India",
    "DM Distributors": "Biddeford, ME, USA",
    "Doray Kumas": "Turkey",
    "Anand Agencies (Chennai)": "Chennai, Tamil Nadu, India",
    "Wholesale Nail Supply": "USA",
}


def year_for(company: str) -> str:
    if company in EXTRA_YEARS:
        return EXTRA_YEARS[company]
    low = company.lower()
    for k, v in YEAR_MAP.items():
        if k in low:
            return str(v)
    return ""


def find_expand_row(market: str, company: str) -> dict | None:
    for r in EXPAND_LISTS.get(market, []):
        if r.get("Company") == company:
            return dict(r)
    # fuzzy: ROC-LON may be under different name in expand
    for lst in EXPAND_LISTS.values():
        for r in lst:
            if r.get("Company") == company:
                return dict(r)
    return None


def make_row(market: str, company: str) -> dict:
    row = find_expand_row(market, company)
    if not row:
        row = {h: "" for h in HEADERS}
        row["Company"] = company
        row["Business Type"] = "Distributor"
        row["Ownership"] = "Independent / privately held (public filings not always available)"
        row["Distribution Type"] = "Wholesale / B2B distribution"
        row["Data Confidence"] = "Medium"
        row["Quality Score"] = "0.75"
        row["Website"] = ""
        row["Summary"] = f"{company} restored after Founded-year filter fix."
        row["Data Sources"] = "Restored from expand curated / prior fill maps"
    year = year_for(company) or str(row.get("Founded") or "").strip()
    if not year or year.lower().startswith("not"):
        year = year_for(company)
    row["Founded"] = year
    if not str(row.get("Headquarters") or "").strip() or str(row["Headquarters"]).lower().startswith(
        "not"
    ):
        row["Headquarters"] = KNOWN_HQ.get(company, row.get("Headquarters") or "")
    if not str(row.get("Cities / Regions") or "").strip() or str(row["Cities / Regions"]).lower().startswith(
        "not"
    ):
        row["Cities / Regions"] = row.get("Cities / Regions") or row.get("Headquarters") or ""
    # clear weak placeholders on key fields if still empty
    for h in HEADERS:
        if not str(row.get(h) or "").strip():
            if h in ("Contact Person", "Role", "Email", "Phone / WhatsApp", "LinkedIn"):
                row[h] = "Not confirmed publicly"
            elif h not in ("Founded", "Headquarters", "Cities / Regions", "Company"):
                row[h] = row.get(h) or "Not confirmed publicly"
    src = str(row.get("Data Sources") or "")
    note = f"Restored; Founded={row['Founded']}"
    row["Data Sources"] = (src + "; " if src else "") + note
    return row


def process(market: str, path: Path) -> tuple[int, list[str]]:
    names = [n for n in WRONG.get(market, []) if n not in SKIP_RESTORE]
    names += READD.get(market, [])
    if not names:
        return 0, []

    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=False))
    hdr_i = next(
        i
        for i, r in enumerate(rows[:12])
        if r and r[0].value and str(r[0].value).strip() == "Company"
    )
    header = [str(c.value or "").strip() for c in rows[hdr_i]]
    existing = {
        str(r[0].value).strip()
        for r in rows[hdr_i + 1 :]
        if r and r[0].value
    }

    added = []
    first_data = hdr_i + 2
    # append after last data row
    next_row = ws.max_row + 1
    for company in names:
        if company in existing:
            continue
        row = make_row(market, company)
        if not row.get("Founded"):
            print(f"  SKIP no year: {company}")
            continue
        for col, h in enumerate(header, 1):
            ws.cell(next_row, col).value = row.get(h, "")
        next_row += 1
        added.append(f"{company}={row['Founded']}")

    if "Run Summary" in wb.sheetnames:
        ws2 = wb["Run Summary"]
        ws2["A11"] = f"Restored after year-filter bug: {', '.join(added)}"

    wb.save(path)
    wb.close()
    return len(added), added


def main() -> None:
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        market = path.parent.name
        n, added = process(market, path)
        print(f"{market}: restored {n}")
        for a in added:
            print(f"  + {a}")

    # recount
    print("\nFinal counts:")
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        fi = [str(h or "").strip() for h in data[hdr_i]].index("Founded")

        def weak(v):
            if v is None or not str(v).strip():
                return True
            s = str(v).strip().lower()
            if s.startswith("not listed") or s.startswith("n/a"):
                return True
            import re

            return not re.search(r"\b(1[5-9]\d{2}|20[0-2]\d)\b", str(v))

        rows = [r for r in data[hdr_i + 1 :] if r and r[0]]
        weak_n = sum(1 for r in rows if weak(r[fi]))
        print(f"  {path.parent.name}: companies={len(rows)} weak_founded={weak_n}")
        wb.close()


if __name__ == "__main__":
    main()
