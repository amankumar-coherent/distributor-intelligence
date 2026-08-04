"""Fill Founded where researched; remove companies still without a real year."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

# High-confidence fills from website / LinkedIn / registry / press
FILLS: dict[str, tuple[str, str]] = {
    # company name -> (founded, source note appended to Data Sources)
    "Dearsundays": ("2017", "LinkedIn/dear-sundays; press Amy Ling Lin founded 2017"),
    "DrinkUpNY": ("2006", "FoodBev/press: DrinkUpNY.com launched end of 2006"),
    "Prosemic": ("2012", "prosemicsr.com/development-history; LinkedIn entity 2019"),
    "Casa Fashions (Karur)": ("2016", "casafashions.com: created in 2016"),
    "SyFabrics": ("2003", "LA business tax start 2003-05-16 (Sytradingpost Inc dba SyFabrics)"),
    "Beauty Manufacturing Solutions (BMS Nail)": (
        "1922",
        "LinkedIn beauty-manufacturing-solutions-corp founded 1922",
    ),
    "Doctor Althea Global": ("2014", "doctoraltheaglobal.com/about; DermApproved founded 2014"),
    "Sewing Machine Plus": ("1974", "sewingmachinesplus.com: Family-Owned & Trusted Since 1974"),
    "Heartland Elderberry Farms": (
        "2020",
        "heartlandelderberryfarms.com: Carter farm ops began early 2020",
    ),
    "Elder Farms": (
        "1899",
        "elderfarms.com About: Buehler family farm origin 1899 (Mount Vernon, MO)",
    ),
    "The Nailhub / pro e-wholesalers": (
        "2014",
        "LinkedIn owner The Nail Hub from Dec 2014; thenailhub.com",
    ),
    "Central India Fabrics": (
        "2009",
        "centralindiafabrics.net/profile: Established 2009 Indore",
    ),
    "Hethstia": ("2012", "hethstia.com/about: Hethstia Company born 2012"),
    "Kaganails": (
        "1996",
        "Guangzhou Kaga Cosmetics est. 1996 (UK distributor of Kaga line)",
    ),
}


def is_weak_founded(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    if not s:
        return True
    low = s.lower()
    if low in {"n/a", "na", "none", "-", "—", "null", "unknown"}:
        return True
    if low.startswith("not listed") or "not listed on" in low or low.startswith("n/a —"):
        return True
    # already a year? (allow 1500–2025)
    if re.fullmatch(r"1[5-9]\d{2}|20[0-2]\d", s):
        return False
    if re.search(r"\b(1[5-9]\d{2}|20[0-2]\d)\b", s) and not low.startswith("not"):
        return False
    return True


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
    fi = header.index("Founded")
    si = header.index("Data Sources") if "Data Sources" in header else None

    keep: list[tuple] = []
    filled: list[str] = []
    removed: list[str] = []

    for r in rows[hdr_i + 1 :]:
        company = str(r[ci].value or "").strip()
        if not company:
            continue
        vals = [c.value for c in r[: len(header)]]
        founded = vals[fi]
        if is_weak_founded(founded):
            if company in FILLS:
                year, note = FILLS[company]
                vals[fi] = year
                if si is not None:
                    prev = str(vals[si] or "").strip()
                    vals[si] = (prev + "; " if prev else "") + f"Founded fill: {note}"
                filled.append(f"{company}={year}")
                keep.append(tuple(vals))
            else:
                removed.append(company)
        else:
            keep.append(tuple(vals))

    # rewrite data rows preserving title + blank + header
    header_row = hdr_i + 1
    first_data = hdr_i + 2
    while ws.max_row >= first_data:
        ws.delete_rows(first_data)
    for i, vals in enumerate(keep):
        for col, val in enumerate(vals, 1):
            ws.cell(first_data + i, col).value = val

    if "Run Summary" in wb.sheetnames:
        ws2 = wb["Run Summary"]
        ws2["A8"] = f"Founded fill pass: filled {len(filled)}, removed {len(removed)}, remaining {len(keep)}"
        if filled:
            ws2["A9"] = "Filled: " + "; ".join(filled)
        if removed:
            ws2["A10"] = "Removed (no Founded): " + "; ".join(removed)

    wb.save(path)
    wb.close()
    return {"filled": filled, "removed": removed, "remaining": len(keep)}


def main() -> None:
    print(f"{'Market':<42} {'N':>4}  filled  removed")
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        stats = process(path)
        name = path.parent.name
        print(
            f"{name:<42} {stats['remaining']:>4}  {len(stats['filled']):>6}  {len(stats['removed']):>7}"
        )
        for x in stats["filled"]:
            print(f"    FILL {x}")
        for x in stats["removed"]:
            print(f"    DROP {x}")


if __name__ == "__main__":
    main()
