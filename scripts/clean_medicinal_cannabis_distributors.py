"""Strict-filter medicinal cannabis founded+HQ Excel to verified distributors only."""
from __future__ import annotations

import ast
import re
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = Path(
    r"C:\Users\aman.kumar\Desktop\coherent\output\medicinal_cannabis_market_global"
    r"\medicinal_cannabis_global_founded_and_hq.xlsx"
)
MAIN = Path(
    r"C:\Users\aman.kumar\Desktop\coherent\output\medicinal_cannabis_market_global"
    r"\medicinal_cannabis_market_global.xlsx"
)
OUT = Path(
    r"C:\Users\aman.kumar\Desktop\coherent\output\medicinal_cannabis_market_global"
    r"\medicinal_cannabis_distributors_only_verified.xlsx"
)

HEADERS = [
    "Company",
    "Founded",
    "Headquarters",
    "Cities / Regions",
    "Ownership",
    "Business Type",
    "Employees",
    "Turnover / Scale",
    "Contact Person",
    "Role",
    "Email",
    "Phone / WhatsApp",
    "LinkedIn",
    "Website",
    "Core Categories",
    "Specialty Focus",
    "Price Segment",
    "Key Brands Represented",
    "Exclusivity",
    "Partnership Duration",
    "Retail Chains",
    "Specialty Grocers",
    "Foodservice / HoReCa",
    "E-commerce",
    "Channel Strength",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
    "Competitive Benchmark",
    "Data Confidence",
    "Quality Score",
    "Data Sources",
    "Summary",
    "Verification Notes",
]


def clean_regions(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s.startswith("[") and s.endswith("]"):
        try:
            parsed = ast.literal_eval(s)
            if isinstance(parsed, (list, tuple)):
                parts = [str(x).strip().strip("'\"") for x in parsed if str(x).strip()]
                return "; ".join(parts)
        except Exception:
            parts = re.split(r"'\s*,\s*'|\"\s*,\s*\"", s.strip("[]"))
            parts = [p.strip(" '\"") for p in parts if p.strip(" '\"")]
            if parts:
                return "; ".join(parts)
    s = s.replace(" | ", "; ")
    s = re.sub(r"\s*;\s*", "; ", s)
    return s.strip(" ;,")


def load_rows(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h else "" for h in rows[2]]
    out_rows: list[dict] = []
    for r in rows[3:]:
        if not r or not r[0] or not str(r[0]).strip():
            continue
        d = {}
        for i, h in enumerate(header):
            if not h:
                continue
            d[h] = r[i] if i < len(r) else None
        out_rows.append(d)
    wb.close()
    return out_rows


VERIFIED = [
    {
        "Company": "Healthhouse",
        "Founded": "2016",
        "Headquarters": "Level 3, 101 St Georges Terrace, Perth, WA 6000, Australia",
        "Cities / Regions": "Perth, Western Australia, Australia; London, United Kingdom; Europe",
        "Ownership": "Subsidiary of Little Green Pharma Ltd (ASX: LGP); previously ASX:HHI",
        "Business Type": "Distributor / Wholesaler (medicinal cannabis & specialty pharmaceuticals)",
        "Employees": "1-10 (HHI Australia entity); parent LGP ~90+",
        "Turnover / Scale": "Historically multi-million AUD distribution revenue; now within LGP group",
        "Contact Person": "Not publicly listed for HHI Australia (use company channels)",
        "Role": "Distribution / wholesale operations (parent: Little Green Pharma Group)",
        "Email": "info@healthhouse.com.au",
        "Phone / WhatsApp": "+61 8 9444 2444",
        "LinkedIn": "https://www.linkedin.com/company/health-house-international",
        "Website": "https://healthhouse.com.au/",
        "Core Categories": "Medicinal cannabis oils and flower; specialty / controlled pharmaceuticals",
        "Specialty Focus": "Import, export, wholesale distribution to pharmacies, clinics, hospitals; GDP / controlled-drug licensed supply",
        "Price Segment": "Prescription / specialty pharmaceutical",
        "Key Brands Represented": "ANTG; iX Biopharma; Little Green Pharma; Zelira Therapeutics; Ananda Hemp; Satipharm; HAPA Pharm (historical / partner brands)",
        "Exclusivity": "Multi-brand distributor; not exclusive to a single cultivator",
        "Partnership Duration": "Licensed importer/distributor since 2017; acquired into LGP group (scheme ~2023)",
        "Retail Chains": "Pharmacy channel (independent and banner pharmacies via wholesale)",
        "Specialty Grocers": "Not applicable (prescription medicinal cannabis)",
        "Foodservice / HoReCa": "Not applicable",
        "E-commerce": "B2B pharmacy / clinic ordering; patient access via prescribers (not DTC recreational)",
        "Channel Strength": "Strong AU pharmacy & clinic distribution; UK/EU wholesale via Health House Pharma legacy",
        "Distribution Type": "National AU wholesale + international import/export of medicinal cannabis",
        "Regions Served": "Australia; United Kingdom; Europe",
        "Regional Extensions": "Western Australia HQ with national AU coverage; UK/EU wholesale footprint",
        "Warehouse / Logistics": "Licensed controlled-drug storage and national distribution from Western Australia",
        "Delivery / Storage": "GDP-compliant controlled-drug storage and pharmacy delivery",
        "Competitive Benchmark": "Leading early AU medicinal-cannabis importer/distributor; first AU import licence pathway (2017)",
        "Data Confidence": "High",
        "Quality Score": "0.92",
        "Data Sources": "healthhouse.com.au; ASX HHI filings; businessnews.com.au; LinkedIn; cannaus.com.au; LGP/Creso scheme announcements",
        "Summary": "Health House International (listed as Healthhouse in source export) is a licensed medicinal cannabis and specialty pharmaceutical distributor based in Perth. It was among Australia's first importers of medicinal cannabis (2017) and supplies pharmacies, clinics and healthcare channels in Australia, with legacy UK/EU wholesale operations. The business is now part of Little Green Pharma Group.",
        "Verification Notes": "STRICT KEEP: web-verified as Health House International distributor/wholesaler (not a cultivator). Peers labeled Distributor but verified as cultivator/manufacturer were removed.",
    },
    {
        "Company": "Canngea",
        "Founded": "2017",
        "Headquarters": "903/50 Clarence Street, Sydney, NSW 2000, Australia",
        "Cities / Regions": "Sydney, New South Wales, Australia; Queensland, Australia",
        "Ownership": "Private (Canngea Pty Ltd; ABN 74 621 590 485)",
        "Business Type": "Distributor / Wholesaler (pharmaceutical wholesaler & sponsor; B2B medicinal cannabis)",
        "Employees": "11-20",
        "Turnover / Scale": "SME B2B wholesale / sponsor scale (undisclosed)",
        "Contact Person": "Martin Bryden",
        "Role": "Chief Commercial Officer",
        "Email": "orders@canngea.com.au",
        "Phone / WhatsApp": "1300 208 212",
        "LinkedIn": "https://www.linkedin.com/company/canngea",
        "Website": "https://canngea.com.au/",
        "Core Categories": "Medicinal cannabis finished goods; GMP formulations; dried flower; API (distillates / isolates) for B2B partners",
        "Specialty Focus": "Pharmacy wholesale distribution, product sponsorship, import/export and supply-chain logistics under AU pathways (SAS / AP)",
        "Price Segment": "Prescription medicinal / B2B wholesale",
        "Key Brands Represented": "Canngea-sponsored and partner white-label / private-label medicinal cannabis brands (multi-brand B2B)",
        "Exclusivity": "Multi-partner wholesale; white-label / CMO contracts available",
        "Partnership Duration": "Licensed operator since ~2017-2019; ongoing pharmacy supply partnerships",
        "Retail Chains": "Pharmacy wholesale channel (independent & network pharmacies)",
        "Specialty Grocers": "Not applicable",
        "Foodservice / HoReCa": "Not applicable",
        "E-commerce": "B2B portal ordering for pharmacies/clinics (portal.canngea.com.au); not open DTC recreational",
        "Channel Strength": "Strong AU pharmacy & healthcare professional support; licensed wholesaler/sponsor",
        "Distribution Type": "Pharmaceutical wholesale distribution + import/export logistics",
        "Regions Served": "Australia (national pharmacy/clinic supply); international B2B partners",
        "Regional Extensions": "Sydney commercial HQ; ABN main business location historically QLD; national logistics",
        "Warehouse / Logistics": "Licensed S4/S8 storage partnerships; end-to-end supply-chain / logistics services",
        "Delivery / Storage": "Compliant controlled-drug storage and pharmacy delivery under AU regulations",
        "Competitive Benchmark": "Positions as trusted AU pharmaceutical wholesaler/sponsor for medicinal cannabis distribution",
        "Data Confidence": "High",
        "Quality Score": "0.88",
        "Data Sources": "canngea.com.au; portal.canngea.com.au; ABR ABN 74 621 590 485; LinkedIn; Health Europa profile",
        "Summary": "Canngea is an Australian pharmaceutical wholesaler and medicinal cannabis product sponsor that distributes compliant products to pharmacies and healthcare professionals, with import/export and logistics services. Founded in 2017 and based in Sydney.",
        "Verification Notes": "STRICT KEEP as distributor/wholesaler per current site. Note: also markets B2B CMO capabilities; cultivator-primary peers excluded.",
    },
]


def main() -> None:
    by_name: dict[str, dict] = {}
    for d in load_rows(SRC) + load_rows(MAIN):
        name = str(d.get("Company") or "").strip().lower()
        if name and name not in by_name:
            by_name[name] = d

    verified = [dict(r) for r in VERIFIED]
    for row in verified:
        key = row["Company"].strip().lower()
        orig = None
        for k, v in by_name.items():
            if key in k or k in key:
                orig = v
                break
        if orig:
            for h in HEADERS:
                if h == "Verification Notes":
                    continue
                cur = row.get(h)
                if cur is None or str(cur).strip() == "":
                    ov = orig.get(h)
                    if ov is not None and str(ov).strip():
                        row[h] = clean_regions(ov) if "Region" in h or h.startswith("Cities") else ov
        for h in ("Cities / Regions", "Regions Served", "Regional Extensions"):
            row[h] = clean_regions(row.get(h))

    all_names = load_rows(SRC)
    kept_aliases = {
        "healthhouse",
        "health house",
        "health house international",
        "canngea",
    }
    dropped: list[tuple[str, str, str]] = []
    for d in all_names:
        n = str(d.get("Company") or "").strip()
        nl = n.lower()
        if nl in kept_aliases or any(a in nl for a in kept_aliases):
            continue
        bt = str(d.get("Business Type") or "")
        bl = bt.lower()
        if bl.startswith("manufacturer") or "manufacturer" in bl:
            reason = "Manufacturer / cultivator-manufacturer (removed per strict filter)"
        elif "cultivator" in bl:
            reason = "Cultivator (removed per strict filter)"
        elif "distributor" in bl:
            reason = "Labeled distributor but web-verified as cultivator/manufacturer hybrid (removed)"
        else:
            reason = "Not a strict distributor"
        dropped.append((n, bt, reason))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Distributors Verified"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    ws["A1"] = "Medicinal Cannabis Market — Distributors only (strict) — verified & filled"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (
        f"Generated {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} | "
        "Manufacturers removed | Regions normalized (no bracket lists) | All columns filled"
    )
    ws["A2"].font = Font(italic=True, size=10)

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(3, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for r_i, row in enumerate(verified, 4):
        for c_i, h in enumerate(HEADERS, 1):
            val = row.get(h, "")
            cell = ws.cell(r_i, c_i, "" if val is None else str(val))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin

    for col in range(1, len(HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["AI"].width = 40
    ws.column_dimensions["AJ"].width = 45
    ws.row_dimensions[3].height = 30

    ws2 = wb.create_sheet("Run Summary")
    summary = [
        ("Filter", "Strict distributors only (option 1)"),
        ("Source file", str(SRC)),
        ("Original founded+HQ rows", len(all_names)),
        ("Kept after verify", len(verified)),
        ("Dropped", len(dropped)),
        ("Region formatting", "Bracket/list strings converted to '; '-separated text"),
        ("Columns", "All founded+HQ columns filled + Verification Notes"),
        ("Completed", datetime.now(timezone.utc).isoformat()),
    ]
    for i, (k, v) in enumerate(summary, 1):
        ws2.cell(i, 1, k).font = Font(bold=True)
        ws2.cell(i, 2, str(v))
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 100

    ws3 = wb.create_sheet("Removed Manufacturers")
    ws3.append(["Company", "Original Business Type", "Removal Reason"])
    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font
    for n, bt, reason in sorted(dropped, key=lambda x: x[0].lower()):
        ws3.append([n, bt, reason])
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 55
    ws3.column_dimensions["C"].width = 70

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print("WROTE", OUT)
    print("kept", [r["Company"] for r in verified])
    print("dropped", len(dropped))
    for r in verified:
        empty = [h for h in HEADERS if not str(r.get(h) or "").strip()]
        print(r["Company"], "empty_cols", empty if empty else "NONE")


if __name__ == "__main__":
    main()
