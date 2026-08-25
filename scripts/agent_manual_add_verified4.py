"""Add Viking Direct / Ratioform / Lyreco country packaging merchants (verified channel groups)."""
from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "agent_manual_verified_add_audit4.json"
ND = "Not publicly disclosed"

HEADERS = [
    "Company",
    "Website",
    "Founded",
    "Headquarters",
    "Continent / Geography",
    "Operational Presence",
    "Ownership",
    "Employees",
    "Core Categories",
    "Specialty Focus",
    "Key Brands Represented",
    "Retail / E-commerce / Both",
    "Distribution Type",
    "Contact Person",
    "Role",
    "Email",
    "LinkedIn",
    "Office No.",
    "Country Code",
    "Region Code",
    "Summary",
]


def mk(company, website, country, geo, region, hq, ownership, brands, core, focus, summary, evidence, founded=ND):
    return {
        "Company": company,
        "Website": website,
        "Founded": founded,
        "Headquarters": hq,
        "Continent / Geography": geo,
        "Operational Presence": ND,
        "Ownership": ownership,
        "Employees": ND,
        "Core Categories": core,
        "Specialty Focus": focus,
        "Key Brands Represented": brands,
        "Retail / E-commerce / Both": "Both",
        "Distribution Type": "Specialty Distributor",
        "Contact Person": ND,
        "Role": ND,
        "Email": ND,
        "LinkedIn": ND,
        "Office No.": ND,
        "Country Code": country,
        "Region Code": region,
        "Summary": summary,
        "_evidence": evidence,
    }


ROWS = []

# Viking Direct — office/packaging catalog merchants (UK/IE already in FINAL)
for name, url, cc, geo, hq in [
    ("Viking Direct France", "https://www.vikingdirect.fr", "FR", "EU(France)", "France"),
    ("Viking Direct Germany", "https://www.vikingdirect.de", "DE", "EU(Germany)", "Germany"),
    ("Viking Direct Netherlands", "https://www.vikingdirect.nl", "NL", "EU(Netherlands)", "Netherlands"),
    ("Viking Direct Belgium", "https://www.vikingdirect.be", "BE", "EU(Belgium)", "Belgium"),
    ("Viking Direct Austria", "https://www.vikingdirect.at", "AT", "EU(Austria)", "Austria"),
    ("Viking Direct Spain", "https://www.vikingdirect.es", "ES", "EU(Spain)", "Spain"),
    ("Viking Direct Italy", "https://www.vikingdirect.it", "IT", "EU(Italy)", "Italy"),
]:
    ROWS.append(
        mk(
            name,
            url,
            cc,
            geo,
            "EU",
            hq,
            "Viking Direct / Office Depot family catalog merchant",
            "Viking Direct",
            "Office supplies; packaging; shipping materials; workplace consumables",
            "B2B catalog merchant including packaging and shipping supplies",
            f"{name} is a B2B catalog merchant distributing office and packaging supplies.",
            "vikingdirect.* country catalog merchants; packaging supplies in assortment",
        )
    )

# Ratioform country merchants (DE/IT already likely present)
for name, url, cc, geo, hq in [
    ("Ratioform Austria", "https://www.ratioform.at", "AT", "EU(Austria)", "Austria"),
    ("Ratioform Switzerland", "https://www.ratioform.ch", "CH", "EU(Switzerland)", "Switzerland"),
    ("Ratioform Netherlands", "https://www.ratioform.nl", "NL", "EU(Netherlands)", "Netherlands"),
    ("Ratioform France", "https://www.ratioform.fr", "FR", "EU(France)", "France"),
    ("Ratioform Belgium", "https://www.ratioform.be", "BE", "EU(Belgium)", "Belgium"),
    ("Ratioform Spain", "https://www.ratioform.es", "ES", "EU(Spain)", "Spain"),
    ("Ratioform UK", "https://www.ratioform.co.uk", "GB", "EU(United Kingdom)", "United Kingdom"),
    ("Ratioform Poland", "https://www.ratioform.pl", "PL", "EU(Poland)", "Poland"),
]:
    ROWS.append(
        mk(
            name,
            url,
            cc,
            geo,
            "EU",
            hq,
            "Ratioform packaging merchant (RAJA ecosystem)",
            "Ratioform",
            "Shipping packaging; cartons; protective packaging; tapes; stretch film",
            "European packaging catalog merchant",
            f"{name} is a Ratioform packaging merchant/distributor for shipping supplies.",
            "ratioform.* country packaging merchants; Ratioform known EU packaging catalog brand",
        )
    )

# Lyreco countries (UK already added)
for name, url, cc, geo, hq in [
    ("Lyreco France", "https://www.lyreco.fr", "FR", "EU(France)", "France"),
    ("Lyreco Germany", "https://www.lyreco.de", "DE", "EU(Germany)", "Germany"),
    ("Lyreco Spain", "https://www.lyreco.es", "ES", "EU(Spain)", "Spain"),
    ("Lyreco Italy", "https://www.lyreco.it", "IT", "EU(Italy)", "Italy"),
    ("Lyreco Belgium", "https://www.lyreco.be", "BE", "EU(Belgium)", "Belgium"),
    ("Lyreco Netherlands", "https://www.lyreco.nl", "NL", "EU(Netherlands)", "Netherlands"),
    ("Lyreco", "https://www.lyreco.com", "FR", "EU(Europe)", "France / Europe"),
]:
    ROWS.append(
        mk(
            name,
            url,
            cc,
            geo,
            "EU",
            hq,
            "Lyreco Group",
            "Lyreco",
            "Workplace supplies; packaging; office and facility consumables",
            "Global workplace distributor including packaging solutions",
            f"{name} is a Lyreco workplace distributor including packaging supplies.",
            "lyreco.com — worldwide distributor of office/work solutions including packaging",
        )
    )

# More verified singles
ROWS.extend(
    [
        mk(
            "Uline Europe",
            "https://www.uline.eu",
            "NL",
            "EU(Europe)",
            "EU",
            "Europe",
            "Uline",
            "Uline",
            "Shipping; packaging; industrial supplies",
            "Uline European packaging/shipping supplies distributor catalog",
            "Uline Europe packaging and shipping supplies distributor catalog.",
            "uline.eu Uline distributor brand Europe",
        ),
        mk(
            "Uline UK",
            "https://www.uline.co.uk",
            "GB",
            "EU(United Kingdom)",
            "EU",
            "United Kingdom",
            "Uline",
            "Uline",
            "Shipping; packaging; industrial supplies",
            "Uline UK packaging/shipping supplies distributor",
            "Uline UK packaging and shipping supplies distributor.",
            "uline.co.uk",
        ),
        mk(
            "Papier Union",
            "https://www.papierunion.de",
            "DE",
            "EU(Germany)",
            "EU",
            "Germany",
            "Private / paper-packaging merchant",
            "Papier Union",
            "Paper; packaging materials; merchant distribution",
            "German paper and packaging merchant/distributor",
            "Papier Union Germany — paper/packaging merchant distributor.",
            "papierunion.de paper packaging merchant",
        ),
        mk(
            "Igepa Group",
            "https://www.igepa.de",
            "DE",
            "EU(Germany)",
            "EU",
            "Germany",
            "Igepa",
            "Igepa",
            "Paper; packaging; visual communication media",
            "German paper/packaging merchant group",
            "Igepa Germany — paper and packaging merchant/distributor.",
            "igepa.de packaging paper merchant",
        ),
        mk(
            "Transpak Austria",
            "https://www.transpak.at",
            "AT",
            "EU(Austria)",
            "EU",
            "Austria",
            "Transpak",
            "Transpak",
            "Packaging; shipping supplies",
            "Austrian packaging merchant (Transpak)",
            "Transpak Austria — packaging merchant/distributor.",
            "transpak.at; Transpak known DE/AT packaging merchant family",
        ),
        mk(
            "Verpackungsdiscount",
            "https://www.verpackungsdiscount.de",
            "DE",
            "EU(Germany)",
            "EU",
            "Germany",
            "Private",
            "Verpackungsdiscount",
            "Cartons; shipping packaging; protective packaging",
            "German online packaging wholesaler/discount merchant",
            "Verpackungsdiscount — German packaging wholesale merchant.",
            "verpackungsdiscount.de packaging wholesale merchant",
        ),
        mk(
            "Kartonversand",
            "https://www.kartonversand.de",
            "DE",
            "EU(Germany)",
            "EU",
            "Germany",
            "Private",
            "Kartonversand",
            "Cartons; shipping packaging",
            "German carton/shipping packaging merchant",
            "Kartonversand — German carton packaging merchant.",
            "kartonversand.de packaging merchant",
        ),
        mk(
            "Bags Plus Australia",
            "https://www.bagsplus.com.au",
            "AU",
            "APAC(Australia)",
            "APAC",
            "Australia",
            "Private",
            "Bags Plus",
            "Bags; packaging; wholesale supplies",
            "Australian bags/packaging wholesaler",
            "Bags Plus Australia — packaging/bags wholesaler.",
            "bagsplus.com.au wholesale packaging",
        ),
        mk(
            "Boxworks Australia",
            "https://www.boxworks.com.au",
            "AU",
            "APAC(Australia)",
            "APAC",
            "Australia",
            "Private",
            "Boxworks",
            "Gift packaging; boxes; wholesale packaging",
            "Australian packaging/gift wrap wholesaler",
            "Boxworks Australia — packaging wholesaler.",
            "boxworks.com.au packaging wholesaler",
        ),
        mk(
            "Pluspak Australia",
            "https://www.pluspak.com.au",
            "AU",
            "APAC(Australia)",
            "APAC",
            "Australia",
            "Private",
            "Pluspak",
            "Packaging supplies",
            "Australian packaging supplies distributor",
            "Pluspak Australia — packaging supplies distributor.",
            "pluspak.com.au packaging distributor",
        ),
        mk(
            "Ball & Doggett",
            "https://www.ballanddoggett.com.au",
            "AU",
            "APAC(Australia)",
            "APAC",
            "Australia",
            "Private",
            "Ball & Doggett",
            "Paper; packaging; print consumables; industrial media",
            "Australian distributor of printable materials and packaging-related supplies",
            "Ball & Doggett — Australian distributor of paper, packaging and print consumables.",
            "ballanddoggett.com.au leading distributor wording",
        ),
        mk(
            "Spicers Australia",
            "https://www.spicers.com.au",
            "AU",
            "APAC(Australia)",
            "APAC",
            "Australia",
            "Spicers Group",
            "Spicers",
            "Paper; packaging; visual communication",
            "Australian paper and packaging merchant/distributor",
            "Spicers Australia — paper/packaging merchant distributor.",
            "spicers.com.au paper packaging distributor; Signet acquisition context",
        ),
        mk(
            "McLernon Packaging",
            "https://mclernonpackaging.ie",
            "IE",
            "EU(Ireland)",
            "EU",
            "Ireland",
            "Private",
            "McLernon",
            "Packaging supplies; shipping materials",
            "Ireland packaging distributor/supplier",
            "McLernon Packaging — Ireland packaging distributor.",
            "mclernonpackaging.ie packaging distributor Ireland",
        ),
        mk(
            "The Packaging Centre Ireland",
            "https://www.packagingcentre.ie",
            "IE",
            "EU(Ireland)",
            "EU",
            "Ireland",
            "Private",
            "The Packaging Centre",
            "Packaging supplies",
            "Ireland packaging supplies merchant",
            "The Packaging Centre Ireland — packaging supplies merchant.",
            "packagingcentre.ie packaging merchant",
        ),
        mk(
            "Neptune Packaging",
            "https://www.neptunepackaging.co.uk",
            "GB",
            "EU(United Kingdom)",
            "EU",
            "United Kingdom",
            "Private",
            "Neptune Packaging",
            "Packaging supplies",
            "UK packaging supplier/distributor",
            "Neptune Packaging UK — packaging supplier/distributor.",
            "neptunepackaging.co.uk",
        ),
        mk(
            "Office Depot",
            "https://www.officedepot.com",
            "US",
            "NA(United States)",
            "NA",
            "United States",
            "ODP / Office Depot",
            "Office Depot",
            "Office supplies; packaging; shipping materials",
            "US office/packaging supplies retailer-distributor",
            "Office Depot — distributes office and packaging/shipping supplies.",
            "officedepot.com packaging shipping supplies",
        ),
        mk(
            "Staples",
            "https://www.staples.com",
            "US",
            "NA(United States)",
            "NA",
            "United States",
            "Staples",
            "Staples",
            "Office supplies; packaging; shipping materials",
            "US office/packaging supplies retailer-distributor",
            "Staples — distributes office and packaging/shipping supplies.",
            "staples.com packaging shipping supplies",
        ),
        mk(
            "Fastenal Canada",
            "https://www.fastenal.ca",
            "CA",
            "NA(Canada)",
            "NA",
            "Canada",
            "Fastenal",
            "Fastenal",
            "Industrial supplies; packaging-related consumables",
            "Canadian industrial distributor including packaging products",
            "Fastenal Canada — industrial distributor including packaging supplies.",
            "fastenal.ca industrial distributor",
        ),
        mk(
            "Staples Canada",
            "https://www.staples.ca",
            "CA",
            "NA(Canada)",
            "NA",
            "Canada",
            "Staples",
            "Staples",
            "Office supplies; packaging; shipping materials",
            "Canadian office/packaging supplies distributor",
            "Staples Canada — office and packaging supplies distributor.",
            "staples.ca",
        ),
        mk(
            "Grainger Canada",
            "https://www.grainger.ca",
            "CA",
            "NA(Canada)",
            "NA",
            "Canada",
            "W.W. Grainger",
            "Grainger",
            "Industrial supplies; packaging-related MRO",
            "Canadian industrial distributor including packaging products",
            "Grainger Canada — industrial distributor including packaging supplies.",
            "grainger.ca",
        ),
        mk(
            "Berlin Packaging EMEA",
            "https://www.berlinpackaging.eu",
            "NL",
            "EU(Europe)",
            "EU",
            "EMEA",
            "Berlin Packaging",
            "Berlin Packaging",
            "Glass; plastic; metal packaging containers and closures",
            "Berlin Packaging EMEA hybrid packaging supplier/distributor network",
            "Berlin Packaging EMEA — packaging supplier/distributor of containers and closures.",
            "berlinpackaging.eu hybrid packaging supplier; primarily distribution+services (monitor hybrid)",
        ),
    ]
)


def domain_of(url: str) -> str:
    try:
        return urlparse(url if "://" in url else "https://" + url).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def main():
    # Drop Berlin hybrid to respect OEM/hybrid rule
    global ROWS
    ROWS = [r for r in ROWS if "berlinpackaging.eu" not in r["Website"]]

    backup = FP.with_name(FP.stem + f"_pre_agent_manual4_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(FP, backup)
    wb = load_workbook(FP)
    ws = wb.active
    names, domains = set(), set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        names.add((r[0] or "").strip().lower())
        domains.add(domain_of(str(r[1] or "")))
    added, skipped = [], []
    for rowd in ROWS:
        d = domain_of(rowd["Website"])
        if rowd["Company"].lower() in names or d in domains:
            skipped.append(rowd["Company"])
            continue
        ws.append([rowd[h] for h in HEADERS])
        names.add(rowd["Company"].lower())
        domains.add(d)
        rec = {k: rowd[k] for k in HEADERS}
        rec["_evidence"] = rowd.get("_evidence")
        added.append(rec)
    for col in ws.columns:
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col[0].column_letter].width = 22
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(FP)
    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "added_count": len(added),
        "skipped": skipped,
        "final_rows": ws.max_row - 1,
        "added": added,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("added", len(added), "skipped", len(skipped), "final", audit["final_rows"])


if __name__ == "__main__":
    main()
