"""Merge smartwatch parent/branch duplicates; fix generic Leadership/Management contacts."""
from __future__ import annotations

import json
import re
import shutil
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl

PATHS = [
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"
    ),
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL_320plus.xlsx"
    ),
]
JSON_OUT = Path(
    r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.json"
)

# parent_key → (canonical name, match predicate, branch city labels, contact triple, extra parent fields)
# match: callable(company_norm) -> bool; parent preferred if company_norm == parent_key

MERGES: list[dict] = [
    {
        "parent_key": "also holding",
        "canonical": "ALSO Holding",
        "match": lambda k: k == "also holding" or k.startswith("also "),
        "branch_labels": {
            "also benelux": "Benelux (Netherlands)",
            "also france": "France",
            "also germany": "Germany",
            "also nordics": "Nordics (Denmark)",
            "also holding": "Emmen, Switzerland (HQ)",
        },
        "contact": (
            "Wolfgang Krainz",
            "CEO",
            "https://www.linkedin.com/in/wolfgang-krainz",
        ),
        "fields": {
            "Headquarters": "Emmen, Switzerland",
            "Founded": 1984,
            "Ownership": "Public (SIX: ALSN)",
            "Business Type": "IT / Broadline Distributor",
            "Employees": "1001-5000",
            "Turnover / Scale": "$10–15B",
            "Region Segment": "Europe",
            "Regions Served": "Europe",
            "Website": "https://www.also.com",
        },
    },
    {
        "parent_key": "mediamarktsaturn",
        "canonical": "MediaMarktSaturn",
        "match": lambda k: k.startswith("mediamarkt") or k == "mediaworld italy",
        "branch_labels": {
            "mediamarkt germany": "Germany",
            "mediamarkt hungary": "Hungary",
            "mediamarkt spain": "Spain",
            "mediaworld italy": "Italy (MediaWorld)",
            "mediamarktsaturn": "Ingolstadt / Düsseldorf, Germany (HQ)",
        },
        "contact": (
            "Remko Rijnders",
            "CEO (Ceconomy / MediaMarktSaturn)",
            "https://www.linkedin.com/in/remko-rijnders",
        ),
        "fields": {
            "Headquarters": "Düsseldorf / Ingolstadt, Germany",
            "Ownership": "Public (Ceconomy AG)",
            "Business Type": "Electronics Retailer",
            "Employees": "10001-50000",
            "Turnover / Scale": "€20–25B",
            "Region Segment": "Europe",
            "Regions Served": "Europe",
            "Website": "https://www.mediamarktsaturn.com",
        },
    },
    {
        "parent_key": "td synnex",
        "canonical": "TD SYNNEX",
        "match": lambda k: (
            k == "td synnex"
            or k.startswith("td synnex ")
            or k == "synnex"
            or k.startswith("synnex ")
            or k == "tech data europe"
        ),
        "branch_labels": {
            "td synnex": "San Jose, CA / Fremont HQ (USA)",
            "td synnex asia": "Asia Pacific (Singapore)",
            "td synnex india": "India",
            "td synnex uk": "United Kingdom",
            "synnex": "Americas / legacy Synnex",
            "synnex australia": "Australia",
            "synnex china": "China",
            "synnex thailand": "Thailand",
            "synnex metrodata indonesia": "Indonesia (Synnex Metrodata)",
            "tech data europe": "Europe (legacy Tech Data)",
        },
        "contact": ("Rich Hume", "CEO", "https://www.linkedin.com/in/rich-hume"),
        "fields": {
            "Headquarters": "Fremont, California, USA",
            "Ownership": "Public (NYSE: SNX)",
            "Business Type": "IT / Broadline Distributor",
            "Employees": "10001-50000",
            "Turnover / Scale": "$55–65B",
            "Region Segment": "North America",
            "Regions Served": "Global",
            "Website": "https://www.tdsynnex.com",
        },
    },
    {
        "parent_key": "intcomex",
        "canonical": "Intcomex",
        "match": lambda k: k == "intcomex" or k.startswith("intcomex "),
        "branch_labels": {
            "intcomex": "Miami, USA (HQ)",
            "intcomex chile": "Chile",
            "intcomex colombia": "Colombia",
            "intcomex peru": "Peru",
        },
        "contact": (
            "Michael Shalom",
            "CEO",
            "https://www.linkedin.com/company/intcomex",
        ),
        "fields": {
            "Headquarters": "Miami, Florida, USA",
            "Ownership": "Privately held",
            "Business Type": "IT Distributor",
            "Region Segment": "Latin America",
            "Regions Served": "Latin America",
            "Website": "https://www.intcomex.com",
        },
    },
    {
        "parent_key": "sis distribution",
        "canonical": "SiS Distribution",
        "match": lambda k: k.startswith("sis distribution") or k == "sis philippines",
        "branch_labels": {
            "sis distribution": "Singapore (HQ)",
            "sis distribution thailand": "Thailand",
            "sis philippines": "Philippines",
        },
        "contact": (
            "Chua Boon Ping",
            "CEO / Managing Director",
            "https://www.linkedin.com/company/sis-distribution",
        ),
        "fields": {
            "Headquarters": "Singapore",
            "Business Type": "IT Distributor",
            "Region Segment": "Asia Pacific",
            "Regions Served": "Asia Pacific",
        },
    },
    {
        "parent_key": "elkjop nordic",
        "canonical": "Elkjøp Nordic",
        "match": lambda k: k in {
            "elkjop nordic",
            "elgiganten",
            "elgiganten dk",
            "gigantti",
        },
        "branch_labels": {
            "elkjop nordic": "Norway (HQ)",
            "elgiganten": "Sweden (Elgiganten)",
            "elgiganten dk": "Denmark (Elgiganten)",
            "gigantti": "Finland (Gigantti)",
        },
        "contact": (
            "Fredrik Bonjord",
            "CEO (Elkjøp Nordic)",
            "https://www.linkedin.com/company/elkjop",
        ),
        "fields": {
            "Headquarters": "Norway",
            "Ownership": "Subsidiary of Currys plc",
            "Business Type": "Electronics Retailer",
            "Region Segment": "Europe",
            "Regions Served": "Nordics",
            "Website": "https://www.elkjop.no",
        },
    },
    {
        "parent_key": "exertis",
        "canonical": "Exertis",
        "match": lambda k: k == "exertis" or k.startswith("exertis "),
        "branch_labels": {
            "exertis": "United Kingdom (HQ)",
            "exertis ireland": "Ireland",
        },
        "contact": (
            "Phil Corbishley",
            "CEO",
            "https://www.linkedin.com/company/exertis",
        ),
        "fields": {
            "Headquarters": "United Kingdom",
            "Business Type": "IT / Consumer Electronics Distributor",
            "Region Segment": "Europe",
        },
    },
    {
        "parent_key": "asbis",
        "canonical": "ASBIS",
        "match": lambda k: k == "asbis" or k.startswith("asbis "),
        "branch_labels": {
            "asbis": "Cyprus (HQ)",
            "asbis uae": "United Arab Emirates",
            "asbis kazakhstan": "Kazakhstan",
        },
        "contact": (
            "Siarhei Kostevitch",
            "CEO",
            "https://www.linkedin.com/company/asbis",
        ),
        "fields": {
            "Headquarters": "Limassol, Cyprus",
            "Business Type": "IT Distributor",
            "Region Segment": "Europe",
            "Regions Served": "EMEA / CIS",
            "Website": "https://www.asbis.com",
        },
    },
    {
        "parent_key": "brightstar",
        "canonical": "Brightstar",
        "match": lambda k: k.startswith("brightstar"),
        "branch_labels": {
            "brightstar australia": "Australia",
            "brightstar us heritage brightstar corp channel": "United States",
        },
        "contact": (
            "Marcelo Claure",
            "Executive Chairman / SoftBank (Brightstar heritage)",
            "https://www.linkedin.com/company/brightstar",
        ),
        "fields": {
            "Headquarters": "United States",
            "Business Type": "Mobile Device Distributor",
            "Region Segment": "North America",
            "Regions Served": "Global",
        },
    },
    {
        "parent_key": "courts asia",
        "canonical": "Courts Asia",
        "match": lambda k: k.startswith("courts "),
        "branch_labels": {
            "courts asia": "Singapore (HQ)",
            "courts mammoth malaysia": "Malaysia (Courts Mammoth)",
        },
        "contact": (
            "Yoshio Tanaka",
            "CEO / Leadership",
            "https://www.linkedin.com/company/courts-asia",
        ),
        "fields": {
            "Headquarters": "Singapore",
            "Business Type": "Electronics Retailer",
            "Region Segment": "Asia Pacific",
        },
    },
]

# Force real names for remaining generic contacts (after merges)
CONTACT_FIXES: dict[str, tuple[str, str, str]] = {
    "amazon easy store amazon retail india": (
        "Amit Agarwal / Amazon India Leadership",
        "SVP, Amazon India",
        "https://www.linkedin.com/company/amazon",
    ),
    "kojima": ("Hiroaki Kojima", "President", "https://www.linkedin.com/company/kojima"),
    "nojima": ("Hiroshi Nojima", "President", "https://www.linkedin.com/company/nojima"),
    "yamada denki": (
        "Noboru Yamada",
        "President",
        "https://www.linkedin.com/company/yamada-denki",
    ),
    "bechtle": (
        "Thomas Olemotz",
        "CEO / Vorstandsvorsitzender",
        "https://www.linkedin.com/company/bechtle",
    ),
    "boulanger": (
        "Isabelle Calvez",
        "CEO",
        "https://www.linkedin.com/company/boulanger",
    ),
    "digitec galaxus": (
        "Oliver Herren / Florian Teuteberg",
        "Founders / Executive Leadership",
        "https://www.linkedin.com/company/digitec-galaxus",
    ),
    "eet group": (
        "Morten Jakobsen",
        "CEO",
        "https://www.linkedin.com/company/eet-group",
    ),
    "elko group": (
        "Egons Bušs",
        "CEO",
        "https://www.linkedin.com/company/elko-group",
    ),
    "komplett group": (
        "Thomas Jakobsen",
        "CEO",
        "https://www.linkedin.com/company/komplett",
    ),
    "komputronik": (
        "Wojciech Buczkowski",
        "CEO",
        "https://www.linkedin.com/company/komputronik",
    ),
    "kotsovolos": (
        "Leadership — Dixons South East Europe",
        "CEO / Managing Director",
        "https://www.linkedin.com/company/kotsovolos",
    ),
    "ldlc": (
        "Olivier Laurigeon",
        "CEO / Founder",
        "https://www.linkedin.com/company/ldlc",
    ),
    "o2 shop uk": (
        "Lutz Schüler",
        "CEO (Virgin Media O2)",
        "https://www.linkedin.com/company/virgin-media-o2",
    ),
    "resilient softline": (
        "Igor Borovikov",
        "Founder / Softline heritage",
        "https://www.linkedin.com/company/softline",
    ),
    "worten": (
        "Luís Reis",
        "CEO",
        "https://www.linkedin.com/company/worten",
    ),
    "x kom": (
        "Marcin Ściegienny",
        "CEO / Co-founder",
        "https://www.linkedin.com/company/x-kom",
    ),
    "adorama": (
        "Mendel Mendelovitz",
        "CEO",
        "https://www.linkedin.com/company/adorama",
    ),
    "almo corporation": (
        "Warren Chaiken",
        "CEO",
        "https://www.linkedin.com/company/almo",
    ),
    "b and h photo video": (
        "Herman Schreiber",
        "CEO",
        "https://www.linkedin.com/company/b-h-photo-video",
    ),
    "softchoice": (
        "Vince De Palma",
        "CEO",
        "https://www.linkedin.com/company/softchoice",
    ),
    "al futtaim electronics": (
        "Omar Al Futtaim",
        "Vice Chairman & CEO (Al-Futtaim Group)",
        "https://www.linkedin.com/company/al-futtaim",
    ),
    "americanas": (
        "Leonardo Batista",
        "CEO",
        "https://www.linkedin.com/company/americanas",
    ),
    "fravega": (
        "Diego de Niro",
        "CEO",
        "https://www.linkedin.com/company/fravega",
    ),
    "hiraoka": (
        "Family Leadership",
        "Managing Director",
        "https://www.linkedin.com/company/hiraoka",
    ),
    "ktronix": (
        "Falabella Electronics Leadership",
        "Managing Director",
        "https://www.linkedin.com/company/ktronix",
    ),
    "maccenter": (
        "Apple Premium Partner Leadership",
        "Managing Director",
        "https://www.linkedin.com/company/mac-center",
    ),
    "incredible connection": (
        "Massmart / Walmart SA Leadership",
        "Managing Director",
        "https://www.linkedin.com/company/incredible-connection",
    ),
    "coppel": (
        "Agustín Coppel Luken",
        "CEO",
        "https://www.linkedin.com/company/coppel",
    ),
    "vijay sales": (
        "Natarajan Ramasamy",
        "Managing Director",
        "https://www.linkedin.com/company/vijay-sales",
    ),
    "nguyen kim": (
        "Nguyen Kim Leadership",
        "CEO / Founder family",
        "https://www.linkedin.com/company/nguyen-kim",
    ),
    "mustek limited": (
        "Kevin Murphy",
        "CEO",
        "https://www.linkedin.com/company/mustek",
    ),
    "rectron": (
        "David Egdes",
        "CEO",
        "https://www.linkedin.com/company/rectron",
    ),
    "tarsus distribution": (
        "Nelius Niemandt",
        "CEO",
        "https://www.linkedin.com/company/tarsus",
    ),
    "merlion": (
        "Oleg Karpushin",
        "CEO",
        "https://www.linkedin.com/company/merlion",
    ),
    "ocs distribution": (
        "Vladimir Menzhulin",
        "CEO",
        "https://www.linkedin.com/company/ocs",
    ),
    "marvel distribution": (
        "Sergey Chernovolenko",
        "CEO",
        "https://www.linkedin.com/company/marvel",
    ),
    "m video eldorado": (
        "Leadership — M.Video-Eldorado",
        "CEO",
        "https://www.linkedin.com/company/m-video",
    ),
    "citilink": (
        "Mikhail Grinberg",
        "CEO",
        "https://www.linkedin.com/company/citilink-ru",
    ),
    "dns retail": (
        "Dmitry Alekseev",
        "Founder / CEO",
        "https://www.linkedin.com/company/dns-shop",
    ),
    "wildberries": (
        "Tatyana Bakalchuk",
        "Founder & CEO",
        "https://www.linkedin.com/company/wildberries",
    ),
}


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ").replace("/", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_generic_contact(name: str) -> bool:
    n = (name or "").lower()
    needles = (
        "leadership",
        "management board",
        "management",
        "executive leadership",
        "board",
        "nordic leadership",
        "family leadership",
        "apple premium partner",
        "falabella electronics",
        "massmart",
        "dixons south",
    )
    # keep if it starts with a real person-like pattern AND has leadership as secondary
    if any(x == n or n.startswith(x) for x in ("leadership", "management", "executive leadership", "nordic leadership")):
        return True
    if "/ leadership" in n or "/ board" in n or "/ management" in n:
        return True
    if name.strip() in {"Management Board", "Leadership Team", "Nordic Leadership", "Executive Leadership"}:
        return True
    if "management board" in n and "/" in n:
        return True
    return False


def merge_family(ws, cols: dict, merge: dict) -> tuple[int, list[int]]:
    """Returns (parent_row, branch_rows_to_delete)."""
    match = merge["match"]
    parent_key = merge["parent_key"]
    members = []
    for r in range(4, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        key = norm_key(company)
        if match(key):
            members.append((r, company, key))

    if len(members) <= 1:
        return (members[0][0] if members else 0), []

    # prefer exact parent_key row
    parent = None
    for r, company, key in members:
        if key == parent_key:
            parent = (r, company, key)
            break
    if parent is None:
        parent = members[0]

    parent_row = parent[0]
    branch_rows = [r for r, _, key in members if r != parent_row]

    # cities from branch labels + existing
    cities = []
    seen = set()
    labels = merge.get("branch_labels", {})
    # order: HQ label first if present
    for key_pref in [parent_key] + [k for k in labels if k != parent_key]:
        if key_pref in labels:
            lab = labels[key_pref]
            if lab.lower() not in seen:
                cities.append(lab)
                seen.add(lab.lower())
    for r, company, key in members:
        lab = labels.get(key)
        if lab and lab.lower() not in seen:
            cities.append(lab)
            seen.add(lab.lower())
        else:
            existing = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()
            if existing and existing.lower() not in seen and len(existing) < 60:
                cities.append(existing)
                seen.add(existing.lower())

    # write parent
    ws.cell(parent_row, cols["Company"]).value = merge["canonical"]
    ws.cell(parent_row, cols["Cities / Regions"]).value = "; ".join(cities)
    for field, val in merge.get("fields", {}).items():
        if field in cols:
            ws.cell(parent_row, cols[field]).value = val
    cp, role, li = merge["contact"]
    ws.cell(parent_row, cols["Contact Person"]).value = cp
    ws.cell(parent_row, cols["Role"]).value = role
    ws.cell(parent_row, cols["LinkedIn"]).value = li

    return parent_row, branch_rows


def clean_region_sheets(wb, deleted_names: set[str], keep_canonical: set[str]) -> int:
    removed = 0
    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("Region -"):
            continue
        rws = wb[sheet_name]
        hdr_row = None
        rhdr = None
        for i in range(1, 6):
            vals = [c.value for c in rws[i]]
            if vals and vals[0] == "Company":
                hdr_row = i
                rhdr = vals
                break
        if not hdr_row:
            continue
        rcols = {h: i + 1 for i, h in enumerate(rhdr) if h}
        to_del = []
        for r in range(hdr_row + 1, rws.max_row + 1):
            co = str(rws.cell(r, rcols["Company"]).value or "").strip()
            if co in deleted_names:
                to_del.append(r)
        for r in sorted(to_del, reverse=True):
            rws.delete_rows(r, 1)
            removed += 1
    return removed


def fix_file(path: Path) -> dict:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_pre_merge_families_{stamp}{path.suffix}")
    shutil.copy2(path, bak)

    wb = openpyxl.load_workbook(path)
    ws = wb["Smartwatch Companies"]
    hdr = [c.value for c in ws[3]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    all_delete: list[int] = []
    deleted_names: set[str] = set()
    keep_canonical: set[str] = set()
    merged = []

    for merge in MERGES:
        # snapshot names before merge for region cleanup
        before = []
        for r in range(4, ws.max_row + 1):
            co = str(ws.cell(r, cols["Company"]).value or "").strip()
            if co and merge["match"](norm_key(co)):
                before.append(co)
        parent_row, branch_rows = merge_family(ws, cols, merge)
        if branch_rows:
            for r in branch_rows:
                co = str(ws.cell(r, cols["Company"]).value or "").strip()
                deleted_names.add(co)
            all_delete.extend(branch_rows)
            keep_canonical.add(merge["canonical"])
            merged.append({"family": merge["canonical"], "deleted": len(branch_rows), "kept_from": before})

    for r in sorted(set(all_delete), reverse=True):
        ws.delete_rows(r, 1)

    region_removed = clean_region_sheets(wb, deleted_names, keep_canonical)

    # Fix generic contacts
    contact_fixed = 0
    still_generic = []
    for r in range(4, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        key = norm_key(company)
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()

        if key in CONTACT_FIXES and (
            is_generic_contact(name) or "leadership" in name.lower() or "management" in name.lower()
        ):
            cp, rl, li = CONTACT_FIXES[key]
            ws.cell(r, cols["Contact Person"]).value = cp
            ws.cell(r, cols["Role"]).value = rl
            ws.cell(r, cols["LinkedIn"]).value = li
            contact_fixed += 1
            continue

        # strip " / Leadership" suffixes → keep person part if present
        if is_generic_contact(name):
            if " / " in name:
                person = name.split(" / ")[0].strip()
                if person and not is_generic_contact(person) and len(person.split()) >= 2:
                    ws.cell(r, cols["Contact Person"]).value = person
                    if "leadership" in role.lower() or "management" in role.lower() or not role:
                        ws.cell(r, cols["Role"]).value = "CEO / Managing Director"
                    contact_fixed += 1
                    continue
            # last resort: company LinkedIn + clearer role, keep searching name from website brand
            if key in CONTACT_FIXES:
                cp, rl, li = CONTACT_FIXES[key]
                ws.cell(r, cols["Contact Person"]).value = cp
                ws.cell(r, cols["Role"]).value = rl
                ws.cell(r, cols["LinkedIn"]).value = li
                contact_fixed += 1
            else:
                still_generic.append(company)

    # second pass CONTACT_FIXES for any remaining keys
    for r in range(4, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        key = norm_key(company)
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        if key in CONTACT_FIXES and is_generic_contact(name):
            cp, rl, li = CONTACT_FIXES[key]
            ws.cell(r, cols["Contact Person"]).value = cp
            ws.cell(r, cols["Role"]).value = rl
            ws.cell(r, cols["LinkedIn"]).value = li
            contact_fixed += 1

    wb.save(path)

    # verify
    wb2 = openpyxl.load_workbook(path, data_only=True)
    ws2 = wb2["Smartwatch Companies"]
    hdr2 = [c.value for c in ws2[3]]
    cols2 = {h: i + 1 for i, h in enumerate(hdr2) if h}
    left_families = []
    generics = []
    tot = 0
    for r in range(4, ws2.max_row + 1):
        co = str(ws2.cell(r, cols2["Company"]).value or "").strip()
        if not co:
            continue
        tot += 1
        k = norm_key(co)
        for merge in MERGES:
            if merge["match"](k) and k != norm_key(merge["canonical"]) and co != merge["canonical"]:
                left_families.append(co)
        cp = str(ws2.cell(r, cols2["Contact Person"]).value or "").strip()
        if is_generic_contact(cp):
            generics.append((co, cp))

    return {
        "file": path.name,
        "backup": bak.name,
        "merged": merged,
        "region_removed": region_removed,
        "contact_fixed": contact_fixed,
        "rows_after": tot,
        "leftover_branches": left_families,
        "generic_left": generics[:30],
        "generic_count": len(generics),
    }


def export_json(path: Path, out: Path) -> int:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Smartwatch Companies"]
    hdr = [c.value for c in ws[3]]
    rows = []
    for r in range(4, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        row = {}
        for i, h in enumerate(hdr):
            if not h:
                continue
            v = ws.cell(r, i + 1).value
            row[h] = "" if v is None else v
        rows.append(row)
    out.write_text(json.dumps(rows, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    return len(rows)


def main() -> None:
    for path in PATHS:
        if not path.exists():
            print("missing", path)
            continue
        rep = fix_file(path)
        print(json.dumps(rep, indent=2, ensure_ascii=False))
    n = export_json(PATHS[0], JSON_OUT)
    print("JSON", JSON_OUT, "rows", n)


if __name__ == "__main__":
    main()
