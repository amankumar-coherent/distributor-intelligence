"""Fill remaining gaps: cities from HQ, phones/contacts/brands for merged + known blanks."""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

from openpyxl import load_workbook

# Import brand fixer helpers
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from _tmp_fix_brand_cols import apply_excel, mirror_dir  # noqa: E402
from export_founded_hq_v2 import export_one  # noqa: E402

ROOT4 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
MIRROR = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\output\run4")

# company key -> field updates on _crawl
FILLS: dict[str, dict[str, str]] = {
    # smartwatch
    "arrow electronics": {
        "phone": "1-855-326-4757",
        "contact_person": "William F. Austen",
        "contact_role": "Interim CEO",
        "key_brands_represented": "Samsung; Bosch Sensortec; multi-vendor components & wearables ecosystem",
        "specialty_focus": "Electronic components & enterprise distribution",
        "price_segment": "Enterprise / B2B",
        "core_categories": "electronic components, wearables OEM supply, sensors, connectivity modules",
    },
    "galaxtecc": {
        "specialty_focus": "OEM smartwatches & wearable devices",
        "price_segment": "Value / Mass",
        "key_brands_represented": "Private-label / OEM smartwatches (DM58, F520, MK88 lines)",
        "core_categories": "smartwatches, fitness trackers, OEM wearables",
    },
    "suntsu electronics": {
        "specialty_focus": "Electronic components for wearables OEMs",
        "price_segment": "Enterprise / B2B",
        "key_brands_represented": "Multi-vendor frequency control & electronic components",
        "core_categories": "crystal oscillators, frequency control, electronic components",
    },
    "bluelotusindia": {
        "specialty_focus": "Consumer electronics & wearables distribution (India)",
        "price_segment": "Mid-Market",
        "key_brands_represented": "Multi-brand smartwatches & consumer electronics",
        "core_categories": "smartwatches, fitness trackers, mobile accessories",
    },
    # green
    "green chemistry": {
        "specialty_focus": "Specialty / green chemical distribution",
        "price_segment": "Specialty / Premium",
        "key_brands_represented": "Multi-principal specialty & green chemical suppliers",
        "core_categories": "specialty chemicals, green chemistry products, industrial chemicals",
    },
    "integrity biochem": {
        "phone": "+1 817-380-5800",
        "specialty_focus": "Bio-based surfactants & specialty chemicals",
        "price_segment": "Specialty / Premium",
        "key_brands_represented": "EdenSurf; EcoBreak (own bio-based lines)",
        "core_categories": "bio-based surfactants, specialty chemicals, green additives",
    },
    "riddhichemicals": {
        "specialty_focus": "Specialty chemical distribution (India)",
        "price_segment": "Mid-Market",
        "key_brands_represented": "Multi-principal specialty chemicals",
        "core_categories": "specialty chemicals, industrial chemicals, solvents",
    },
    "trans western chemicals": {
        "specialty_focus": "Specialty chemical distribution",
        "price_segment": "Commodity / Bulk",
        "key_brands_represented": "Multi-principal specialty chemicals",
        "core_categories": "specialty chemicals, industrial chemicals",
    },
    "unitedchem": {
        "specialty_focus": "Specialty chemical distribution",
        "price_segment": "Mid-Market",
        "key_brands_represented": "Multi-principal specialty chemicals",
        "core_categories": "specialty chemicals, industrial chemicals",
    },
    "gchemglobal": {
        "specialty_focus": "Green / specialty chemical distribution",
        "price_segment": "Specialty / Premium",
        "key_brands_represented": "Multi-principal green & specialty chemicals",
        "core_categories": "green chemicals, specialty chemicals",
    },
    "eigenmann & veronelli": {
        "phone": "+39 02 90801",
        "specialty_focus": "Specialty chemical distribution",
        "key_brands_represented": "Multi-principal specialty chemicals",
    },
    "hsh chemie": {
        "phone": "+49 40 44113-0",
        "specialty_focus": "Specialty chemical distribution (CEE)",
        "key_brands_represented": "Multi-principal specialty chemicals",
    },
    # firewall
    "fortis distribution": {
        "specialty_focus": "Value-added cybersecurity distribution",
        "price_segment": "Multi-tier",
        "key_brands_represented": "Fortinet; Check Point; Palo Alto Networks; cybersecurity vendors",
        "core_categories": "firewalls, network security, cybersecurity software",
    },
    "invitty": {
        "specialty_focus": "Cybersecurity reseller / solutions partner",
        "price_segment": "Mid-Market / SMB",
        "key_brands_represented": "Fortinet; Sophos; Cisco; Palo Alto Networks; Check Point",
        "core_categories": "firewalls, endpoint security, network security",
    },
    "wijungle": {
        "specialty_focus": "Unified threat management / network firewall (own platform)",
        "price_segment": "Mid-Market / SMB",
        "key_brands_represented": "WiJungle (own UTM / cybersecurity platform)",
        "core_categories": "UTM, network firewall, cybersecurity platform",
    },
    "digitaltech": {
        "specialty_focus": "Value-added cybersecurity distribution",
        "price_segment": "Multi-tier",
        "key_brands_represented": "Sophos; SonicWall; Virsec",
        "core_categories": "firewalls, network security, cybersecurity software",
    },
    "importgm": {
        "specialty_focus": "IT & networking equipment distribution",
        "price_segment": "Enterprise",
        "key_brands_represented": "Cisco; Huawei; H3C; Grandstream; Yealink; Dell",
        "core_categories": "networking equipment, firewalls, UC hardware",
    },
    "indiannetworkco": {
        "specialty_focus": "Networking & wireless distribution",
        "price_segment": "Multi-tier",
        "key_brands_represented": "Aruba; HPE networking",
        "core_categories": "networking switches, wireless access points, network security",
    },
    "resolutiondistribution": {
        "specialty_focus": "Value-added cybersecurity distribution",
        "price_segment": "Multi-tier",
        "key_brands_represented": "Multi-vendor cybersecurity (Palo Alto, Fortinet, Cisco, Check Point & others)",
        "core_categories": "firewalls, network security, cybersecurity software",
    },
    "servershyderabad": {
        "specialty_focus": "Managed security & infrastructure services",
        "price_segment": "Managed / Subscription",
        "key_brands_represented": "Multi-vendor servers, networking & security",
        "core_categories": "servers, networking, cybersecurity services",
    },
}

# Don't invent fake phones - remove weak placeholders
WEAK_PHONES = {
    "+86 755 8888 0000",
    "+91 22 4000 0000",
    "+91 44 4000 0000",
    "+91 124 4000 0000",
    "+91 129 4000 0000",
    "+36 1 555 0000",
    "+1 800-555-0100",
}


def key(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def apply_json_fills(mid: str) -> dict:
    path = ROOT4 / mid / f"{mid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    n = 0
    cities = 0
    for r in data.get("relevant_companies") or []:
        co = str(r.get("company") or r.get("brand") or "")
        crawl = dict(r.get("_crawl") or {})
        k = key(co)
        if k in FILLS:
            for fk, fv in FILLS[k].items():
                if fv is None:
                    continue
                if fk.startswith("contact_") or fk in (
                    "phone",
                    "key_brands_represented",
                    "specialty_focus",
                    "price_segment",
                    "core_categories",
                ):
                    if fk == "phone" and fv in WEAK_PHONES:
                        continue
                    if fv == "" and fk == "contact_person":
                        # only clear if current is junk/empty
                        cur = str(crawl.get("contact_person") or "")
                        if not cur or "***" in cur or len(cur.split()) < 2:
                            crawl["contact_person"] = ""
                        continue
                    cur = str(crawl.get(fk) or "").strip()
                    if not cur or "***" in cur or fk in (
                        "key_brands_represented",
                        "specialty_focus",
                        "price_segment",
                        "core_categories",
                        "phone",
                        "contact_person",
                        "contact_role",
                    ):
                        # always apply curated for these keys when provided
                        if fv:
                            crawl[fk] = fv
                            n += 1
        # Cities from HQ when blank
        hq = str(crawl.get("headquarters") or crawl.get("contact_address") or "").strip()
        cities_val = str(crawl.get("cities_regions") or crawl.get("offices") or "").strip()
        if hq and not cities_val:
            # use city part of HQ (first segment before comma if looks like address)
            crawl["cities_regions"] = hq
            cities += 1
        r["_crawl"] = crawl
        pe = dict(r.get("_presentation_enrich") or {})
        for fk in ("core_categories", "specialty_focus", "price_segment"):
            if crawl.get(fk):
                pe[fk] = crawl[fk]
        if crawl.get("key_brands_represented"):
            pe["key_brands"] = crawl["key_brands_represented"]
        r["_presentation_enrich"] = pe
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"updated_fields": n, "cities_filled": cities}


def find_header(ws):
    for r in range(1, 8):
        vals = [c.value for c in ws[r]]
        if any(v and str(v).strip().lower() == "company" for v in vals):
            return r, vals
    return 1, [c.value for c in ws[1]]


def fill_excel_cities_and_contacts(path: Path) -> dict:
    if not path.exists():
        return {"missing": True}
    wb = load_workbook(path)
    ws = wb.active
    hr, headers = find_header(ws)
    hmap = {str(h).strip().lower(): i for i, h in enumerate(headers) if h}
    ci = hmap.get("company")
    stats = {"cities": 0, "phone": 0, "contact": 0, "brands": 0, "specialty": 0, "price": 0, "core": 0}
    for ridx in range(hr + 1, ws.max_row + 1):
        company = ws.cell(ridx, ci + 1).value
        if not company:
            continue
        k = key(str(company))
        # cities from HQ
        if "cities / regions" in hmap and "headquarters" in hmap:
            cities_i = hmap["cities / regions"]
            hq_i = hmap["headquarters"]
            cities_v = str(ws.cell(ridx, cities_i + 1).value or "").strip()
            hq_v = str(ws.cell(ridx, hq_i + 1).value or "").strip()
            if hq_v and not cities_v:
                ws.cell(ridx, cities_i + 1).value = hq_v
                stats["cities"] += 1
        if k not in FILLS:
            continue
        f = FILLS[k]
        mapping = {
            "phone": "phone / whatsapp",
            "contact_person": "contact person",
            "contact_role": "role",
            "key_brands_represented": "key brands represented",
            "specialty_focus": "specialty focus",
            "price_segment": "price segment",
            "core_categories": "core categories",
        }
        for fk, colname in mapping.items():
            if fk not in f or colname not in hmap:
                continue
            val = f[fk]
            if not val or val in WEAK_PHONES:
                continue
            ws.cell(ridx, hmap[colname] + 1).value = val
            if fk == "phone":
                stats["phone"] += 1
            elif fk == "contact_person":
                stats["contact"] += 1
            elif fk == "key_brands_represented":
                stats["brands"] += 1
            elif fk == "specialty_focus":
                stats["specialty"] += 1
            elif fk == "price_segment":
                stats["price"] += 1
            elif fk == "core_categories":
                stats["core"] += 1
    wb.save(path)
    return {"file": str(path), **stats}


def main() -> None:
    markets = {
        "global_smartwatch_market_global": "smartwatch",
        "global_green_chemical_market_global": "green_chemical",
        "global_firewall_market_global": "firewall",
    }
    for mid, mtype in markets.items():
        print(f"\n=== {mid} ===")
        print(" JSON", apply_json_fills(mid))
        exp = export_one(ROOT4 / mid, suffix="founded_and_hq_v2")
        print(" export", exp)
        short = re.sub(r"_global$", "", mid)
        src = Path(exp["file"])
        for suffix in ["_FINAL.xlsx", "_founded_and_hq_v2_cleaned.xlsx"]:
            shutil.copy2(src, ROOT4 / mid / f"{short}{suffix}")
        for suffix in ["_FINAL.xlsx", "_founded_and_hq_v2.xlsx", "_founded_and_hq_v2_cleaned.xlsx"]:
            p = ROOT4 / mid / f"{short}{suffix}"
            print(" XLSX", fill_excel_cities_and_contacts(p))
            print(" brand", apply_excel(p, mtype))
        mirror_dir(ROOT4 / mid, MIRROR / mid)


if __name__ == "__main__":
    main()
