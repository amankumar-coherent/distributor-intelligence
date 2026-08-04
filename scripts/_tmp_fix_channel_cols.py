"""Market-aware fill for Retail / Specialty Grocers / Foodservice / E-com /
Distribution Type / Regions / Regional Extensions / Warehouse / Delivery.

Smartwatch & Firewall: Foodservice + Specialty Grocers = Not applicable.
Green Chemical: Foodservice only when life-sciences/food ingredients; else No.
"""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

from openpyxl import load_workbook

ROOT4 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
MIRROR = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\output\run4")

MARKETS = {
    "global_smartwatch_market_global": "smartwatch",
    "global_green_chemical_market_global": "green_chemical",
    "global_firewall_market_global": "firewall",
}

SMARTWATCH_RETAILERS = {
    "walmart",
    "target",
    "macys",
    "macy",
    "jcpenney",
    "nordstrom rack",
    "nordstrom",
    "zales",
    "b&h photo video",
    "jomashop",
    "newegg commerce",
    "cover company",
    "wildberries",
}
SMARTWATCH_ECOM = SMARTWATCH_RETAILERS | {
    "chinavasion",
    "everfulwholesale",
    "usmphones",
    "eumobilewholesalers",
    "digikey",
    "mouser electronics",
    "newark",
    "farnell",
    "rs components",
}

STREET_RE = re.compile(
    r"\b(?:street|st\.|suite|office|road|rd\.|avenue|ave\.|blvd|building|floor|"
    r"room\s+\d|no\.\s*\d|\d{1,5}\s+\w)",
    re.I,
)

COUNTRY_FROM_HQ = re.compile(
    r"(united states|usa|u\.s\.a\.|india|china|united kingdom|uk|germany|france|"
    r"italy|spain|netherlands|uae|united arab emirates|saudi arabia|singapore|"
    r"japan|australia|canada|brazil|mexico|israel|turkey|switzerland|poland|"
    r"hungary|czech|austria|belgium|sweden|norway|denmark|finland|ireland|"
    r"south africa|hong kong|taiwan|korea|malaysia|thailand|indonesia|vietnam|"
    r"nepal|bahrain|qatar|kuwait|oman|egypt|nigeria|kenya|uruguay|argentina|"
    r"chile|colombia|peru|russia|ukraine|romania|bulgaria|greece|portugal)",
    re.I,
)


def key(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def looks_like_street(val: str) -> bool:
    s = (val or "").strip()
    if not s:
        return False
    if STREET_RE.search(s) and (re.search(r"\d", s) or "suite" in s.lower() or "office" in s.lower()):
        return True
    if len(s) > 80 and "," in s and re.search(r"\d", s):
        return True
    return False


def hq_country(hq: str) -> str:
    m = COUNTRY_FROM_HQ.search(hq or "")
    if not m:
        return ""
    raw = m.group(1).lower()
    mapping = {
        "usa": "United States",
        "u.s.a.": "United States",
        "united states": "United States",
        "uk": "United Kingdom",
        "united kingdom": "United Kingdom",
        "uae": "United Arab Emirates",
        "united arab emirates": "United Arab Emirates",
    }
    return mapping.get(raw, raw.title())


def infer_distribution_type(market: str, business_type: str, company: str) -> str:
    bt = (business_type or "").lower()
    co = company.lower()
    if market == "smartwatch":
        if key(company) in SMARTWATCH_RETAILERS or "retail" in bt:
            return "Direct retail / Retailer"
        if any(x in bt for x in ("wholesale",)):
            return "Wholesale"
        if any(x in bt for x in ("manufactur", "oem")) or "galaxtecc" in co:
            return "Direct / Manufacturer-led"
        if "distribut" in bt or "partner" in bt:
            return "Distributor"
        return "Distributor"
    if market == "green_chemical":
        if "manufactur" in bt:
            return "Direct / Manufacturer-led"
        if "trad" in bt:
            return "Trading / Wholesale"
        if "bulk" in bt:
            return "Bulk B2B"
        return "Distributor"
    # firewall
    if any(x in bt for x in ("msp", "managed")):
        return "Managed service / MSP"
    if any(x in bt for x in ("integrat", "si")):
        return "System integrator"
    if any(x in bt for x in ("reseller", "var", "vad")):
        return "Value-added distributor / Reseller"
    if "distribut" in bt or "partner" in bt:
        return "Distributor"
    if "service" in bt:
        return "Service provider"
    return "Distributor"


def infer_channels(market: str, company: str, business_type: str, summary: str, website: str) -> dict[str, str]:
    bt = (business_type or "").lower()
    text = f"{company} {business_type} {summary} {website}".lower()
    k = key(company)
    dist = infer_distribution_type(market, business_type, company)

    if market == "smartwatch":
        retail = "Yes" if (k in SMARTWATCH_RETAILERS or "retail" in bt) else "No"
        specialty = "Not applicable"
        foodservice = "Not applicable"
        ecom = "Yes" if (
            k in SMARTWATCH_ECOM
            or "e-commerce" in bt
            or "ecommerce" in bt
            or any(x in text for x in ("shop online", "buy online", "add to cart", "webshop", "online store"))
            or dist == "Direct retail / Retailer"
        ) else ("Yes" if any(x in k for x in ("digikey", "mouser", "newegg", "farnell", "newark")) else "No")
        # Big distributors usually have B2B e-commerce portals
        if dist == "Distributor" and any(
            x in k for x in ("ingram", "synnex", "d&h", "esprinet", "arrow", "td synnex", "tech data")
        ):
            ecom = "Yes"
        warehouse = "Yes" if dist in ("Distributor", "Wholesale", "Direct retail / Retailer") else "No"
        delivery = "Yes" if warehouse == "Yes" or ecom == "Yes" else "No"
        return {
            "retail_chains": retail,
            "specialty_grocers": specialty,
            "foodservice": foodservice,
            "ecommerce": ecom,
            "distribution_type": dist,
            "warehouse_logistics": warehouse,
            "delivery_storage": delivery,
        }

    if market == "green_chemical":
        food_related = any(
            x in text
            for x in (
                "food",
                "nutrition",
                "flavour",
                "flavor",
                "fragrance",
                "personal care",
                "cosmetic",
                "pharma",
                "life science",
                "horeca",
                "foodservice",
            )
        )
        retail = "No"  # B2B chemical channel, not grocery retail
        specialty = "Yes" if food_related and any(x in text for x in ("specialty", "natural", "organic", "retail")) else "No"
        # Prefer Not applicable when purely industrial
        if not food_related:
            specialty = "Not applicable"
            foodservice = "Not applicable"
        else:
            foodservice = "Yes" if any(x in text for x in ("foodservice", "horeca", "food", "nutrition", "beverage")) else "No"
        ecom = "Yes" if any(x in text for x in ("e-commerce", "ecommerce", "online portal", "webshop", "customer portal")) else "No"
        # Major distributors often have digital ordering
        if any(x in k for x in ("brenntag", "imcd", "azelis", "barentz", "caldic", "dksh", "univar")):
            ecom = "Yes"
        warehouse = "Yes" if dist.startswith("Distributor") or "Trading" in dist or dist == "Bulk B2B" else "No"
        delivery = "Yes" if warehouse == "Yes" else "No"
        return {
            "retail_chains": retail,
            "specialty_grocers": specialty,
            "foodservice": foodservice,
            "ecommerce": ecom,
            "distribution_type": dist,
            "warehouse_logistics": warehouse,
            "delivery_storage": delivery,
        }

    # firewall — no grocery/foodservice
    retail = "No"
    specialty = "Not applicable"
    foodservice = "Not applicable"
    ecom = "Yes" if any(
        x in text for x in ("e-commerce", "ecommerce", "partner portal", "online ordering", "marketplace")
    ) or any(
        x in k
        for x in (
            "ingram",
            "synnex",
            "td synnex",
            "d&h",
            "also",
            "redington",
            "westcon",
            "arrow",
            "scansource",
            "connection",
            "cdw",
            "softchoice",
            "softcat",
        )
    ) else "No"
    warehouse = "Yes" if "distribut" in dist.lower() or "distributor" in bt else "No"
    if any(x in dist.lower() for x in ("managed", "service", "integrator")):
        warehouse = "No"
        delivery = "No"
    else:
        delivery = "Yes" if warehouse == "Yes" or ecom == "Yes" else "No"
    return {
        "retail_chains": retail,
        "specialty_grocers": specialty,
        "foodservice": foodservice,
        "ecommerce": ecom,
        "distribution_type": dist,
        "warehouse_logistics": warehouse,
        "delivery_storage": delivery,
    }


def channel_strength(ch: dict[str, str]) -> str:
    score = 0
    for k in ("retail_chains", "specialty_grocers", "foodservice", "ecommerce"):
        if ch.get(k) == "Yes":
            score += 1
    if ch.get("warehouse_logistics") == "Yes":
        score += 1
    return str(min(5, max(1, score)))


def fix_regions(existing: str, hq: str, cities: str) -> str:
    s = (existing or "").strip()
    if s and not looks_like_street(s) and len(s) >= 2:
        return s
    # derive from cities if they look like countries/regions
    c = (cities or "").strip()
    if c and not looks_like_street(c):
        # if cities are city names, still useful as served markets
        countries = COUNTRY_FROM_HQ.findall(c)
        if countries:
            seen = []
            for x in countries:
                pretty = hq_country(x) or x.title()
                if pretty not in seen:
                    seen.append(pretty)
            if seen:
                return ", ".join(seen)
    country = hq_country(hq)
    return country or "Global"


def fix_regional_extensions(existing: str, hq: str, cities: str, regions: str) -> str:
    s = (existing or "").strip()
    if s and not looks_like_street(s):
        # keep clean multi-city / multi-region lists
        return s
    c = (cities or "").strip()
    if c and not looks_like_street(c) and c.lower() != (hq or "").lower():
        return c
    # if regions has multiple places, use as extensions hint
    r = (regions or "").strip()
    if r and "," in r:
        return r
    return ""


def apply_json(mid: str, market: str) -> dict:
    path = ROOT4 / mid / f"{mid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    n = 0
    for r in data.get("relevant_companies") or []:
        company = str(r.get("company") or r.get("brand") or "")
        bt = str(r.get("role") or r.get("channel_type") or r.get("company_function") or "")
        summary = str(r.get("summary") or r.get("company_summary") or r.get("polished_summary") or "")
        crawl = dict(r.get("_crawl") or {})
        website = str(crawl.get("website") or r.get("website") or r.get("domain") or "")
        hq = str(crawl.get("headquarters") or crawl.get("contact_address") or "")
        cities = str(crawl.get("cities_regions") or crawl.get("offices") or "")
        ch = infer_channels(market, company, bt, summary, website)
        ch["channel_strength"] = channel_strength(ch)
        regions = fix_regions(str(crawl.get("regions_served") or ""), hq, cities)
        extensions = fix_regional_extensions(
            str(crawl.get("regional_extensions") or crawl.get("cities_regions") or ""),
            hq,
            cities,
            regions,
        )
        crawl["retail_chains"] = ch["retail_chains"]
        crawl["specialty_grocers"] = ch["specialty_grocers"]
        crawl["foodservice"] = ch["foodservice"]
        crawl["ecommerce"] = ch["ecommerce"]
        crawl["distribution_type"] = ch["distribution_type"]
        crawl["warehouse_logistics"] = ch["warehouse_logistics"]
        crawl["delivery_storage"] = ch["delivery_storage"]
        crawl["channel_strength"] = ch["channel_strength"]
        crawl["regions_served"] = regions
        crawl["regional_extensions"] = extensions
        # also keep offices clean for export Regional Extensions
        if looks_like_street(str(crawl.get("cities_regions") or "")):
            # don't wipe cities if that's all we have for Cities/Regions column
            pass
        r["_crawl"] = crawl
        pe = dict(r.get("_presentation_enrich") or {})
        pe.update(
            {
                "retail_chains": ch["retail_chains"],
                "specialty_grocers": ch["specialty_grocers"],
                "foodservice": ch["foodservice"],
                "ecommerce": ch["ecommerce"],
                "distribution_type": ch["distribution_type"],
                "warehouse_logistics": ch["warehouse_logistics"],
                "delivery_storage": ch["delivery_storage"],
                "channel_strength": ch["channel_strength"],
                "regions_served": regions,
                "regional_extensions": extensions,
            }
        )
        r["_presentation_enrich"] = pe
        n += 1
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"rows": n}


def find_header(ws):
    for r in range(1, 8):
        vals = [c.value for c in ws[r]]
        if any(v and str(v).strip().lower() == "company" for v in vals):
            return r, vals
    return 1, [c.value for c in ws[1]]


def apply_excel(path: Path, market: str, json_path: Path) -> dict:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    by_co = {}
    for r in data.get("relevant_companies") or []:
        co = key(str(r.get("company") or r.get("brand") or ""))
        if co:
            by_co[co] = r.get("_crawl") or {}

    wb = load_workbook(path)
    ws = wb.active
    hr, headers = find_header(ws)
    hmap = {str(h).strip().lower(): i for i, h in enumerate(headers) if h}
    ci = hmap.get("company")
    if ci is None:
        return {"error": "no company"}
    colmap = {
        "retail chains": "retail_chains",
        "specialty grocers": "specialty_grocers",
        "foodservice / horeca": "foodservice",
        "e-commerce": "ecommerce",
        "distribution type": "distribution_type",
        "regions served": "regions_served",
        "regional extensions": "regional_extensions",
        "warehouse / logistics": "warehouse_logistics",
        "delivery / storage": "delivery_storage",
        "channel strength": "channel_strength",
    }
    updated = 0
    for ridx in range(hr + 1, ws.max_row + 1):
        company = ws.cell(ridx, ci + 1).value
        if not company:
            continue
        crawl = by_co.get(key(str(company)))
        if not crawl:
            bt = str(ws.cell(ridx, hmap["business type"] + 1).value or "") if "business type" in hmap else ""
            summary = str(ws.cell(ridx, hmap["summary"] + 1).value or "") if "summary" in hmap else ""
            website = str(ws.cell(ridx, hmap["website"] + 1).value or "") if "website" in hmap else ""
            hq = str(ws.cell(ridx, hmap["headquarters"] + 1).value or "") if "headquarters" in hmap else ""
            cities = (
                str(ws.cell(ridx, hmap["cities / regions"] + 1).value or "") if "cities / regions" in hmap else ""
            )
            ch = infer_channels(market, str(company), bt, summary, website)
            ch["channel_strength"] = channel_strength(ch)
            regions = fix_regions(
                str(ws.cell(ridx, hmap["regions served"] + 1).value or "") if "regions served" in hmap else "",
                hq,
                cities,
            )
            extensions = fix_regional_extensions(
                str(ws.cell(ridx, hmap["regional extensions"] + 1).value or "")
                if "regional extensions" in hmap
                else "",
                hq,
                cities,
                regions,
            )
            crawl = {**ch, "regions_served": regions, "regional_extensions": extensions}
        for excel_col, crawl_key in colmap.items():
            if excel_col not in hmap:
                continue
            val = crawl.get(crawl_key)
            if val is None:
                continue
            ws.cell(ridx, hmap[excel_col] + 1).value = val
            updated += 1
    # Save via temp to avoid partial writes; fall back if destination locked
    tmp = path.with_suffix(path.suffix + ".tmp")
    wb.save(tmp)
    try:
        tmp.replace(path)
        out = str(path)
        locked = False
    except PermissionError:
        alt = path.with_name(path.stem + "_CHANNEL_FIXED.xlsx")
        try:
            if alt.exists():
                alt.unlink()
        except Exception:
            pass
        tmp.replace(alt)
        out = str(alt)
        locked = True
    return {"file": out, "cells": updated, "locked_fallback": locked}


def main() -> None:
    for mid, market in MARKETS.items():
        print(f"\n=== {mid} ({market}) ===")
        print(" JSON", apply_json(mid, market))
        short = re.sub(r"_global$", "", mid)
        jp = ROOT4 / mid / f"{mid}.json"
        for suffix in ["_FINAL.xlsx", "_founded_and_hq_v2.xlsx", "_founded_and_hq_v2_cleaned.xlsx"]:
            xp = ROOT4 / mid / f"{short}{suffix}"
            print(" XLSX", apply_excel(xp, market, jp))
        # mirror
        dst = MIRROR / mid
        dst.mkdir(parents=True, exist_ok=True)
        for pat in [f"{mid}.json", "*_FINAL.xlsx", "*_founded_and_hq_v2.xlsx", "*_founded_and_hq_v2_cleaned.xlsx"]:
            for p in (ROOT4 / mid).glob(pat):
                try:
                    shutil.copy2(p, dst / p.name)
                except PermissionError:
                    print(f"  skip locked {p.name}")


if __name__ == "__main__":
    main()
