#!/usr/bin/env python3
"""Verify/fix ADC channel & logistics columns (pharma-appropriate, not FMCG).

Columns:
  Retail Chains, Specialty Grocers, Foodservice / HoReCa, E-commerce,
  Distribution Type, Regions Served, Regional Extensions,
  Warehouse / Logistics, Delivery / Storage
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fill_adc_contact_role_linkedin as xls

OUT = xls.OUT
XLSX = xls.XLSX
SUMMARY = OUT / "adc_channel_logistics_summary.json"

NA_PHARMA = "Not applicable (pharma / ADC channel)"
NA_B2B = "N/A (B2B hospital / specialty channel)"
NPD = "Not publicly disclosed"

# company -> overrides for multi-country / retail specialty / logistics detail
KNOWN: dict[str, dict[str, str]] = {
    "McKesson Specialty Health": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B provider / specialty pharmacy portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US national specialty pharmacy & oncology network",
        "Warehouse / Logistics": "Yes — GDP cold-chain specialty DCs nationwide",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment; STAT capable",
    },
    "McKesson Specialty Care Distribution": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty ordering portal",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty care distribution network",
        "Warehouse / Logistics": "Yes — GDP cold-chain specialty DCs",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "McKesson Plasma and Biologics": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty ordering portal",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US plasma & biologics specialty network",
        "Warehouse / Logistics": "Yes — cold-chain biologics DCs",
        "Delivery / Storage": "Cold-chain biologics storage & fulfillment",
    },
    "McKesson Medical-Surgical": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B online ordering (McKesson Connect)",
        "Distribution Type": "Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US medical-surgical distribution network",
        "Warehouse / Logistics": "Yes — national medical distribution centers",
        "Delivery / Storage": "Ambient + cold-chain medical product fulfillment",
    },
    "McKesson Canada": {
        "Retail Chains": "Partner pharmacy network (Canada)",
        "E-commerce": "B2B pharmacy / hospital ordering",
        "Distribution Type": "Distributor",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada national pharmaceutical wholesale",
        "Warehouse / Logistics": "Yes — Canadian pharmaceutical DCs",
        "Delivery / Storage": "Temperature-controlled pharma fulfillment",
    },
    "Biologics by McKesson": {
        "Retail Chains": NA_B2B,
        "E-commerce": "Specialty pharmacy patient/provider portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US oncology specialty pharmacy",
        "Warehouse / Logistics": "Yes — oncology specialty cold-chain",
        "Delivery / Storage": "Temperature-controlled oncology specialty fulfillment",
    },
    "Cencora (AmerisourceBergen) Specialty": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty portals (Cencora)",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty distribution & solutions network",
        "Warehouse / Logistics": "Yes — GDP specialty / cold-chain DCs",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "ASD Healthcare": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty ordering",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty physician / oncology channel",
        "Warehouse / Logistics": "Yes — specialty cold-chain DCs",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Besse Medical": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty ordering",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty physician practices",
        "Warehouse / Logistics": "Yes — specialty DCs",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Oncology Supply": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B oncology practice ordering",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US community oncology practices",
        "Warehouse / Logistics": "Yes — oncology specialty DCs",
        "Delivery / Storage": "Temperature-controlled oncology fulfillment",
    },
    "Cardinal Health Specialty Solutions": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty pharmaceutical solutions network",
        "Warehouse / Logistics": "Yes — specialty / cold-chain DCs",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Cardinal Health Specialty Pharmaceutical Distribution": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty ordering",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty pharmaceutical distribution",
        "Warehouse / Logistics": "Yes — specialty / cold-chain DCs",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "CVS Specialty": {
        "Retail Chains": "CVS Pharmacy / specialty retail network",
        "E-commerce": "Patient & provider specialty portals + CVS.com specialty",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US national specialty pharmacy",
        "Warehouse / Logistics": "Yes — specialty pharmacy fulfillment centers",
        "Delivery / Storage": "Temperature-controlled home & clinic delivery",
    },
    "Aetna Specialty Pharmacy": {
        "Retail Chains": "CVS Health / Aetna specialty channel",
        "E-commerce": "Member specialty pharmacy portal",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US payer-aligned specialty pharmacy",
        "Warehouse / Logistics": "Yes — specialty fulfillment (CVS Health)",
        "Delivery / Storage": "Temperature-controlled specialty delivery",
    },
    "Coram CVS Specialty Infusion": {
        "Retail Chains": NA_B2B,
        "E-commerce": "Provider / patient infusion coordination portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty infusion sites & home infusion",
        "Warehouse / Logistics": "Yes — infusion specialty logistics",
        "Delivery / Storage": "Cold-chain infusion drug storage & home delivery",
    },
    "Walgreens Specialty Pharmacy": {
        "Retail Chains": "Walgreens retail specialty pharmacy network",
        "E-commerce": "Specialty patient portal + retail pickup",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US Walgreens specialty pharmacy network",
        "Warehouse / Logistics": "Yes — specialty pharmacy fulfillment",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "AllianceRx Walgreens Pharmacy": {
        "Retail Chains": "Walgreens / AllianceRx specialty network",
        "E-commerce": "Specialty patient/provider portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US specialty pharmacy (AllianceRx)",
        "Warehouse / Logistics": "Yes — specialty fulfillment centers",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Accredo Health Group": {
        "Retail Chains": NA_B2B,
        "E-commerce": "Specialty pharmacy patient/provider portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US national specialty pharmacy (Evernorth)",
        "Warehouse / Logistics": "Yes — specialty cold-chain fulfillment",
        "Delivery / Storage": "Temperature-controlled home & clinic delivery",
    },
    "Optum Specialty Pharmacy": {
        "Retail Chains": NA_B2B,
        "E-commerce": "Specialty pharmacy portals (Optum)",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US Optum specialty pharmacy network",
        "Warehouse / Logistics": "Yes — specialty fulfillment network",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Walmart Specialty Pharmacy": {
        "Retail Chains": "Walmart / Sam's Club specialty pharmacy channel",
        "E-commerce": "Specialty pharmacy portal + store pickup",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US Walmart specialty pharmacy",
        "Warehouse / Logistics": "Yes — specialty pharmacy logistics",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Kroger Specialty Pharmacy": {
        "Retail Chains": "Kroger specialty pharmacy / retail grocery pharmacy",
        "E-commerce": "Specialty portal + retail pickup",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US Kroger specialty pharmacy footprint",
        "Warehouse / Logistics": "Yes — specialty pharmacy logistics",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment",
    },
    "Zuellig Pharma": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B healthcare ordering platforms",
        "Distribution Type": "Distributor",
        "Regions Served": "Asia-Pacific",
        "Regional Extensions": "Singapore HQ; multiple APAC markets",
        "Warehouse / Logistics": "Yes — APAC GDP / cold-chain healthcare DCs",
        "Delivery / Storage": "Temperature-controlled healthcare distribution",
    },
    "Zuellig Pharma Philippines": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B healthcare ordering",
        "Distribution Type": "Distributor",
        "Regions Served": "Philippines",
        "Regional Extensions": "Philippines national healthcare distribution",
        "Warehouse / Logistics": "Yes — PH GDP / cold-chain DCs",
        "Delivery / Storage": "Temperature-controlled pharma fulfillment",
    },
    "DKSH Health": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B market expansion / ordering portals",
        "Distribution Type": "Distributor",
        "Regions Served": "Asia",
        "Regional Extensions": "Multi-country Asia healthcare distribution",
        "Warehouse / Logistics": "Yes — Asia healthcare logistics network",
        "Delivery / Storage": "Temperature-controlled healthcare distribution",
    },
    "Alliance Healthcare UK": {
        "Retail Chains": "Independent / multiple pharmacy partners (UK)",
        "E-commerce": "B2B pharmacy wholesale ordering",
        "Distribution Type": "Distributor",
        "Regions Served": "United Kingdom",
        "Regional Extensions": "UK pharmacy & hospital wholesale network",
        "Warehouse / Logistics": "Yes — UK pharmaceutical wholesale DCs",
        "Delivery / Storage": "Temperature-controlled wholesale fulfillment",
    },
    "AAH Pharmaceuticals": {
        "Retail Chains": "UK pharmacy wholesale customers",
        "E-commerce": "B2B pharmacy ordering",
        "Distribution Type": "Distributor",
        "Regions Served": "United Kingdom",
        "Regional Extensions": "UK national pharmaceutical wholesale",
        "Warehouse / Logistics": "Yes — UK wholesale DCs",
        "Delivery / Storage": "Temperature-controlled wholesale fulfillment",
    },
    "Avantor Inc.": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B e-commerce / VWR catalog ordering",
        "Distribution Type": "Distributor",
        "Regions Served": "Global",
        "Regional Extensions": "Americas; EMEA; APAC lab distribution",
        "Warehouse / Logistics": "Yes — global lab / life-science DCs",
        "Delivery / Storage": "Ambient + cold-chain lab product fulfillment",
    },
    "VWR International": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B online catalog (VWR)",
        "Distribution Type": "Distributor",
        "Regions Served": "Global",
        "Regional Extensions": "Multi-region lab supply network (Avantor)",
        "Warehouse / Logistics": "Yes — lab distribution centers",
        "Delivery / Storage": "Ambient + cold-chain lab fulfillment",
    },
    "Fisher Scientific": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B online catalog (Thermo Fisher)",
        "Distribution Type": "Distributor",
        "Regions Served": "Global",
        "Regional Extensions": "Global lab products distribution",
        "Warehouse / Logistics": "Yes — Thermo Fisher distribution network",
        "Delivery / Storage": "Ambient + cold-chain lab fulfillment",
    },
    "Marken (UPS Healthcare)": {
        "Retail Chains": NA_B2B,
        "E-commerce": "Clinical logistics portals",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "Global",
        "Regional Extensions": "Global clinical supply chain / depot network",
        "Warehouse / Logistics": "Yes — GDP clinical depots & hubs worldwide",
        "Delivery / Storage": "Full temperature-range clinical logistics",
    },
    "Phoenix Pharma SE": {
        "Retail Chains": "European pharmacy wholesale customers",
        "E-commerce": "B2B pharmacy wholesale portals",
        "Distribution Type": "Distributor",
        "Regions Served": "Europe",
        "Regional Extensions": "Multi-country European wholesale (PHOENIX group)",
        "Warehouse / Logistics": "Yes — European pharmaceutical wholesale DCs",
        "Delivery / Storage": "Temperature-controlled wholesale fulfillment",
    },
    "BioCare Specialty Distribution (BioCareSD)": {
        "Retail Chains": NA_B2B,
        "E-commerce": "B2B specialty ordering",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "Cold-chain specialty DCs (Tempe AZ; Olive Branch MS) + forward stocking",
        "Warehouse / Logistics": "Yes — cold-chain specialty DCs (Tempe AZ; Olive Branch MS)",
        "Delivery / Storage": "Temperature-controlled specialty fulfillment; STAT/emergency delivery",
    },
}

COUNTRY_FROM_HQ = re.compile(
    r"(United States|USA|U\.S\.A\.|United Kingdom|Germany|France|Italy|Spain|Portugal|"
    r"Netherlands|Belgium|Switzerland|Austria|Sweden|Norway|Denmark|Finland|Poland|"
    r"Czech|Hungary|Romania|Greece|Ireland|Canada|Mexico|Brazil|Argentina|Chile|"
    r"Colombia|Peru|India|China|Japan|South Korea|Korea|Singapore|Malaysia|Thailand|"
    r"Indonesia|Philippines|Vietnam|Australia|New Zealand|UAE|United Arab Emirates|"
    r"Saudi Arabia|Egypt|Israel|Turkey|South Africa|Nigeria|Pakistan|Taiwan|Hong Kong|"
    r"Portugal|Luxembourg)",
    re.I,
)

RETAIL_SPECIALTY = re.compile(
    r"\b(CVS|Walgreens|Walmart|Kroger|AllianceRx|Shoppers Drug|Jean Coutu|"
    r"CenterWell|Humana Pharmacy|Kaiser)\b",
    re.I,
)


def country_from_text(*texts: str) -> str:
    for t in texts:
        m = COUNTRY_FROM_HQ.search(t or "")
        if not m:
            continue
        c = m.group(1)
        low = c.lower()
        if low in {"usa", "u.s.a."}:
            return "United States"
        if low == "korea":
            return "South Korea"
        if low in {"uae", "united arab emirates"}:
            return "United Arab Emirates"
        return c
    return ""


def cities_clean(row: dict) -> str:
    raw = str(row.get("Cities / Regions") or "").strip()
    if not raw or raw.lower() in {"n/a", "not publicly disclosed", "unknown"}:
        return ""
    # drop non-geo junk
    if re.search(r"distribution|specialty|eu distribution", raw, re.I) and not re.search(
        r"[A-Za-z]{3,}.*,",
        raw,
    ):
        # might still be geo
        pass
    return raw


def fix_row(row: dict) -> dict[str, str]:
    co = row["Company"]
    bt = str(row.get("Business Type") or "Distributor").strip()
    if bt not in {"Distributor", "Specialty Distributor"}:
        bt = "Specialty Distributor" if "specialty" in bt.lower() else "Distributor"
    price = str(row.get("Price Segment") or "")
    hq = str(row.get("Headquarters") or "")
    cities = cities_clean(row)
    core = str(row.get("Core Categories") or "")
    is_research = "Research / Laboratory" in price or "Life science reagents" in core
    is_specialty = bt == "Specialty Distributor" or "Specialty pharmacy" in core

    country = country_from_text(hq, cities, str(row.get("Regions Served") or ""))
    if not country:
        country = str(row.get("Regions Served") or "").split(";")[0].strip() or NPD

    # defaults
    retail = NA_B2B
    if RETAIL_SPECIALTY.search(co) or RETAIL_SPECIALTY.search(str(row.get("Specialty Focus") or "")):
        retail = "Retail specialty pharmacy / pharmacy chain channel"
    elif re.search(r"\b(pharmacy wholesale|independent pharmacy|pharmacy partners)\b", str(row.get("Specialty Focus") or ""), re.I):
        retail = "Pharmacy wholesale customers (not grocery retail)"

    ecommerce = "B2B ordering portal"
    if is_research:
        ecommerce = "B2B online catalog / lab ordering"
    elif is_specialty:
        ecommerce = "Specialty pharmacy / B2B provider portals"
    elif "Mixed specialty & retail" in price:
        ecommerce = "B2B pharmacy / hospital wholesale ordering"

    # Regions Served: primary country or known region
    regions = country
    # Regional Extensions: prefer cities; else descriptive footprint; avoid identical duplicate when cities exist
    if cities and cities.lower() != country.lower() and country.lower() not in cities.lower():
        extensions = cities
    elif cities:
        extensions = cities
    else:
        extensions = f"{country} (primary market)" if country != NPD else NPD

    if is_research:
        warehouse = "Yes — lab / life-science distribution storage (ambient + cold as required)"
        delivery = "Lab product fulfillment; cold-chain for biologics/reagents when required"
    elif is_specialty:
        warehouse = "Yes — GDP / specialty cold-chain capable DCs"
        delivery = "Temperature-controlled specialty fulfillment (2–8°C / frozen as required)"
    else:
        warehouse = "Yes — pharmaceutical wholesale DCs; cold-chain capable"
        delivery = "Temperature-controlled pharmaceutical wholesale fulfillment"

    out = {
        "Retail Chains": retail,
        "Specialty Grocers": NA_PHARMA,
        "Foodservice / HoReCa": NA_PHARMA,
        "E-commerce": ecommerce,
        "Distribution Type": bt,
        "Regions Served": regions,
        "Regional Extensions": extensions,
        "Warehouse / Logistics": warehouse,
        "Delivery / Storage": delivery,
    }

    if co in KNOWN:
        out.update(KNOWN[co])
        # keep grocers/foodservice always N/A
        out["Specialty Grocers"] = NA_PHARMA
        out["Foodservice / HoReCa"] = NA_PHARMA
        if "Distribution Type" not in KNOWN[co]:
            out["Distribution Type"] = bt

    # sanitize bare Yes leftovers if somehow present
    for k in ("Warehouse / Logistics", "Delivery / Storage", "E-commerce"):
        if str(out.get(k) or "").strip().lower() in {"yes", "y", "true"}:
            if k == "E-commerce":
                out[k] = ecommerce
            elif k == "Warehouse / Logistics":
                out[k] = warehouse
            else:
                out[k] = delivery

    return out


def main() -> None:
    bak = XLSX.with_name(XLSX.stem + "_pre_channel_logistics.xlsx")
    if not bak.exists():
        bak.write_bytes(XLSX.read_bytes())
        print(f"Backup: {bak.name}", flush=True)

    rows = xls.load_rows(XLSX)
    changed = 0
    samples = []
    for r in rows:
        before = {c: r.get(c) for c in (
            "Retail Chains", "Specialty Grocers", "Foodservice / HoReCa", "E-commerce",
            "Distribution Type", "Regions Served", "Regional Extensions",
            "Warehouse / Logistics", "Delivery / Storage",
        )}
        fix = fix_row(r)
        for k, v in fix.items():
            r[k] = v
        after = {c: r.get(c) for c in before}
        if before != after:
            changed += 1
        if len(samples) < 10:
            samples.append({"company": r["Company"], **{k: str(v)[:80] for k, v in fix.items()}})

    rows.sort(key=lambda x: x["Company"].lower())
    xls.write_xlsx(XLSX, rows)

    def ctr(col: str) -> list:
        return Counter(str(r.get(col) or "")[:70] for r in rows).most_common(8)

    same_re = sum(
        1
        for r in rows
        if str(r.get("Regions Served") or "").strip() == str(r.get("Regional Extensions") or "").strip()
    )
    bare_yes = sum(
        1
        for r in rows
        for c in ("E-commerce", "Warehouse / Logistics", "Delivery / Storage")
        if str(r.get(c) or "").strip().lower() == "yes"
    )
    grocer_ok = sum(1 for r in rows if "Not applicable" in str(r.get("Specialty Grocers") or ""))
    food_ok = sum(1 for r in rows if "Not applicable" in str(r.get("Foodservice / HoReCa") or ""))
    dt_match = sum(
        1
        for r in rows
        if str(r.get("Distribution Type") or "") == str(r.get("Business Type") or "")
        or r["Company"] in KNOWN
    )

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total": len(rows),
        "rows_changed": changed,
        "retail_chains": ctr("Retail Chains"),
        "specialty_grocers_na": grocer_ok,
        "foodservice_na": food_ok,
        "ecommerce": ctr("E-commerce"),
        "distribution_type": ctr("Distribution Type"),
        "regions_served_top": ctr("Regions Served"),
        "regions_eq_extensions": same_re,
        "warehouse_top": ctr("Warehouse / Logistics"),
        "delivery_top": ctr("Delivery / Storage"),
        "bare_yes_remaining": bare_yes,
        "distribution_type_aligned": dt_match,
        "samples": samples,
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: summary[k] for k in summary if k != "samples"}, indent=2))
    print("\nSamples:")
    for s in samples[:6]:
        print(f"  {s['company'][:40]}")
        for k in (
            "Retail Chains", "Specialty Grocers", "Foodservice / HoReCa", "E-commerce",
            "Distribution Type", "Regions Served", "Regional Extensions",
            "Warehouse / Logistics", "Delivery / Storage",
        ):
            print(f"    {k[:22]:22s}: {s.get(k, '')[:65]}")


if __name__ == "__main__":
    main()
