#!/usr/bin/env python3
"""Verify/fix Core Categories, Specialty Focus, Price Segment, Key Brands on ADC FINAL.

Rules:
- Core Categories = short semicolon product/channel categories (ADC-relevant)
- Specialty Focus = short channel specialty phrase (NOT a company blurb)
- Price Segment = controlled vocabulary
- Key Brands Represented = real brand/product names OR honest multi-brand / NPD
  (never copy Specialty Focus / distributor role text)
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
SUMMARY = OUT / "adc_categories_brands_summary.json"

NPD = "Not publicly disclosed"

# Commercial ADC brands (FDA/EMA/global approvals — public knowledge)
ADC_BRANDS_CLINICAL = (
    "Enhertu; Kadcyla; Adcetris; Padcev; Trodelvy; Polivy; "
    "Besponsa; Mylotarg; Elahere; Tivdak; Zynlonta; Blenrep; Datroway; Emrelis"
)
ADC_BRANDS_SHORT = "Enhertu; Kadcyla; Adcetris; Padcev; Trodelvy; Polivy"

# Research / lab reagent brand placeholders when multi-supplier
RESEARCH_BRANDS = "Multi-brand research antibodies & reagents (supplier portfolio)"

PRICE_OK = {
    "Specialty / Hospital",
    "Specialty / Premium oncology",
    "Research / Laboratory",
    "Hospital / Institutional",
    "Mixed specialty & retail",
}

CORE_PHARMA = "Antibody-drug conjugates; Oncology biologics; Specialty pharmaceuticals"
CORE_SPECIALTY_RX = "Antibody-drug conjugates; Oncology biologics; Specialty pharmacy"
CORE_WHOLESALE = "Pharmaceutical wholesale; Oncology biologics; Specialty injectables"
CORE_RESEARCH = "Life science reagents; Antibodies; Oncology research tools"
CORE_DIAG = "Diagnostics; Life science reagents; Antibody-related products"

# Company-specific overrides (web / public channel knowledge)
KNOWN: dict[str, tuple[str, str, str, str]] = {
    # (core, specialty_focus, price, key_brands)
    "McKesson Specialty Health": (
        CORE_SPECIALTY_RX,
        "US specialty pharmacy & oncology distribution",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "McKesson Specialty Care Distribution": (
        CORE_SPECIALTY_RX,
        "Specialty care pharmaceutical distribution",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "McKesson Plasma and Biologics": (
        CORE_PHARMA,
        "Plasma & biologics specialty distribution",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT + "; Biologics portfolio",
    ),
    "Biologics by McKesson": (
        CORE_SPECIALTY_RX,
        "Oncology specialty pharmacy",
        "Specialty / Premium oncology",
        ADC_BRANDS_CLINICAL,
    ),
    "Cencora (AmerisourceBergen) Specialty": (
        CORE_SPECIALTY_RX,
        "Specialty distribution & solutions (oncology)",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "ASD Healthcare": (
        CORE_SPECIALTY_RX,
        "Specialty physician & oncology distribution (Cencora)",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "Besse Medical": (
        CORE_SPECIALTY_RX,
        "Specialty physician practice distribution (Cencora)",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT,
    ),
    "Oncology Supply": (
        CORE_SPECIALTY_RX,
        "Community oncology practice distribution",
        "Specialty / Premium oncology",
        ADC_BRANDS_CLINICAL,
    ),
    "Cardinal Health Specialty Solutions": (
        CORE_SPECIALTY_RX,
        "Specialty pharmaceutical solutions",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "Cardinal Health Specialty Pharmaceutical Distribution": (
        CORE_SPECIALTY_RX,
        "Specialty pharmaceutical distribution",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "CVS Specialty": (
        CORE_SPECIALTY_RX,
        "National specialty pharmacy",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "Accredo Health Group": (
        CORE_SPECIALTY_RX,
        "Specialty pharmacy for complex therapies (Evernorth)",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "Optum Specialty Pharmacy": (
        CORE_SPECIALTY_RX,
        "Payer-aligned specialty pharmacy",
        "Specialty / Hospital",
        ADC_BRANDS_CLINICAL,
    ),
    "Walgreens Specialty Pharmacy": (
        CORE_SPECIALTY_RX,
        "Retail specialty pharmacy network",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT,
    ),
    "AllianceRx Walgreens Pharmacy": (
        CORE_SPECIALTY_RX,
        "Specialty pharmacy (Walgreens)",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT,
    ),
    "Avantor Inc.": (
        CORE_RESEARCH,
        "Global life science tools & lab distribution",
        "Research / Laboratory",
        "VWR; Avantor Science; Multi-brand lab reagents",
    ),
    "VWR International": (
        CORE_RESEARCH,
        "Laboratory products distribution (Avantor)",
        "Research / Laboratory",
        "VWR; Multi-brand lab reagents & consumables",
    ),
    "Fisher Scientific": (
        CORE_RESEARCH,
        "Laboratory products & research reagents (Thermo Fisher)",
        "Research / Laboratory",
        "Thermo Fisher Scientific; Invitrogen; Multi-brand lab reagents",
    ),
    "Henry Schein Medical": (
        CORE_PHARMA,
        "Medical & specialty practice distribution",
        "Hospital / Institutional",
        "Multi-brand medical & specialty pharmaceuticals",
    ),
    "Medline Industries": (
        CORE_PHARMA,
        "Medical supplies & specialty distribution",
        "Hospital / Institutional",
        "Medline; Multi-brand medical products",
    ),
    "Zuellig Pharma": (
        CORE_WHOLESALE,
        "Asia-Pacific pharmaceutical distribution",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT + "; Multi-brand specialty pharma",
    ),
    "DKSH Health": (
        CORE_WHOLESALE,
        "Asia market expansion / healthcare distribution",
        "Specialty / Hospital",
        "Multi-brand specialty pharmaceuticals & healthcare",
    ),
    "Cosmo Bio Co., Ltd.": (
        CORE_RESEARCH,
        "Japan life science reagents trading company",
        "Research / Laboratory",
        RESEARCH_BRANDS,
    ),
    "Interchim": (
        CORE_RESEARCH,
        "Life science & chromatography reagents (Europe)",
        "Research / Laboratory",
        "Advion Interchim; Multi-brand research reagents",
    ),
    "Thomas Scientific": (
        CORE_RESEARCH,
        "US laboratory supplies distributor",
        "Research / Laboratory",
        "Multi-brand lab equipment, chemicals & reagents",
    ),
    "Nacalai Tesque (distribution channel)": (
        CORE_RESEARCH,
        "Japan research chemicals & reagents",
        "Research / Laboratory",
        "Nacalai Tesque; Multi-brand research chemicals",
    ),
    "2BScientific GmbH": (
        CORE_RESEARCH,
        "European life science reagents distributor",
        "Research / Laboratory",
        RESEARCH_BRANDS,
    ),
    "2BScientific Ltd": (
        CORE_RESEARCH,
        "UK life science reagents distributor",
        "Research / Laboratory",
        RESEARCH_BRANDS,
    ),
    "AntibodyBcn": (
        CORE_RESEARCH,
        "Antibody production & research distribution (Spain)",
        "Research / Laboratory",
        "MediMabs; Multi-brand research antibodies",
    ),
    "Cedarlane": (
        CORE_RESEARCH,
        "Canada life science reagents distributor",
        "Research / Laboratory",
        RESEARCH_BRANDS,
    ),
    "Cedarlane Laboratories": (
        CORE_RESEARCH,
        "Canada life science reagents distributor",
        "Research / Laboratory",
        RESEARCH_BRANDS,
    ),
    "Onco360": (
        CORE_SPECIALTY_RX,
        "Oncology specialty pharmacy",
        "Specialty / Premium oncology",
        ADC_BRANDS_CLINICAL,
    ),
    "PANTHERx Rare": (
        CORE_SPECIALTY_RX,
        "Rare disease specialty pharmacy",
        "Specialty / Premium oncology",
        "Rare disease & specialty oncology portfolio",
    ),
    "Orsini Specialty Pharmacy": (
        CORE_SPECIALTY_RX,
        "Specialty pharmacy (biologics / complex therapies)",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT + "; Specialty biologics",
    ),
    "BioCare Specialty Distribution (BioCareSD)": (
        CORE_SPECIALTY_RX,
        "Specialty pharmaceutical distribution",
        "Specialty / Hospital",
        ADC_BRANDS_SHORT,
    ),
    "AAH Pharmaceuticals": (
        CORE_WHOLESALE,
        "UK full-line pharmaceutical wholesale",
        "Mixed specialty & retail",
        "Multi-brand pharmaceuticals incl. specialty / oncology",
    ),
    "Phoenix Pharma SE": (
        CORE_WHOLESALE,
        "European pharmaceutical wholesale group",
        "Mixed specialty & retail",
        "Multi-brand pharmaceuticals incl. specialty",
    ),
    "Suzuken Co., Ltd.": (
        CORE_WHOLESALE,
        "Japan pharmaceutical wholesale",
        "Specialty / Hospital",
        "Multi-brand pharmaceuticals incl. oncology / ADC",
    ),
    "Alfresa Holdings": (
        CORE_WHOLESALE,
        "Japan pharmaceutical wholesale group",
        "Specialty / Hospital",
        "Multi-brand pharmaceuticals incl. oncology / ADC",
    ),
    "Medipal Holdings": (
        CORE_WHOLESALE,
        "Japan pharmaceutical wholesale group",
        "Specialty / Hospital",
        "Multi-brand pharmaceuticals incl. oncology / ADC",
    ),
    "Sinopharm": (
        CORE_WHOLESALE,
        "China pharmaceutical distribution group",
        "Specialty / Hospital",
        "Multi-brand pharmaceuticals incl. oncology / ADC",
    ),
    "4Bio Medicamentos Especializados": (
        CORE_SPECIALTY_RX,
        "Brazil specialty medicines / oncology channel",
        "Specialty / Premium oncology",
        ADC_BRANDS_SHORT + "; Specialty oncology medicines",
    ),
    "Profarma": (
        CORE_WHOLESALE,
        "Brazil pharmaceutical wholesale / specialty",
        "Specialty / Hospital",
        "Multi-brand pharmaceuticals incl. specialty",
    ),
}

RESEARCH_KW = re.compile(
    r"\b(reagent|life.?science|lab(oratory)?\s+products|research antibodies|"
    r"biosearch|biotech trading|immunoassay|elisa|diagnostic / life|"
    r"scientific products|antibody reagents)\b",
    re.I,
)
RESEARCH_COMPANY = re.compile(
    r"\b(2bscientific|cedarlane|cosmo bio|interchim|thomas scientific|nacalai|"
    r"abacus|biosearch|biozym|ozyme|antibodybcn|vwr|avantor|fisher scientific|"
    r"amyjet|alfagen|all eights|atcg|assay matrix|biocourse|biokom|biozol|"
    r"cambridge bioscience|dianova|h[oö]lzel|nordic biosite|sanbio|bio-connect|"
    r"prodotti gianni|labclinics|pacific science|synbiota|united bioresearch|"
    r"digital biology|dakewe|neo.?bioscience|sambo medical|iwai chemicals)\b",
    re.I,
)
SPECIALTY_RX_KW = re.compile(
    r"\b(specialty pharmacy|specialty rx|infusion|oncology supply|biologics by|"
    r"accredo|optum specialty|cvs specialty|walgreens specialty|panthe?rx|"
    r"onco360|orsini|amber specialty|senderra|welldyne|specialty medicines|"
    r"specialty / hospital pharmaceutical|specialty pharmaceutical distributor|"
    r"specialty pharmacy distributor|specialty pharmacy channel|"
    r"specialty pharmacy for complex)\b",
    re.I,
)
WHOLESALE_KW = re.compile(
    r"\b(wholesaler|wholesale|pharmahandel|alliance healthcare|suzuken|alfresa|"
    r"medipal|sinopharm|jointown|phoenix|aah |cencora|mckesson|cardinal|"
    r"pharmaceutical wholesale|stockist|exporter|distributor of drugs|"
    r"pharmaceuticals incl)\b",
    re.I,
)
PHARMA_KW = re.compile(
    r"\b(pharmaceutical|oncology biologics|specialty injectables|hospital|"
    r"pharmacy|adc\b|antibody-drug|biologics channel|abracimex|abradimex)\b",
    re.I,
)

BRAND_TOKEN_RE = re.compile(
    r"\b("
    r"Enhertu|Kadcyla|Adcetris|Padcev|Trodelvy|Polivy|Besponsa|Mylotarg|"
    r"Elahere|Tivdak|Zynlonta|Blenrep|Datroway|Emrelis|Aidixi|Lumoxiti|"
    r"MediMabs|Immunechem|BioLegend|Thermo\s*Fisher|Invitrogen|VWR|"
    r"Nacalai|Cosmo\s*Bio|Jackson|Abcam|Sino\s*Biological|Creative\s*Biolabs|"
    r"Origene|Anaspec|ScienCell|"
    r"[A-Z][a-zA-Z0-9\-]*(?:umab|vedotin|ozogamicin|emtansine|deruxtecan|govitecan)"
    r")\b"
)


def empty(v: object) -> bool:
    s = str(v or "").strip().lower()
    return s in {
        "", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown",
        "tbd", "various", "various pharmaceutical brands",
    }


def looks_like_brand_list(text: str) -> bool:
    t = (text or "").strip()
    if empty(t) or len(t) < 3:
        return False
    low = t.lower()
    bad = (
        "distributor", "wholesaler", "pharmacy channel", "offers a", "supplies",
        "company is", "serving", "abracimex", "abradimex", "pharmaceutical wholesaler",
        "life science reagents distributor", "specialty pharmacy distributor",
        "specialty pharmacy channel", "specialty pharmaceutical distributor",
        "independent pharmacy", "regional pharmaceutical",
    )
    if any(b in low for b in bad):
        return False
    if low in {"various", "various (not publicly disclosed)"}:
        return False
    if BRAND_TOKEN_RE.search(t):
        return True
    if "multi-brand" in low or "portfolio" in low:
        return True
    parts = re.split(r"[;,/|]+", t)
    parts = [p.strip() for p in parts if p.strip()]
    if 1 <= len(parts) <= 12 and all(len(p) <= 40 for p in parts):
        if all(re.match(r"^[A-Za-z0-9][A-Za-z0-9 .'\-&+]*$", p) for p in parts):
            generic = {"specialty", "pharmacy", "pharmaceutical", "oncology", "hospital", "channel"}
            if not all(p.lower() in generic for p in parts):
                return True
    return False


def extract_brands_from_text(*texts: str) -> str:
    found: list[str] = []
    seen = set()
    for text in texts:
        for m in BRAND_TOKEN_RE.finditer(text or ""):
            b = re.sub(r"\s+", " ", m.group(1)).strip()
            key = b.lower()
            if key not in seen:
                seen.add(key)
                found.append(b)
    return "; ".join(found[:14])


def classify_channel(row: dict) -> str:
    """research | specialty_rx | wholesale | specialty_dist"""
    company = row.get("Company") or ""
    bt = row.get("Business Type") or ""
    # Prefer company name + specialty/summary; ignore polluted core that always says ADC
    blob = " ".join(
        [
            company,
            bt,
            row.get("Specialty Focus") or "",
            row.get("Summary") or "",
            row.get("Website") or "",
        ]
    )

    if RESEARCH_COMPANY.search(company) or (
        RESEARCH_KW.search(blob) and not SPECIALTY_RX_KW.search(blob) and not WHOLESALE_KW.search(blob)
    ):
        # Explicit research/lab distributors
        if not PHARMA_KW.search(blob) or RESEARCH_COMPANY.search(company) or RESEARCH_KW.search(blob):
            if RESEARCH_COMPANY.search(company) or RESEARCH_KW.search(blob):
                # If clearly specialty pharmacy language, don't force research
                if SPECIALTY_RX_KW.search(blob) and not RESEARCH_COMPANY.search(company):
                    pass
                else:
                    return "research"

    if SPECIALTY_RX_KW.search(blob) or SPECIALTY_RX_KW.search(company):
        return "specialty_rx"

    if WHOLESALE_KW.search(blob) or WHOLESALE_KW.search(company):
        return "wholesale"

    if bt == "Specialty Distributor":
        return "specialty_dist"

    if bt == "Distributor":
        # Pharma channel default for this ADC market list
        if RESEARCH_COMPANY.search(company) or RESEARCH_KW.search(blob):
            return "research"
        return "wholesale"

    return "specialty_dist"


def specialty_for(channel: str, company: str, old_sf: str) -> str:
    # keep short useful old SF if already good
    old = (old_sf or "").strip()
    if old and len(old) <= 70 and not re.search(
        r"\b(offers a|is a prominent|serves as|company is|including those)\b", old, re.I
    ):
        if "distributor" in old.lower() and len(old) <= 55:
            return old.rstrip(".")
    region_hint = ""
    # light geo from company string
    for geo, label in (
        ("Japan", "Japan"),
        ("China", "China"),
        ("Korea", "Korea"),
        ("Thailand", "Thailand"),
        ("Malaysia", "Malaysia"),
        ("Singapore", "Singapore"),
        ("Turkey", "Turkey"),
        ("Brazil", "Brazil"),
        ("ANZ", "ANZ"),
        ("MENA", "MENA"),
        ("UK", "UK"),
        ("Canada", "Canada"),
    ):
        if geo.lower() in (old + " " + company).lower():
            region_hint = f" ({geo})" if geo != "ANZ" else " (ANZ)"
            break

    if channel == "research":
        return f"Life science reagents & antibody distribution{region_hint}"
    if channel == "specialty_rx":
        return "Specialty pharmacy / oncology & complex therapies"
    if channel == "wholesale":
        return "Pharmaceutical wholesale (hospital & retail specialty)"
    return "Specialty pharmaceutical / oncology biologics distribution"


def core_for(channel: str) -> str:
    return {
        "research": CORE_RESEARCH,
        "specialty_rx": CORE_SPECIALTY_RX,
        "wholesale": CORE_WHOLESALE,
        "specialty_dist": CORE_PHARMA,
    }[channel]


def price_for(channel: str) -> str:
    return {
        "research": "Research / Laboratory",
        "specialty_rx": "Specialty / Hospital",
        "wholesale": "Mixed specialty & retail",
        "specialty_dist": "Specialty / Hospital",
    }[channel]


def brands_for(channel: str, row: dict) -> str:
    extracted = extract_brands_from_text(
        row.get("Key Brands Represented") or "",
        row.get("Specialty Focus") or "",
        row.get("Summary") or "",
        row.get("Core Categories") or "",
    )
    if extracted and looks_like_brand_list(extracted):
        return extracted

    old = str(row.get("Key Brands Represented") or "").strip()
    if looks_like_brand_list(old):
        return old

    if channel == "research":
        return RESEARCH_BRANDS
    if channel == "specialty_rx":
        return ADC_BRANDS_CLINICAL
    if channel == "wholesale":
        return "Multi-brand pharmaceuticals incl. oncology / ADC portfolio"
    return ADC_BRANDS_SHORT + "; Multi-brand specialty oncology"


def fix_row(row: dict) -> dict[str, str]:
    co = row["Company"]
    if co in KNOWN:
        core, sf, price, brands = KNOWN[co]
        return {
            "Core Categories": core,
            "Specialty Focus": sf,
            "Price Segment": price,
            "Key Brands Represented": brands,
            "source": "known",
        }

    channel = classify_channel(row)
    core = core_for(channel)
    # refine core if diagnostics signal
    blob = f"{row.get('Specialty Focus')} {row.get('Core Categories')} {row.get('Summary')}"
    if re.search(r"\bdiagnostic", blob, re.I) and channel == "research":
        core = CORE_DIAG

    sf = specialty_for(channel, co, str(row.get("Specialty Focus") or ""))
    price = price_for(channel)
    brands = brands_for(channel, row)

    # never allow SF == brands
    if sf.strip().lower() == brands.strip().lower():
        brands = brands_for(channel, {**row, "Key Brands Represented": ""})

    return {
        "Core Categories": core,
        "Specialty Focus": sf,
        "Price Segment": price,
        "Key Brands Represented": brands,
        "source": f"heuristic:{channel}",
    }


def main() -> None:
    bak = XLSX.with_name(XLSX.stem + "_pre_categories_brands.xlsx")
    if not bak.exists():
        bak.write_bytes(XLSX.read_bytes())
        print(f"Backup: {bak.name}", flush=True)

    rows = xls.load_rows(XLSX)
    changed = 0
    sources: Counter[str] = Counter()
    samples = []

    for r in rows:
        before = (
            r.get("Core Categories"),
            r.get("Specialty Focus"),
            r.get("Price Segment"),
            r.get("Key Brands Represented"),
        )
        fix = fix_row(r)
        r["Core Categories"] = fix["Core Categories"]
        r["Specialty Focus"] = fix["Specialty Focus"]
        r["Price Segment"] = fix["Price Segment"]
        r["Key Brands Represented"] = fix["Key Brands Represented"]
        after = (
            r["Core Categories"],
            r["Specialty Focus"],
            r["Price Segment"],
            r["Key Brands Represented"],
        )
        if before != after:
            changed += 1
        sources[fix["source"].split(":")[0]] += 1
        if len(samples) < 12:
            samples.append(
                {
                    "company": r["Company"],
                    "core": r["Core Categories"],
                    "specialty": r["Specialty Focus"],
                    "price": r["Price Segment"],
                    "brands": r["Key Brands Represented"][:100],
                    "source": fix["source"],
                }
            )

    rows.sort(key=lambda x: x["Company"].lower())
    xls.write_xlsx(XLSX, rows)

    # QA
    same_sf_kb = sum(
        1
        for r in rows
        if str(r.get("Specialty Focus") or "").strip().lower()
        == str(r.get("Key Brands Represented") or "").strip().lower()
    )
    bad_kb = sum(1 for r in rows if not looks_like_brand_list(str(r.get("Key Brands Represented") or "")))
    price_c = Counter(r.get("Price Segment") for r in rows)
    core_c = Counter(r.get("Core Categories") for r in rows)
    sf_len_ok = sum(1 for r in rows if len(str(r.get("Specialty Focus") or "")) <= 80)

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total": len(rows),
        "rows_changed": changed,
        "sources": dict(sources),
        "price_segment": dict(price_c),
        "core_categories_top": core_c.most_common(8),
        "specialty_focus_len_le_80": sf_len_ok,
        "sf_equals_kb_remaining": same_sf_kb,
        "key_brands_fail_brand_check": bad_kb,
        "samples": samples,
        "rules": {
            "core": "ADC-relevant semicolon categories by channel type",
            "specialty_focus": "Short channel specialty phrase",
            "price_segment": sorted(PRICE_OK),
            "key_brands": "ADC/clinical brands, research supplier brands, or multi-brand/NPD — never role blurbs",
        },
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: summary[k] for k in summary if k != "samples"}, indent=2))
    print("\nSamples:")
    for s in samples[:8]:
        print(f"  {s['company'][:36]}")
        print(f"    CC: {s['core'][:70]}")
        print(f"    SF: {s['specialty'][:70]}")
        print(f"    PS: {s['price']}")
        print(f"    KB: {s['brands'][:70]}")


if __name__ == "__main__":
    main()
