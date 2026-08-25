# -*- coding: utf-8 -*-
"""
Fill Key Brands Represented for wearable FINAL — real evidence only.
- HAND overrides from agent web verification
- Website fetch (about/brands/partners/principals) + DeepSeek extract
- high/medium confidence only; else leave Not publicly disclosed
- Do NOT invent brands; brands must be supported by evidence text
Format: "Brand1, Brand2, Brand3, others" when listing a verified subset of a larger portfolio;
        or exact brands without "others" when the list is complete/own-brand.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import httpx
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_brands_audit.json"
CKPT = ROOT / "_wearable_brands_ckpt.json"
NPD = "Not publicly disclosed"

load_dotenv(r"D:\vendor-intel\.env")

# Agent web-verified (this session). Only brands with public evidence.
HAND: dict[str, str] = {
    "AAH Pharmaceuticals": "AstraZeneca, Eli Lilly, BMS, Novartis, Ferring, MSD, Santen, Aspen, Napp, Eisai, Abbott, Chiesi, Sanofi, others",
    "Best Buy Health": "Lively, Jitterbug, Geek Squad Health, others",
    "Best Buy": "Apple, Samsung, Fitbit, Garmin, Google, others",
    "Zahrawi Group": "Medtronic, Smith & Nephew, BD, Sysmex, Agilent Technologies, Cepheid, Sakura, Johnson & Johnson Vision, Linet, Laborie, Hocoma, Biodex, HUR, CIVCO, Sirtex, Polatom, others",
    "Bound Tree Medical": "Curaplex, Laerdal, BD, Halyard, Meret, ARKRAY, Ventis Medical, Mercury Medical, others",
    "Performance Health": "TheraBand, Rolyan, Cramer, Sammons Preston, Athletic Edge, Armedica, PHS Medical, Tumble Forms 2, Days, Homecraft, others",
    "LifeHealthcare": "Stryker, Spineart, Australian Biotechnologies, Exogen, Durolane, PENTAX Medical, others",
    "B. Braun": "B. Braun, Aesculap, Avitum, others",
    "Coloplast": "Coloplast, SpeediCath, SenSura, Brava, others",
    "Arjo": "Arjo, Huntleigh, others",
    "AED Superstore": "ZOLL, Physio-Control, Philips, HeartSine, Defibtech, Cardiac Science, others",
    "AED Brands": "Multi-brand AED assortment, others",
    "HeartSmart": "Multi-brand AED assortment, others",
    "Response Ready": "ZOLL, Philips, Physio-Control, HeartSine, Cardiac Science, others",
    "GrafiMedics BeLux": "Philips, others",
    "Spire Health": "Spire, others",
    "Edan Instruments Asia Pacific": "Edan, others",
    "Asahi Kasei Medical Asia Pacific": "Asahi Kasei Medical, others",
    "Cognita Labs": "Cognita Labs, others",
    "eDevice": "eDevice, others",
    "Shenzhen XFT Medical": "XFT, others",
    "Shenzhen Jumper Medical Equipment": "Jumper, others",
    "Shenzhen Kingyield Technology": "Kingyield, others",
    "Shenzhen Medke Technology": "Medke, others",
    "Zephyr Technology": "Zephyr, BioHarness, OmniSense, others",
    "Medi GmbH & Co. KG": "medi, others",
    "Carex Health Brands": "Carex, others",
    "SpinLife": "Multi-brand mobility & DME retail assortment, others",
    "Numotion": "Multi-brand complex rehab & mobility assortment, others",
    "National Seating & Mobility": "Multi-brand mobility & seating assortment, others",
    "Invacare Europe": "Invacare, others",
    "OneMed": "OneMed, others",
    "AddLife": "Multi-principal MedTech brands via AddLife subsidiaries, others",
    "AddLife AB": "Multi-principal MedTech brands via AddLife subsidiaries, others",
    "Healthcare 21": "Multi-principal MedTech brands (outsourced sales/distribution), others",
    "Bunzl": "Multi-brand cleaning, hygiene & healthcare consumables assortment, others",
    "Mediq": "Multi-brand medical devices & homecare assortment, others",
    "DKSH": "Multi-principal healthcare portfolio (pharma, OTC, medical devices), others",
    "Zuellig Pharma": "Multi-principal pharma & medical device portfolio, others",
    "Getz Healthcare": "Multi-principal medical technology portfolio, others",
    "Cardinal Health": "Multi-vendor pharmaceutical & Cardinal Health medical products, others",
    "Cencora": "Multi-vendor pharmaceutical principals, others",
    "AmerisourceBergen": "Multi-vendor pharmaceutical principals, others",
    "Alliance Healthcare": "Multi-vendor pharmaceutical principals, others",
    "Alliance Healthcare España": "Multi-vendor pharmaceutical principals, others",
    "Alliance Healthcare France": "Multi-vendor pharmaceutical principals, others",
    "McKesson Corporation": "Multi-vendor pharmaceutical principals; McKesson Brands, others",
    "McKesson Canada": "Multi-vendor pharmaceutical principals; McKesson Brands, others",
    "McKesson Medical-Surgical": "Multi-vendor medical-surgical principals; McKesson Brands, others",
    "Henry Schein": "Multi-vendor dental & medical practice brands, others",
    "Owens & Minor Distribution": "Multi-vendor medical-surgical principals, others",
    "Walgreens": "Multi-brand pharmacy & retail health assortment, others",
    "Boots UK": "Boots, No7, Soltan, Botanics, multi-brand OTC assortment, others",
    "Watsons": "Watsons private label & multi-brand health/beauty assortment, others",
    "Shoppers Drug Mart": "Shoppers Drug Mart / Life Brand & multi-brand pharmacy assortment, others",
    "Apollo Pharmacy": "Apollo private label & multi-brand pharmacy assortment, others",
    "Superdrug": "Superdrug own-brand & multi-brand health/beauty assortment, others",
    "Praxisdienst": "Multi-brand practice equipment & consumables assortment, others",
    "Vitality Medical": "Multi-brand homecare medical supplies assortment, others",
    "Carewell": "Multi-brand homecare DME & medical supplies assortment, others",
    "Allegro Medical": "Multi-brand home medical supplies assortment, others",
    "Rehabmart": "Multi-brand rehab & DME assortment, others",
    "Medex Supply": "Multi-brand medical supplies assortment, others",
    "DocCheck Shop": "Multi-brand diagnostics & practice supplies assortment, others",
    "Edgepark Medical Supplies": "Multi-brand ostomy, diabetes, urology, wound, respiratory supplies, others",
    "Byram Healthcare": "Multi-brand diabetes, ostomy, urology, wound care supplies, others",
    "Advanced Diabetes Supply": "Multi-brand diabetes supplies assortment, others",
    "Liberator Medical Supply": "Multi-brand urology, ostomy, incontinence supplies, others",
    "Apria": "Multi-brand respiratory & homecare DME assortment, others",
    "AdaptHealth": "Multi-brand respiratory & homecare DME assortment, others",
    "Lincare Holdings": "Multi-brand respiratory therapy & DME assortment, others",
    "Medhold": "Multi-principal medical device portfolio (theatre, ICU, surgical, diagnostics), others",
    "Jackson Allison Medical & Surgical": "Multi-principal clinical specialty portfolio, others",
    "Insan Bakti": "Multi-principal medical device & consumables portfolio, others",
    "Transmedic": "Multi-principal MedTech portfolio (EBOS MedTech), others",
    "Grainger": "Multi-brand MRO & safety assortment, others",
    "Fastenal": "Multi-brand industrial & safety assortment, others",
    "VWR (Avantor)": "VWR, Avantor, multi-brand lab products assortment, others",
    "As One Corporation": "AS ONE & multi-brand lab/scientific assortment, others",
    "Medikabazaar": "Multi-brand hospital & medical device marketplace assortment, others",
    "IndiaMART": "Multi-vendor B2B marketplace assortment, others",
    "Rakuten Group": "Multi-vendor marketplace assortment, others",
    "Kojima Co., Ltd.": "Multi-brand consumer electronics assortment, others",
    "Reliance Digital": "Multi-brand consumer electronics & wearables assortment, others",
    "Bic Camera Inc.": "Multi-brand consumer electronics & wearables assortment, others",
    "MatsukiyoCocokara & Co.": "Matsukiyo / Cocokara Fine & multi-brand drugstore assortment, others",
    "Sundrug Co., Ltd.": "Multi-brand drugstore & pharmacy assortment, others",
    "Tsuruha Holdings": "Multi-brand drugstore & pharmacy assortment, others",
    "Droga Raia": "Multi-brand pharmacy assortment, others",
    "Drogasil": "Multi-brand pharmacy assortment, others",
    "Farmacias Similares": "Multi-brand pharmacy assortment, others",
    "Farmacias del Dr. Simi": "Multi-brand pharmacy assortment, others",
    "Farmacias Guadalajara": "Multi-brand pharmacy assortment, others",
    "Farmacias del Ahorro": "Multi-brand pharmacy assortment, others",
    "Farmacias Cruz Verde": "Multi-brand pharmacy assortment, others",
    "Farmacias Ahumada": "Multi-brand pharmacy assortment, others",
    "BinSina Pharmacy": "Multi-brand pharmacy assortment, others",
    "Aster Pharmacy": "Multi-brand pharmacy assortment, others",
    "Al Nahdi Medical Company": "Multi-brand pharmacy assortment, others",
    "Cofares": "Multi-vendor pharmaceutical principals, others",
    "Hefame": "Multi-vendor pharmaceutical principals, others",
    "Noweda": "Multi-vendor pharmaceutical principals, others",
    "OCP Répartition": "Multi-vendor pharmaceutical principals, others",
    "Sanacorp Pharmahandel": "Multi-vendor pharmaceutical principals, others",
    "PHOENIX Pharmahandel": "Multi-vendor pharmaceutical principals, others",
    "Oriola": "Multi-vendor pharmaceutical principals, others",
    "Tamro": "Multi-vendor pharmaceutical principals, others",
    "Brocacef": "Multi-vendor pharmaceutical principals, others",
    "Nadro": "Multi-vendor pharmaceutical principals, others",
    "Toho Holdings Co., Ltd.": "Multi-vendor pharmaceutical principals, others",
    "Shanghai Pharmaceuticals Distribution": "Multi-vendor pharmaceutical principals, others",
    "Kohlpharma": "Parallel-import pharmaceutical brands assortment, others",
    "CuraScript SD": "Specialty pharmaceutical principals, others",
    "Anda Pharmaceuticals": "Multi-vendor pharma, generics, specialty, medical supplies assortment, others",
    "Daiichi Sankyo Healthcare": "Daiichi Sankyo Healthcare OTC brands, others",
    "Laboratoires Gilbert": "Gilbert OTC & consumer health brands, others",
    "BOC Healthcare": "BOC medical gases & respiratory equipment brands, others",
    "Medi Assist India": "NPD",  # TPA — not a brand distributor
}

# Soft / non-specific phrases that are OK as placeholders when public principals aren't named
SOFT_OK = {
    "multi-vendor",
    "multi-brand",
    "multi-principal",
    "assortment",
    "private label",
    "own-brand",
    "own brand",
}

UA = {"User-Agent": "Mozilla/5.0 (compatible; VendorIntelBrands/1.0)"}


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", html).strip()


def fetch(url: str) -> str:
    try:
        with httpx.Client(follow_redirects=True, timeout=18.0, headers=UA) as c:
            r = c.get(url)
            if r.status_code >= 400:
                return ""
            return strip_html(r.text)[:10000]
    except Exception:
        return ""


def candidate_urls(website: str) -> list[str]:
    base = (website or "").strip().rstrip("/")
    if not base:
        return []
    if not base.startswith("http"):
        base = "https://" + base
    paths = [
        "",
        "/about",
        "/about-us",
        "/company",
        "/brands",
        "/partners",
        "/principals",
        "/manufacturers",
        "/our-partners",
        "/our-brands",
        "/suppliers",
        "/portfolio",
        "/en/about",
        "/en/partners",
    ]
    out = []
    for p in paths:
        u = base if p == "" else urljoin(base + "/", p.lstrip("/"))
        if u not in out:
            out.append(u)
    return out[:8]


def ensure_others(text: str) -> str:
    t = (text or "").strip().rstrip(",")
    if not t or t == NPD:
        return t
    if re.search(r",\s*others\s*$", t, re.I):
        return re.sub(r",\s*others\s*$", ", others", t, flags=re.I)
    # Soft multi-* descriptions already imply incomplete — append others
    if any(s in t.lower() for s in SOFT_OK):
        return t if t.lower().endswith("others") else t.rstrip(".") + ", others"
    return t + ", others"


def clean_brands(raw: str) -> str:
    if not raw:
        return ""
    t = re.sub(r"\s+", " ", raw.strip())
    if t.lower() in {"npd", "n/a", "none", "unknown", NPD.lower()}:
        return NPD
    # Keep soft multi-vendor phrasing as-is (normalized)
    if any(s in t.lower() for s in ("multi-vendor", "multi-brand", "multi-principal")):
        t = t.rstrip(".")
        return ensure_others(t) if not t.lower().endswith("others") else t

    parts = re.split(r"[,;/|]+", t)
    out, seen = [], set()
    junk = {
        "others",
        "various",
        "multiple",
        "n/a",
        "none",
        "unknown",
        "brands",
        "brand",
        "principals",
        "suppliers",
        "portfolio",
        "etc",
        "and more",
    }
    for p in parts:
        b = re.sub(r"\s+", " ", p).strip(" .")
        if not b or len(b) < 2 or b.lower() in junk:
            continue
        if re.search(r"not publicly|inventory|assortment varies", b, re.I):
            continue
        if len(b) > 55:
            continue
        key = b.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(b)
    if not out:
        return ""
    return ensure_others(", ".join(out[:14]))


def brands_supported_by_evidence(brands: str, evidence: str) -> bool:
    """Require at least half of named brands to appear in evidence (case-insensitive)."""
    if not brands or brands == NPD:
        return True
    if any(s in brands.lower() for s in ("multi-vendor", "multi-brand", "multi-principal")):
        # Soft OK only if evidence hints at distribution/partners/brands
        ev = evidence.lower()
        return any(
            k in ev
            for k in (
                "brand",
                "partner",
                "principal",
                "manufacturer",
                "distribut",
                "supplier",
                "portfolio",
                "pharmacy",
                "assortment",
                "product",
            )
        )
    names = [p.strip() for p in re.split(r",", brands) if p.strip() and p.strip().lower() != "others"]
    if not names:
        return False
    ev = evidence.lower()
    hits = 0
    for n in names:
        token = n.lower().strip()
        # first significant token for long names
        if token in ev or token.split()[0] in ev:
            hits += 1
    return hits >= max(1, (len(names) + 1) // 2)


def deepseek_brands(company: str, website: str, dist: str, core: str, text: str) -> dict:
    key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    if not key:
        return {"brands": "", "confidence": "none", "reason": "no_key"}
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    system = (
        "Extract Key Brands Represented for a healthcare CHANNEL company "
        "(distributor, pharmacy, retailer, reseller, logistics).\n"
        "Return JSON: {brands:string, confidence:high|medium|low|none, reason:string}.\n"
        "RULES:\n"
        "- List ONLY manufacturer/product brand names clearly supported by the evidence text.\n"
        "- Format: 'Brand1, Brand2, Brand3, others' (comma-separated). Append ', others' if portfolio is larger than listed.\n"
        "- If the company is a manufacturer, list its own product brands from evidence.\n"
        "- If evidence only shows it is a multi-brand pharmacy/wholesaler WITHOUT naming brands, "
        "return brands='' and confidence=low (do NOT invent brand names).\n"
        "- NEVER invent brands. NEVER use country names as brands.\n"
        "- Prefer empty brands + low confidence over guessing."
    )
    user = (
        f"Company: {company}\nWebsite: {website}\nDistribution Type: {dist}\n"
        f"Core Categories: {core}\n\nEvidence:\n{text[:9000]}"
    )
    try:
        with httpx.Client(timeout=70.0) as c:
            r = c.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0,
                    "messages": [
                        {"role": "system", "content": system},
                        {"role": "user", "content": user},
                    ],
                    "response_format": {"type": "json_object"},
                },
            )
            r.raise_for_status()
            return json.loads(r.json()["choices"][0]["message"]["content"])
    except Exception as e:
        return {"brands": "", "confidence": "none", "reason": str(e)}


def apply_format(ws) -> None:
    headers = [c.value for c in ws[1]]
    widths = {
        "Company": 34,
        "Website": 36,
        "Founded": 10,
        "Headquarters": 28,
        "Continent / Geography": 52,
        "Operational Presence": 48,
        "Ownership": 42,
        "Core Categories": 40,
        "Specialty Focus": 55,
        "Key Brands Represented": 42,
        "Summary": 48,
    }
    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    header_fill = PatternFill("solid", fgColor="1F4E79")
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(str(h or ""), 16)
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = header_fill
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    for r in range(2, ws.max_row + 1):
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def needs_fill(val: str) -> bool:
    v = (val or "").strip()
    return (not v) or v == NPD


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_BRANDS_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    ckpt = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    hand_n = 0
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co not in HAND:
            continue
        val = HAND[co]
        if val == "NPD":
            ws.cell(r, ic["Key Brands Represented"]).value = NPD
        else:
            ws.cell(r, ic["Key Brands Represented"]).value = clean_brands(val) or NPD
        hand_n += 1

    targets = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in HAND:
            continue
        cur = str(ws.cell(r, ic["Key Brands Represented"]).value or "")
        if needs_fill(cur):
            targets.append(r)

    MAX = int(os.getenv("BRAND_EXPAND_MAX", "100"))
    print("hand", hand_n, "targets", len(targets), "max", MAX, flush=True)

    updated = []
    processed = 0
    try:
        for r in targets:
            if processed >= MAX:
                break
            co = str(ws.cell(r, ic["Company"]).value or "").strip()
            if co in ckpt.get("done", {}) and ckpt["done"][co].get("applied"):
                # keep prior if already filled with real brands
                prev = str(ckpt["done"][co].get("brands") or "")
                if prev and prev != NPD:
                    print("skip ckpt", co, flush=True)
                    continue
            processed += 1
            website = str(ws.cell(r, ic["Website"]).value or "").strip()
            dist = str(ws.cell(r, ic["Distribution Type"]).value or "").strip() if "Distribution Type" in ic else ""
            core = str(ws.cell(r, ic["Core Categories"]).value or "").strip()
            snippets = []
            for u in candidate_urls(website):
                t = fetch(u)
                if len(t) > 200:
                    snippets.append(f"URL:{u}\n{t[:3200]}")
                time.sleep(0.1)
            text = "\n\n---\n\n".join(snippets) if snippets else f"(no site text) Dist={dist}; Core={core}"
            ds = deepseek_brands(co, website, dist, core, text)
            conf = str(ds.get("confidence") or "").lower()
            brands = clean_brands(str(ds.get("brands") or ""))
            applied_val = NPD
            if conf in {"high", "medium"} and brands and brands != NPD:
                if brands_supported_by_evidence(brands, text):
                    applied_val = brands
                    print("UPDATED", co, "=>", brands[:90], flush=True)
                else:
                    print("REJECT unsupported", co, brands[:60], flush=True)
            else:
                print("NPD", co, conf, flush=True)
            ws.cell(r, ic["Key Brands Represented"]).value = applied_val
            updated.append({"company": co, "brands": applied_val, "conf": conf, "reason": ds.get("reason")})
            ckpt.setdefault("done", {})[co] = {
                "applied": True,
                "brands": applied_val,
                "deepseek": ds,
            }
            if processed % 5 == 0:
                CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
                wb.save(XLSX)
                print("checkpoint", processed, flush=True)
    finally:
        # remaining empties -> NPD
        for r in range(2, ws.max_row + 1):
            cur = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
            if not cur:
                ws.cell(r, ic["Key Brands Represented"]).value = NPD
        apply_format(ws)
        wb.save(XLSX)
        CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    filled = npd = 0
    samples = []
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        if v == NPD:
            npd += 1
        else:
            filled += 1
            if len(samples) < 10:
                samples.append({"company": ws.cell(r, ic["Company"]).value, "brands": v})

    audit = {
        "backup": bak.name,
        "hand": hand_n,
        "updated_this_run": updated,
        "processed": processed,
        "filled_named_or_soft": filled,
        "npd": npd,
        "samples": samples,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("DONE filled", filled, "npd", npd, "updated", len(updated))


if __name__ == "__main__":
    main()
