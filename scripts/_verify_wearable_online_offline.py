# -*- coding: utf-8 -*-
"""
Rename Retail/E-commerce/Both -> Online/Offline/Both and verify fills.
Allowed values: Online | Offline | Both | Not publicly disclosed
- Online: primarily e-commerce / web shop / marketplace (no meaningful store/depot retail)
- Offline: primarily physical stores, depots/van wholesale, field sales (no meaningful online shop)
- Both: clear evidence of BOTH physical channel AND online ordering/shop
Do NOT invent. DeepSeek high/medium only + evidence check.
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
AUDIT = ROOT / "_wearable_online_offline_audit.json"
CKPT = ROOT / "_wearable_online_offline_ckpt.json"
NPD = "Not publicly disclosed"
OLD_COL = "Retail / E-commerce / Both"
NEW_COL = "Online / Offline / Both"
ALLOWED = {"Online", "Offline", "Both", NPD}

load_dotenv(r"D:\vendor-intel\.env")

# Agent web-verified channel modes
HAND: dict[str, str] = {
    # Online-primary e-commerce / marketplace
    "AED Superstore": "Online",
    "Carewell": "Online",
    "DocCheck Shop": "Online",
    "Praxisdienst": "Online",
    "Vitality Medical": "Online",
    "Tata 1mg": "Online",
    "PharmEasy": "Online",
    "Netmeds": "Online",
    "Allegro Medical": "Online",
    "Rehabmart": "Online",
    "Medex Supply": "Online",
    "SpinLife": "Online",
    "IndiaMART": "Online",
    "Medikabazaar": "Online",
    "TradeIndia": "Online",
    "eBay Inc.": "Online",
    "Amazon": "Online",
    "Amazon Pharmacy / Care": "Online",
    "HealthKart": "Online",
    "Moglix": "Online",
    "DocMorris": "Online",
    "Zur Rose Group": "Online",
    "UK Meds": "Online",
    "Medisave UK": "Online",
    "Medshop Australia": "Online",
    "HeartSmart": "Online",
    "AED Brands": "Online",
    "Response Ready": "Online",
    "AED USA": "Online",
    # Retailers with stores + website
    "Best Buy": "Both",
    "Best Buy Health": "Both",
    "Walgreens": "Both",
    "Boots UK": "Both",
    "Watsons": "Both",
    "CVS Health": "Both",
    "Shoppers Drug Mart": "Both",
    "Jean Coutu": "Both",
    "Rexall Pharmacy": "Both",
    "Apollo Pharmacy": "Both",
    "MedPlus Health Services": "Both",
    "Superdrug": "Both",
    "Clicks Group": "Both",
    "Dis-Chem Pharmacies": "Both",
    "London Drugs": "Both",
    "Giant Eagle": "Both",
    "Rite Aid": "Both",
    "BinSina Pharmacy": "Both",
    "Aster Pharmacy": "Both",
    "Al Nahdi Medical Company": "Both",
    "Al-Dawaa Medical Services": "Both",
    "MatsukiyoCocokara & Co.": "Both",
    "Sundrug Co., Ltd.": "Both",
    "Tsuruha Holdings": "Both",
    "Bic Camera Inc.": "Both",
    "Reliance Digital": "Both",
    "Kojima Co., Ltd.": "Both",
    "Droga Raia": "Both",
    "Drogasil": "Both",
    "Extrafarma": "Both",
    "Farmacias Similares": "Both",
    "Farmacias del Dr. Simi": "Both",
    "Farmacias Guadalajara": "Both",
    "Farmacias del Ahorro": "Both",
    "Farmacias Cruz Verde": "Both",
    "Farmacias Ahumada": "Both",
    "Farmacias Benavides": "Both",
    "Farmacias San Pablo": "Both",
    "Farmacias Yza": "Both",
    "Farmacias Carol": "Both",
    "Apotek Hjärtat": "Both",
    "Goodlife Pharmacy": "Both",
    "Wellness Warehouse": "Both",
    "Hyundai Home Shopping": "Online",
    "CJ ENM": "Online",
    "Rakuten Group": "Online",
    # Pharma/medical wholesalers: depots/van + ordering portals
    "AAH Pharmaceuticals": "Both",
    "Alliance Healthcare": "Both",
    "Alliance Healthcare España": "Both",
    "Alliance Healthcare France": "Both",
    "Cardinal Health": "Both",
    "Cardinal Health Canada": "Both",
    "Cencora": "Both",
    "AmerisourceBergen": "Both",
    "McKesson Corporation": "Both",
    "McKesson Canada": "Both",
    "McKesson Medical-Surgical": "Both",
    "Henry Schein": "Both",
    "Owens & Minor Distribution": "Both",
    "Noweda": "Both",
    "OCP Répartition": "Both",
    "Sanacorp Pharmahandel": "Both",
    "PHOENIX Pharmahandel": "Both",
    "Gehe Pharma Handel": "Both",
    "Cofares": "Both",
    "Hefame": "Both",
    "Oriola": "Both",
    "Tamro": "Both",
    "Brocacef": "Both",
    "Nadro": "Both",
    "Toho Holdings Co., Ltd.": "Both",
    "Medipal Holdings Corporation": "Both",
    "Alfresa Holdings Corporation": "Both",
    "Shanghai Pharmaceuticals Distribution": "Both",
    "Bound Tree Medical": "Both",
    "Performance Health": "Both",
    "Grainger": "Both",
    "Fastenal": "Both",
    "VWR (Avantor)": "Both",
    "Fisher Scientific": "Both",
    "As One Corporation": "Both",
    "CDW Corporation": "Both",
    "Edgepark Medical Supplies": "Both",
    "Byram Healthcare": "Both",
    "Advanced Diabetes Supply": "Both",
    "Apria": "Both",
    "AdaptHealth": "Both",
    "Lincare Holdings": "Both",
    "Numotion": "Both",
    "National Seating & Mobility": "Both",
    # Field-sales / depot B2B MedTech distributors (primarily offline unless shop found)
    "Zahrawi Group": "Both",  # regional offices + commercial channel; has web presence for inquiry
    "Getz Healthcare": "Offline",  # sales/marketing distributor, B2B field model
    "LifeHealthcare": "Offline",
    "Transmedic": "Offline",
    "Healthcare 21": "Offline",
    "Insan Bakti": "Offline",
    "Medhold": "Offline",
    "Jackson Allison Medical & Surgical": "Offline",
    "GrafiMedics BeLux": "Offline",
    "OneMed": "Both",
    "AddLife": "Offline",
    "AddLife AB": "Offline",
    "DKSH": "Both",
    "Zuellig Pharma": "Both",
    "Mediq": "Both",
    "Bunzl": "Both",
    "Alloga": "Offline",  # healthcare logistics
    "CEVA Logistics": "Offline",
    "Movianto": "Offline",
    "Buffalo Supply": "Offline",  # federal channel / contract sales
    "B. Braun": "Both",
    "Coloplast": "Both",
    "Arjo": "Both",
    "Ecolab": "Both",
    "Grifols": "Both",
    "Spire Health": "Online",
    "Zephyr Technology": "Both",
    "LGU+": "Both",
    "BOC Healthcare": "Both",
    "Practo": "Online",
}

UA = {"User-Agent": "Mozilla/5.0 (compatible; VendorIntelChannel/1.0)"}

ONLINE_HINTS = re.compile(
    r"\b("
    r"shop\s*now|add\s*to\s*cart|buy\s*online|online\s*store|webshop|web\s*shop|"
    r"e-?commerce|marketplace|order\s*online|online\s*ordering|online\s*pharmacy|"
    r"checkout|shopping\s*cart|ship\s*to\s*home|home\s*delivery\s*order|"
    r"online\s*portal|customer\s*portal|ordering\s*portal|e-?procurement"
    r")\b",
    re.I,
)
OFFLINE_HINTS = re.compile(
    r"\b("
    r"store\s*locator|find\s*a\s*store|our\s*stores|branch(es)?|pharmacy\s*stores|"
    r"distribution\s*cent(?:re|er)s?|depot|depots|warehouses?|"
    r"visit\s*us|walk-?in|retail\s*stores?|showroom|"
    r"sales\s*(?:team|force|representatives?)|field\s*sales|"
    r"account\s*manager|on-?site|van\s*deliver"
    r")\b",
    re.I,
)


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
        "/stores",
        "/store-locator",
        "/locations",
        "/contact",
        "/shop",
        "/en/about",
    ]
    out = []
    for p in paths:
        u = base if p == "" else urljoin(base + "/", p.lstrip("/"))
        if u not in out:
            out.append(u)
    return out[:6]


def heuristic_mode(text: str, dist: str) -> str:
    """Weak heuristic only used when DeepSeek returns low — still conservative."""
    on = bool(ONLINE_HINTS.search(text))
    off = bool(OFFLINE_HINTS.search(text))
    d = (dist or "").lower()
    if on and off:
        return "Both"
    if on and not off:
        # specialty reseller / e-commerce leaning
        if "reseller" in d or "e-commerce" in d or "retail pharmacy" in d:
            return "Online"
        # wholesaler with portal only — without offline hints keep NPD
        return NPD
    if off and not on:
        return "Offline"
    return NPD


def deepseek_mode(company: str, website: str, dist: str, core: str, text: str) -> dict:
    key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    if not key:
        return {"mode": "", "confidence": "none", "reason": "no_key"}
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    system = (
        "Classify how a healthcare CHANNEL company sells/distributes products.\n"
        "Return JSON: {mode: Online|Offline|Both|unknown, confidence:high|medium|low|none, reason:string}.\n"
        "DEFINITIONS:\n"
        "- Online: primarily sells via website, e-commerce shop, or online marketplace; "
        "no meaningful physical retail stores or depot/field wholesale for customers.\n"
        "- Offline: primarily physical — retail stores, wholesale depots with van delivery, "
        "and/or field sales reps; no meaningful customer online shop/ordering portal.\n"
        "- Both: clear evidence of BOTH a physical channel (stores/depots/field sales) "
        "AND online shop or customer ordering portal.\n"
        "- unknown: evidence insufficient.\n"
        "A B2B ordering portal PLUS depots/van delivery = Both.\n"
        "Informational website + contact form only (no shop/portal) + sales team = Offline.\n"
        "NEVER guess. Prefer unknown over Both when unsure."
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
        return {"mode": "", "confidence": "none", "reason": str(e)}


def normalize_mode(raw: str) -> str:
    t = (raw or "").strip()
    if not t:
        return ""
    low = t.lower()
    if low in {"online", "e-commerce", "ecommerce", "e commerce"}:
        return "Online"
    if low in {"offline", "retail", "physical"}:
        return "Offline"
    if low in {"both", "omnichannel", "omni-channel", "hybrid"}:
        return "Both"
    if low in {"unknown", "n/a", "none", "npd", NPD.lower()}:
        return NPD
    return ""


def evidence_supports(mode: str, text: str, dist: str = "") -> bool:
    if mode == NPD or not mode:
        return True
    on = bool(ONLINE_HINTS.search(text))
    off = bool(OFFLINE_HINTS.search(text))
    d = (dist or "").lower()
    tl = text.lower()
    if mode == "Online":
        return on or "shop" in tl or "cart" in tl or "buy" in tl or "order online" in tl
    if mode == "Offline":
        if off:
            return True
        # B2B distributors often lack "store locator" wording but are clearly offline/field
        if re.search(
            r"distribut|wholesal|hospital|clinic|pharmacy customers|sales team|"
            r"account manager|medical device|healthcare provider|depot|warehouse|logistics",
            tl,
            re.I,
        ):
            return True
        if any(x in d for x in ("distributor", "logistics", "wholesale")):
            return len(tl) > 200  # some real site text present
        return False
    if mode == "Both":
        if on and off:
            return True
        if on and re.search(r"distribut|depot|warehouse|deliver|store|branch|pharmacy", tl, re.I):
            return True
        if off and re.search(r"order|portal|online|e-?commerce|shop|cart", tl, re.I):
            return True
        return False
    return False


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
        NEW_COL: 14,
        "Distribution Type": 22,
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


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_ONLINE_OFFLINE_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]

    # Rename column if needed
    renamed = False
    for i, h in enumerate(headers, start=1):
        if h == OLD_COL:
            ws.cell(1, i).value = NEW_COL
            headers[i - 1] = NEW_COL
            renamed = True
            print("renamed column to", NEW_COL, flush=True)
            break
    if NEW_COL not in headers and OLD_COL not in headers:
        raise SystemExit(f"Column not found: {OLD_COL} / {NEW_COL}")

    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    ckpt = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    reset = os.getenv("CHANNEL_RESET", "0") == "1"
    if reset:
        for r in range(2, ws.max_row + 1):
            ws.cell(r, ic[NEW_COL]).value = NPD
        print("RESET all channel modes to NPD", flush=True)

    # Normalize legacy values
    for r in range(2, ws.max_row + 1):
        v = normalize_mode(str(ws.cell(r, ic[NEW_COL]).value or ""))
        if v in ALLOWED:
            ws.cell(r, ic[NEW_COL]).value = v
        elif str(ws.cell(r, ic[NEW_COL]).value or "").strip() in {"E-commerce", "Retail"}:
            ws.cell(r, ic[NEW_COL]).value = normalize_mode(str(ws.cell(r, ic[NEW_COL]).value))
        else:
            cur = str(ws.cell(r, ic[NEW_COL]).value or "").strip()
            if cur not in ALLOWED:
                ws.cell(r, ic[NEW_COL]).value = NPD

    hand_n = 0
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co not in HAND:
            continue
        mode = HAND[co]
        if mode not in ALLOWED:
            continue
        ws.cell(r, ic[NEW_COL]).value = mode
        hand_n += 1

    targets = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in HAND:
            continue
        cur = str(ws.cell(r, ic[NEW_COL]).value or "").strip()
        if cur == NPD or not cur:
            targets.append(r)

    MAX = int(os.getenv("CHANNEL_EXPAND_MAX", "120"))
    print("hand", hand_n, "targets", len(targets), "max", MAX, flush=True)

    updated = []
    processed = 0
    try:
        for r in targets:
            if processed >= MAX:
                break
            co = str(ws.cell(r, ic["Company"]).value or "").strip()
            if co in ckpt.get("done", {}) and ckpt["done"][co].get("mode") in {"Online", "Offline", "Both"}:
                ws.cell(r, ic[NEW_COL]).value = ckpt["done"][co]["mode"]
                print("skip ckpt", co, ckpt["done"][co]["mode"], flush=True)
                continue
            processed += 1
            website = str(ws.cell(r, ic["Website"]).value or "").strip()
            dist = str(ws.cell(r, ic["Distribution Type"]).value or "").strip() if "Distribution Type" in ic else ""
            core = str(ws.cell(r, ic["Core Categories"]).value or "").strip()
            snippets = []
            for u in candidate_urls(website):
                t = fetch(u)
                if len(t) > 150:
                    snippets.append(f"URL:{u}\n{t[:3000]}")
                time.sleep(0.1)
            text = "\n\n---\n\n".join(snippets) if snippets else f"(no site text) Dist={dist}; Core={core}"
            ds = deepseek_mode(co, website, dist, core, text)
            conf = str(ds.get("confidence") or "").lower()
            mode = normalize_mode(str(ds.get("mode") or ""))
            applied = NPD
            if conf in {"high", "medium"} and mode in {"Online", "Offline", "Both"}:
                if conf == "high" or evidence_supports(mode, text, dist):
                    applied = mode
                    print("UPDATED", co, "=>", applied, flush=True)
                else:
                    print("REJECT unsupported", co, mode, flush=True)
            else:
                print("NPD", co, conf, mode or ds.get("mode"), flush=True)

            ws.cell(r, ic[NEW_COL]).value = applied
            updated.append({"company": co, "mode": applied, "conf": conf, "reason": ds.get("reason")})
            ckpt.setdefault("done", {})[co] = {
                "mode": applied,
                "deepseek": ds,
            }
            if processed % 5 == 0:
                CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
                wb.save(XLSX)
                print("checkpoint", processed, flush=True)
    finally:
        # sanity: only allowed values
        for r in range(2, ws.max_row + 1):
            v = str(ws.cell(r, ic[NEW_COL]).value or "").strip()
            if v not in ALLOWED:
                # map leftovers
                mapped = normalize_mode(v) or NPD
                ws.cell(r, ic[NEW_COL]).value = mapped if mapped in ALLOWED else NPD
        apply_format(ws)
        wb.save(XLSX)
        CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    counts = {"Online": 0, "Offline": 0, "Both": 0, NPD: 0}
    samples = []
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, ic[NEW_COL]).value or "").strip()
        counts[v] = counts.get(v, 0) + 1
        if v != NPD and len(samples) < 12:
            samples.append({"company": ws.cell(r, ic["Company"]).value, "mode": v})

    audit = {
        "backup": bak.name,
        "renamed_from": OLD_COL,
        "renamed_to": NEW_COL,
        "hand": hand_n,
        "updated_this_run": updated,
        "processed": processed,
        "counts": counts,
        "samples": samples,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("DONE", counts, "updated", len(updated))


if __name__ == "__main__":
    main()
