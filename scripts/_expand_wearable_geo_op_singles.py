# -*- coding: utf-8 -*-
"""
Expand Continent/Geography + Operational Presence for single-country rows.
Uses website About/Locations text + DeepSeek extract. Only high-confidence real countries.
Also applies hand-verified overrides (AAH, Adcock, AddLife, etc.).
Does not modify Website/Founded except AAH HQ correction to Coventry when verified.
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
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_geo_op_expand_singles_audit.json"
CKPT = ROOT / "_wearable_geo_op_expand_singles_ckpt.json"

load_dotenv(r"D:\vendor-intel\.env")

COUNTRY_REGION = {
    "United States": "NA",
    "Canada": "NA",
    "Mexico": "NA",
    "Brazil": "LATAM",
    "Chile": "LATAM",
    "Colombia": "LATAM",
    "Peru": "LATAM",
    "Bolivia": "LATAM",
    "Guatemala": "LATAM",
    "Honduras": "LATAM",
    "El Salvador": "LATAM",
    "Dominican Republic": "LATAM",
    "Puerto Rico": "LATAM",
    "United Kingdom": "EU",
    "Ireland": "EU",
    "France": "EU",
    "Germany": "EU",
    "Spain": "EU",
    "Portugal": "EU",
    "Italy": "EU",
    "Netherlands": "EU",
    "Belgium": "EU",
    "Luxembourg": "EU",
    "Switzerland": "EU",
    "Austria": "EU",
    "Sweden": "EU",
    "Norway": "EU",
    "Denmark": "EU",
    "Finland": "EU",
    "Poland": "EU",
    "Czechia": "EU",
    "Romania": "EU",
    "Hungary": "EU",
    "Greece": "EU",
    "Lithuania": "EU",
    "Slovakia": "EU",
    "Croatia": "EU",
    "Bulgaria": "EU",
    "Turkey": "EU",
    "Bulgaria": "EU",
    "Estonia": "EU",
    "Latvia": "EU",
    "Iceland": "EU",
    "South Africa": "AF",
    "Namibia": "AF",
    "Botswana": "AF",
    "Mauritius": "AF",
    "Nigeria": "AF",
    "Kenya": "AF",
    "Ghana": "AF",
    "Zimbabwe": "AF",
    "Egypt": "ME",
    "Saudi Arabia": "ME",
    "United Arab Emirates": "ME",
    "Qatar": "ME",
    "Kuwait": "ME",
    "Bahrain": "ME",
    "Oman": "ME",
    "Jordan": "ME",
    "Iraq": "ME",
    "Japan": "APAC",
    "China": "APAC",
    "Hong Kong": "APAC",
    "Macau": "APAC",
    "Taiwan": "APAC",
    "South Korea": "APAC",
    "India": "APAC",
    "Singapore": "APAC",
    "Malaysia": "APAC",
    "Thailand": "APAC",
    "Indonesia": "APAC",
    "Philippines": "APAC",
    "Vietnam": "APAC",
    "Brunei": "APAC",
    "Cambodia": "APAC",
    "Laos": "APAC",
    "Myanmar": "APAC",
    "Sri Lanka": "APAC",
    "Australia": "OC",
    "New Zealand": "OC",
    "Costa Rica": "LATAM",
    "Argentina": "LATAM",
    "Uruguay": "LATAM",
}
REGION_ORDER = ["NA", "LATAM", "EU", "AF", "ME", "APAC", "OC"]

# Hand-verified from web (this pass). GEO=physical, OP=commercial.
HAND: dict[str, dict[str, list[str]]] = {
    # AAH: MHRA/aah.co.uk — all depots UK; sells UK only.
    # Revelio employee crumbs in US/India are NOT verified offices/warehouses or sales markets → do NOT add.
    "AAH Pharmaceuticals": {
        "geo": ["United Kingdom"],
        "op": ["United Kingdom"],
        "hq": "Coventry, United Kingdom",
        "note": "HQ Sapphire Court Coventry; UK depots only; commercial UK-only. US/India staff counts not treated as ops countries.",
    },
    # Adcock 2025 AFS: Southern Africa + India manufacturing/ops; Namibia subsidiary
    "Adcock Ingram": {
        "geo": ["South Africa", "India", "Namibia"],
        "op": ["South Africa", "India", "Namibia"],
        "note": "AFS: Southern Africa and India; Adcock Ingram Namibia subsidiary",
    },
    # AddLife annual report: presence in 30 countries incl Nordic/W/C/E Europe + China + Australia
    # Named explicitly in report excerpts: Sweden HQ, UK, Ireland, Denmark, China, Australia + Nordics
    "AddLife AB": {
        "geo": [
            "Sweden",
            "Norway",
            "Denmark",
            "Finland",
            "United Kingdom",
            "Ireland",
            "Spain",
            "Portugal",
            "Italy",
            "Germany",
            "Switzerland",
            "Austria",
            "Czechia",
            "Netherlands",
            "Poland",
            "China",
            "Australia",
        ],
        "op": [
            "Sweden",
            "Norway",
            "Denmark",
            "Finland",
            "United Kingdom",
            "Ireland",
            "Spain",
            "Portugal",
            "Italy",
            "Germany",
            "Switzerland",
            "Austria",
            "Czechia",
            "Netherlands",
            "Poland",
            "China",
            "Australia",
        ],
        "note": "AddLife AR 2025 named markets/employee countries; ~30 total — only named countries filled.",
    },
    "Al-Mansour Medical": {
        "geo": ["Iraq", "Saudi Arabia"],
        "op": ["Iraq", "Saudi Arabia"],
        "note": "Prior verified expand; restored after crash.",
    },
    # Cardinal Health: Canada DCs confirmed; GMPD sold US/Canada/Europe/Asia without full country list → GEO/OP named only US+Canada.
    "Cardinal Health": {
        "geo": ["United States", "Canada"],
        "op": ["United States", "Canada"],
        "note": "CAH newsroom: 43 DCs in US+Canada; Europe/Asia sales mentioned without named country list.",
    },
    # B. Braun AR/management report named manufacturing + major employee countries (not inventing all 64).
    "B. Braun": {
        "geo": [
            "Germany",
            "United States",
            "Malaysia",
            "Poland",
            "Spain",
            "Hungary",
            "France",
            "China",
            "Brazil",
            "Vietnam",
            "India",
            "United Kingdom",
            "Switzerland",
            "Romania",
            "Czechia",
            "Slovakia",
            "South Africa",
            "Italy",
            "Japan",
            "Indonesia",
            "Colombia",
            "Thailand",
            "Argentina",
            "Kenya",
            "Dominican Republic",
            "Bulgaria",
        ],
        "op": [
            "Germany",
            "United States",
            "Malaysia",
            "Poland",
            "Spain",
            "Hungary",
            "France",
            "China",
            "Brazil",
            "Vietnam",
            "India",
            "United Kingdom",
            "Switzerland",
            "Romania",
            "Czechia",
            "Slovakia",
            "South Africa",
            "Italy",
            "Japan",
            "Indonesia",
            "Colombia",
            "Thailand",
            "Argentina",
            "Kenya",
            "Dominican Republic",
            "Bulgaria",
        ],
        "note": "Named plants/sites + reporting countries from B.Braun reports; not inventing remaining of 64.",
    },
    # DKSH locations page / brochure named markets (subset in COUNTRY_REGION).
    "DKSH": {
        "geo": [
            "Switzerland",
            "Australia",
            "Japan",
            "South Korea",
            "Brunei",
            "Cambodia",
            "China",
            "Hong Kong",
            "Macau",
            "Malaysia",
            "Myanmar",
            "New Zealand",
            "India",
            "Indonesia",
            "Philippines",
            "Singapore",
            "Taiwan",
            "Thailand",
            "Vietnam",
            "Laos",
            "Sri Lanka",
            "France",
            "Germany",
            "Netherlands",
            "Poland",
            "United States",
            "Belgium",
            "Denmark",
            "Estonia",
            "Finland",
            "Italy",
            "Latvia",
            "Lithuania",
            "Luxembourg",
            "Norway",
            "Portugal",
            "Spain",
            "Sweden",
            "United Kingdom",
            "Canada",
            "Iceland",
        ],
        "op": [
            "Switzerland",
            "Australia",
            "Japan",
            "South Korea",
            "Brunei",
            "Cambodia",
            "China",
            "Hong Kong",
            "Macau",
            "Malaysia",
            "Myanmar",
            "New Zealand",
            "India",
            "Indonesia",
            "Philippines",
            "Singapore",
            "Taiwan",
            "Thailand",
            "Vietnam",
            "Laos",
            "Sri Lanka",
            "France",
            "Germany",
            "Netherlands",
            "Poland",
            "United States",
            "Belgium",
            "Denmark",
            "Estonia",
            "Finland",
            "Italy",
            "Latvia",
            "Lithuania",
            "Luxembourg",
            "Norway",
            "Portugal",
            "Spain",
            "Sweden",
            "United Kingdom",
            "Canada",
            "Iceland",
        ],
        "note": "DKSH Our Locations named markets only.",
    },
    # Arjo: plants Sweden HQ + Poland/China/UK/Canada/Dominican Republic; largest markets named.
    "Arjo": {
        "geo": [
            "Sweden",
            "Poland",
            "China",
            "United Kingdom",
            "Canada",
            "Dominican Republic",
            "United States",
        ],
        "op": [
            "United States",
            "United Kingdom",
            "France",
            "Canada",
            "Germany",
            "Sweden",
            "Poland",
            "China",
            "Dominican Republic",
            "Australia",
            "Austria",
            "Belgium",
            "Czechia",
            "Denmark",
            "Finland",
            "Ireland",
            "Italy",
            "Netherlands",
            "Norway",
            "Portugal",
            "Spain",
            "Switzerland",
            "Japan",
            "India",
            "New Zealand",
            "Brazil",
            "South Africa",
        ],
        "note": "Arjo plants + named largest markets + country-selector markets (own sales sites).",
    },
    # Coloplast: named production sites; sales subsidiaries 40+ without full list → GEO plants; OP plants + named key markets.
    "Coloplast": {
        "geo": [
            "Denmark",
            "Hungary",
            "China",
            "Costa Rica",
            "United States",
            "France",
            "Germany",
            "Sweden",
            "Iceland",
            "Portugal",
        ],
        "op": [
            "Denmark",
            "Hungary",
            "China",
            "Costa Rica",
            "United States",
            "France",
            "Germany",
            "Sweden",
            "Iceland",
            "Portugal",
            "Canada",
            "Japan",
            "Australia",
            "New Zealand",
        ],
        "note": "Named production network + AR named markets USA/Canada/Japan/Australia/NZ; Portugal site under construction counted as physical footprint.",
    },
    # Cencora / AmerisourceBergen: US + Canada + Alliance Healthcare Europe countries already verified in sheet.
    "Cencora": {
        "geo": [
            "United States",
            "Canada",
            "United Kingdom",
            "France",
            "Germany",
            "Spain",
            "Netherlands",
            "Czechia",
            "Norway",
            "Romania",
            "Turkey",
            "Lithuania",
        ],
        "op": [
            "United States",
            "Canada",
            "United Kingdom",
            "France",
            "Germany",
            "Spain",
            "Netherlands",
            "Czechia",
            "Norway",
            "Romania",
            "Turkey",
            "Lithuania",
        ],
        "note": "US/Canada + Alliance Healthcare European wholesale countries (owned); not inventing unnamed of 50+.",
    },
    "AmerisourceBergen": {
        "geo": [
            "United States",
            "Canada",
            "United Kingdom",
            "France",
            "Germany",
            "Spain",
            "Netherlands",
            "Czechia",
            "Norway",
            "Romania",
            "Turkey",
            "Lithuania",
        ],
        "op": [
            "United States",
            "Canada",
            "United Kingdom",
            "France",
            "Germany",
            "Spain",
            "Netherlands",
            "Czechia",
            "Norway",
            "Romania",
            "Turkey",
            "Lithuania",
        ],
        "note": "Same as Cencora (rebrand); Alliance Healthcare Europe countries.",
    },
    # Grifols: named industrial/plasma hubs.
    "Grifols": {
        "geo": [
            "Spain",
            "United States",
            "Germany",
            "Ireland",
            "Switzerland",
            "Canada",
            "Australia",
            "China",
            "Hungary",
            "Czechia",
            "Austria",
            "Egypt",
        ],
        "op": [
            "Spain",
            "United States",
            "Germany",
            "Ireland",
            "Switzerland",
            "Canada",
            "Australia",
            "China",
            "Hungary",
            "Czechia",
            "Austria",
            "Egypt",
        ],
        "note": "Named hubs/plants/plasma networks from Grifols IAR/press; not inventing all 30+ subsidiaries.",
    },
}

UA = {"User-Agent": "Mozilla/5.0 (compatible; VendorIntelGeo/1.1)"}


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
    paths = ["", "/about", "/about-us", "/company", "/locations", "/our-locations", "/contact", "/contact-us", "/en/about"]
    out = []
    for p in paths:
        u = base if p == "" else urljoin(base + "/", p.lstrip("/"))
        if u not in out:
            out.append(u)
    return out[:6]


def deepseek_extract(company: str, website: str, hq: str, text: str) -> dict:
    key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    if not key:
        return {"physical_countries": [], "commercial_countries": [], "confidence": "none", "reason": "no_key"}
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    system = (
        "Extract country lists for a company from evidence text ONLY. Never invent.\n"
        "physical_countries = countries with HQ/offices/warehouses/plants/stores (physical infrastructure).\n"
        "commercial_countries = countries where the company sells/serves customers.\n"
        "Employee remote headcount alone is NOT enough for physical_countries unless an office/facility is stated.\n"
        "Vague phrases like 'worldwide' or 'global' without named countries → empty lists.\n"
        "Return JSON: {physical_countries:[], commercial_countries:[], confidence:high|medium|low|none, reason:string}."
    )
    user = (
        f"Company: {company}\nWebsite: {website}\nCurrent HQ cell: {hq}\n\nEvidence:\n{text[:8500]}"
    )
    try:
        with httpx.Client(timeout=60.0) as c:
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
        return {
            "physical_countries": [],
            "commercial_countries": [],
            "confidence": "none",
            "reason": str(e),
        }


def canon(c: str) -> str | None:
    t = (c or "").strip()
    if not t:
        return None
    low = t.lower()
    alias = {
        "usa": "United States",
        "us": "United States",
        "uk": "United Kingdom",
        "u.k.": "United Kingdom",
        "uae": "United Arab Emirates",
        "korea": "South Korea",
        "czech republic": "Czechia",
        "hong kong sar": "Hong Kong",
        "england": "United Kingdom",
        "scotland": "United Kingdom",
        "wales": "United Kingdom",
        "northern ireland": "United Kingdom",
    }
    if low in alias:
        return alias[low]
    specials = {
        "united states": "United States",
        "united kingdom": "United Kingdom",
        "united arab emirates": "United Arab Emirates",
        "south africa": "South Africa",
        "south korea": "South Korea",
        "new zealand": "New Zealand",
        "hong kong": "Hong Kong",
        "saudi arabia": "Saudi Arabia",
        "dominican republic": "Dominican Republic",
        "el salvador": "El Salvador",
    }
    if low in specials:
        return specials[low]
    if t in COUNTRY_REGION:
        return t
    # Title case fallback if looks like country word(s)
    if re.fullmatch(r"[A-Za-z][A-Za-z .'-]{1,40}", t):
        return t
    return None


def clean_list(vals) -> list[str]:
    out = []
    seen = set()
    for v in vals or []:
        c = canon(str(v))
        if c and c in COUNTRY_REGION and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def country_from_hq(hq: str) -> str | None:
    if not hq:
        return None
    if "," in hq:
        return canon(hq.split(",")[-1].strip())
    return canon(hq)


def format_geo(countries: list[str]) -> str:
    groups: dict[str, list[str]] = {}
    for c in countries:
        region = COUNTRY_REGION.get(c, "APAC")
        if c in {
            "Saudi Arabia",
            "United Arab Emirates",
            "Qatar",
            "Kuwait",
            "Bahrain",
            "Oman",
            "Jordan",
            "Iraq",
        }:
            region = "ME"
        groups.setdefault(region, [])
        if c not in groups[region]:
            groups[region].append(c)
    if "EU" in groups and "Turkey" in groups["EU"] and "ME" in groups:
        groups["EU"].remove("Turkey")
        if not groups["EU"]:
            del groups["EU"]
        if "Turkey" not in groups.setdefault("ME", []):
            groups["ME"].append("Turkey")
    parts = []
    for region in REGION_ORDER:
        if region in groups and groups[region]:
            parts.append(f"{region}({', '.join(groups[region])})")
    return "; ".join(parts)


def format_op(countries: list[str]) -> str:
    return "; ".join(countries)


def apply_format(ws) -> None:
    headers = [c.value for c in ws[1]]
    widths = {
        "Company": 34,
        "Website": 36,
        "Founded": 10,
        "Headquarters": 28,
        "Continent / Geography": 52,
        "Operational Presence": 48,
        "Ownership": 18,
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
    ws.row_dimensions[1].height = 32
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = header_fill
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    geo_i = headers.index("Continent / Geography") + 1
    op_i = headers.index("Operational Presence") + 1
    for r in range(2, ws.max_row + 1):
        longest = max(len(str(ws.cell(r, geo_i).value or "")), len(str(ws.cell(r, op_i).value or "")), 20)
        ws.row_dimensions[r].height = 15 * max(2, min(8, longest // 48 + 1))
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_EXPAND_SINGLES_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    ckpt = {"done": {}, "results": []}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    # 1) Hand-verified overrides first
    hand_applied = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co not in HAND:
            continue
        info = HAND[co]
        if info.get("hq"):
            ws.cell(r, ic["Headquarters"]).value = info["hq"]
        geo = clean_list(info["geo"])
        op = clean_list(info["op"])
        ws.cell(r, ic["Continent / Geography"]).value = format_geo(geo)
        ws.cell(r, ic["Operational Presence"]).value = format_op(op)
        hand_applied.append({"company": co, "geo": format_geo(geo), "op": format_op(op), "note": info.get("note")})

    # 2) DeepSeek expand for other single-country rows (limit batch for runtime; checkpoint)
    # Process ALL singles but skip if already multi after hand, or already in ckpt with accepted multi
    targets = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in HAND:
            continue
        op = str(ws.cell(r, ic["Operational Presence"]).value or "")
        geo = str(ws.cell(r, ic["Continent / Geography"]).value or "")
        # still single if no comma inside CODE(...) and no multi OP
        single_op = ";" not in op
        single_geo = "," not in geo
        if single_op and single_geo:
            targets.append(r)

    print("hand_applied", len(hand_applied), "deepseek_targets", len(targets), flush=True)

    expanded = []
    # Cap NEW processed companies per run (skip already-checkpointed).
    MAX = int(os.getenv("GEO_EXPAND_MAX", "120"))
    processed_new = 0
    try:
        for idx, r in enumerate(targets, start=1):
            if processed_new >= MAX:
                break
            co = str(ws.cell(r, ic["Company"]).value or "").strip()
            website = str(ws.cell(r, ic["Website"]).value or "").strip()
            hq = str(ws.cell(r, ic["Headquarters"]).value or "").strip()
            if co in ckpt.get("done", {}):
                # Already verified this run/session (applied or intentionally kept single)
                print("skip ckpt", co, flush=True)
                continue

            processed_new += 1
            snippets = []
            for u in candidate_urls(website):
                t = fetch(u)
                if len(t) > 250:
                    snippets.append(f"URL:{u}\n{t[:3500]}")
                time.sleep(0.2)
            text = "\n\n---\n\n".join(snippets) if snippets else f"(no site text) HQ={hq}"
            ds = deepseek_extract(co, website, hq, text)
            conf = str(ds.get("confidence") or "").lower()
            phys = clean_list(ds.get("physical_countries"))
            comm = clean_list(ds.get("commercial_countries"))
            hq_c = country_from_hq(hq)
            if hq_c:
                if hq_c not in phys:
                    phys = [hq_c] + phys
                if hq_c not in comm:
                    comm = [hq_c] + [x for x in comm if x != hq_c]

            applied = False
            # Accept only high/medium with at least one EXTRA country beyond HQ
            extra_phys = [c for c in phys if c != hq_c]
            extra_comm = [c for c in comm if c != hq_c]
            if conf in {"high", "medium"} and (extra_phys or extra_comm):
                # Reject ghost GEO extras when commercial list has no extras
                if extra_phys and not extra_comm:
                    print("reject_inconsistent", co, "phys", extra_phys, "comm", extra_comm, flush=True)
                else:
                    new_geo = format_geo(phys if phys else ([hq_c] if hq_c else []))
                    new_op = format_op(comm if comm else ([hq_c] if hq_c else []))
                    ws.cell(r, ic["Continent / Geography"]).value = new_geo
                    ws.cell(r, ic["Operational Presence"]).value = new_op
                    applied = True
                    expanded.append(
                        {
                            "company": co,
                            "geo": new_geo,
                            "op": new_op,
                            "deepseek": ds,
                            "conf": conf,
                        }
                    )
                    print("EXPANDED", co, "=>", new_geo, "|", new_op, flush=True)
            if not applied:
                print("keep", co, conf, "extras", extra_phys, extra_comm, flush=True)

            ckpt.setdefault("done", {})[co] = {
                "applied": applied,
                "deepseek": ds,
                "phys": phys,
                "comm": comm,
            }
            if processed_new % 5 == 0:
                CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
                wb.save(XLSX)
                print("checkpoint", processed_new, flush=True)
    finally:
        apply_format(ws)
        wb.save(XLSX)
        CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    # recount
    single = multi = 0
    for r in range(2, ws.max_row + 1):
        op = str(ws.cell(r, ic["Operational Presence"]).value or "")
        if ";" in op or ("," in str(ws.cell(r, ic["Continent / Geography"]).value or "")):
            multi += 1
        else:
            single += 1

    audit = {
        "backup": bak.name,
        "aah_note": HAND["AAH Pharmaceuticals"]["note"],
        "hand_applied": hand_applied,
        "deepseek_expanded": expanded,
        "deepseek_targets_remaining_estimate": max(0, len(targets) - MAX),
        "counts_after": {"single_ish": single, "multi_ish": multi},
        "processed_this_run": min(MAX, len(targets)),
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("DONE multi_ish", multi, "single_ish", single, "expanded", len(expanded))
    for h in hand_applied:
        print("HAND", h["company"], h["geo"], h["op"])


if __name__ == "__main__":
    main()
