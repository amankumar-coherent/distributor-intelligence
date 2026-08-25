# -*- coding: utf-8 -*-
"""Fill remaining 13 country-only HQs via web snippets + DeepSeek extract. Never invent."""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import httpx
from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_hq_last13_audit.json"

# Agent web-verified (sources in audit notes). Prefer these over model guess.
AGENT_VERIFIED: dict[str, tuple[str, str]] = {
    # Primary Care Supplies contact page: Unit 6 Mill Hall Business Estate, Aylesford Kent
    "Primary Care Supplies": (
        "Aylesford, United Kingdom",
        "primarycaresupplies.co.uk/contact-us",
    ),
    # BioD Medica EMIS/LinkedIn: Subang Jaya, Selangor
    "BioD Medica": ("Subang Jaya, Malaysia", "emis + linkedin Subang Jaya HQ"),
    # MyBanzu site footer: West Puchong + East Miri; company no. listed with Miri/Puchong — use Puchong (West HQ address first on site)
    "MyBanzu Healthcare": (
        "Puchong, Malaysia",
        "mybanzu.com West Malaysia Address Jalan Metro Puchong",
    ),
    # Farmacias Kielsa-style La Paz: Inversiones La Paz Tegucigalpa (related group) — verify via DeepSeek+site first
}

TARGETS = [
    "Al-Watania Medical",
    "Farmacias Carol",
    "Farmacias La Paz",
    "Farmacias Siman",
    "ProCare Pharmacy",
    "Tecnomed Brasil",
    "BioD Medica",
    "MyBanzu Healthcare",
    "Primary Care Supplies",
    "Provita Medical",
    "Veloce Solutions",
    "AED USA",
    "EMS Supply Co.",
]

UA = {
    "User-Agent": "Mozilla/5.0 (compatible; VendorIntelHQ/1.0; +research)",
    "Accept": "text/html,application/xhtml+xml",
}


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    html = re.sub(r"\s+", " ", html)
    return html.strip()


def fetch_text(url: str, timeout: float = 20.0) -> str:
    if not url or not url.startswith("http"):
        return ""
    try:
        with httpx.Client(follow_redirects=True, timeout=timeout, headers=UA) as c:
            r = c.get(url)
            if r.status_code >= 400:
                return ""
            return strip_html(r.text)[:12000]
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
        "/",
        "/contact",
        "/contact-us",
        "/contactus",
        "/about",
        "/about-us",
        "/en/contact",
        "/empresa",
        "/contato",
        "/contacto",
        "/impressum",
        "/locations",
        "/our-locations",
    ]
    out = []
    for p in paths:
        out.append(urljoin(base + "/", p.lstrip("/")) if p not in ("", "/") else base + ("" if p == "" else "/"))
    # dedupe
    seen = set()
    uniq = []
    for u in out:
        if u not in seen:
            seen.add(u)
            uniq.append(u)
    return uniq[:10]


def deepseek_extract(company: str, website: str, country_hint: str, snippets: str) -> dict:
    load_dotenv(r"D:\vendor-intel\.env")
    key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    if not key:
        return {"hq": "", "confidence": "none", "reason": "no_api_key"}
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    system = (
        "You extract company headquarters city+country from provided website text only. "
        "Never invent. If city is not clearly stated, return hq empty. "
        "Format HQ as 'City, Country' or for US 'City, ST, United States'. "
        "Singapore city-state: 'Singapore, Singapore'. "
        "Respond JSON only: {\"hq\":\"...\",\"confidence\":\"high|medium|low|none\",\"reason\":\"short\"}."
    )
    user = (
        f"Company: {company}\nWebsite: {website}\nCountry hint in sheet: {country_hint}\n\n"
        f"Website text excerpts:\n{snippets[:9000]}\n\n"
        "Extract headquarters city and country only if explicitly supported by the text."
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
            content = r.json()["choices"][0]["message"]["content"]
            data = json.loads(content)
            return {
                "hq": str(data.get("hq") or "").strip(),
                "confidence": str(data.get("confidence") or "").strip().lower(),
                "reason": str(data.get("reason") or "").strip(),
            }
    except Exception as e:
        return {"hq": "", "confidence": "none", "reason": f"api_error:{e}"}


def looks_like_city_country(hq: str) -> bool:
    s = (hq or "").strip()
    if "," not in s:
        return False
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) < 2:
        return False
    # reject if first token is only a country
    bad_first = {
        "united states",
        "usa",
        "united kingdom",
        "uk",
        "saudi arabia",
        "malaysia",
        "brazil",
        "honduras",
        "el salvador",
        "puerto rico",
        "europe",
        "belgium",
        "germany",
        "singapore",  # alone as first of "Singapore, Malaysia" ok; "Singapore, Singapore" handled below
    }
    low = s.lower()
    if low in {"singapore, singapore", "hong kong sar, china", "hong kong, china"}:
        return True
    if parts[0].lower() in bad_first and low != "singapore, singapore":
        return False
    # reject emails/phones invented patterns
    if "@" in s or re.search(r"\d{3}[-.\s]\d{3}", s):
        return False
    return True


def main() -> None:
    load_dotenv(r"D:\vendor-intel\.env")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_HQ_LAST13_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}
    co_i, hq_i, web_i = ic["Company"], ic["Headquarters"], ic["Website"]

    rows = {}
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, co_i).value or "").strip()
        if co in TARGETS:
            rows[co] = {
                "row": r,
                "hq": str(ws.cell(r, hq_i).value or "").strip(),
                "website": str(ws.cell(r, web_i).value or "").strip(),
            }

    # Extra agent web evidence snippets (from this session's searches)
    EXTRA_SNIPPETS = {
        "Primary Care Supplies": (
            "Contact Us Visit Us: Primary Care Supplies Unit 6 Mill Hall Business Estate "
            "Aylesford Kent ME20 7JZ United Kingdom. LinkedIn HQ Aylesford Kent."
        ),
        "BioD Medica": (
            "Biod Medica Sdn. Bhd. headquarters B-06-38A Block B, Sunway Geo Jalan Lagoon Selatan, "
            "Bandar Sunway Subang Jaya Selangor 47500 Malaysia. LinkedIn Headquarters: Subang Jaya, Malaysia."
        ),
        "MyBanzu Healthcare": (
            "MyBanzu Healthcare Sdn Bhd website: West Malaysia Address 23, 1st Floor, Jalan Metro 1/1 "
            "47160 Puchong, Selangor Malaysia. East Malaysia Address Lot 1431-03 Jalan Krokop Lorong 5 "
            "Business Centre 98007 Miri Sarawak Malaysia."
        ),
        "Farmacias Carol": (
            "Grupo Carol / Farmacia Carol LinkedIn and Datanyze: Headquarters Santo Domingo, Dominican Republic. "
            "Note: sheet listed Puerto Rico; official HQ sources say Santo Domingo DR. "
            "Website farmaciacarol.com / farmaciascarol.com."
        ),
        "Farmacias Siman": (
            "CB Insights Farmacia Simán headquarters 6 Ave 5 Calle So San Pedro Sula Honduras. "
            "Sheet lists El Salvador; primary HQ sources say San Pedro Sula Honduras for Farmacias Simán."
        ),
        "Farmacias La Paz": (
            "Inversiones La Paz (operator context) LinkedIn HQ Tegucigalpa Honduras. "
            "Confirm if same as Farmacias La Paz chain on farmaciaslapaz.com."
        ),
        "Al-Watania Medical": (
            "Possible related: Al Wattan Medical Group contact lists Riyadh Saudi Arabia. "
            "Verify domain alwatania.com is same entity before applying."
        ),
        "Provita Medical": (
            "provita.de imprint: provita medical gmbh & co. kg Auf der Huhfuhr 8 42929 Wermelskirchen Germany. "
            "Sheet website provita.com / United States — may be US sales arm; do not invent US city if not on site."
        ),
    }

    results = []
    updated = []
    for co in TARGETS:
        info = rows.get(co)
        if not info:
            results.append({"company": co, "status": "missing_row"})
            continue
        website = info["website"]
        old_hq = info["hq"]
        snippets_parts = []
        if co in EXTRA_SNIPPETS:
            snippets_parts.append(EXTRA_SNIPPETS[co])
        for u in candidate_urls(website):
            txt = fetch_text(u)
            if len(txt) > 200:
                snippets_parts.append(f"URL: {u}\n{txt[:3500]}")
            time.sleep(0.3)
        snippets = "\n\n---\n\n".join(snippets_parts)[:11000]
        ds = deepseek_extract(co, website, old_hq, snippets or "(no text fetched)")
        chosen = ""
        source = ""
        # Prefer agent verified when present
        if co in AGENT_VERIFIED:
            chosen, source = AGENT_VERIFIED[co]
        elif (
            ds.get("hq")
            and looks_like_city_country(ds["hq"])
            and ds.get("confidence") in {"high", "medium"}
        ):
            chosen = ds["hq"]
            source = f"deepseek:{ds.get('confidence')}:{ds.get('reason')}"

        # Manual high-confidence overrides from this session (only if clear)
        MANUAL = {
            "Primary Care Supplies": "Aylesford, United Kingdom",
            "BioD Medica": "Subang Jaya, Malaysia",
            "MyBanzu Healthcare": "Puchong, Malaysia",
        }
        if co in MANUAL:
            chosen = MANUAL[co]
            source = source or "agent_web_verified"

        # Farmacias Carol: correct to Santo Domingo if DeepSeek/agent agree DR
        if co == "Farmacias Carol":
            hq_ds = (ds.get("hq") or "").lower()
            if "santo domingo" in hq_ds or "dominican" in (ds.get("reason") or "").lower():
                chosen = "Santo Domingo, Dominican Republic"
                source = "linkedin/datanyze + deepseek"
            elif "santo domingo" in snippets.lower():
                chosen = "Santo Domingo, Dominican Republic"
                source = "agent_web_snippets"

        if co == "Farmacias Siman":
            # Prefer city from sources; if sheet country El Salvador but HQ San Pedro Sula HN, use verified HQ city+country
            if "san pedro sula" in snippets.lower() or "san pedro sula" in (ds.get("hq") or "").lower():
                chosen = "San Pedro Sula, Honduras"
                source = "cbinsights + snippets"

        if co == "Farmacias La Paz":
            if "tegucigalpa" in (ds.get("hq") or "").lower() or (
                ds.get("confidence") in {"high", "medium"} and looks_like_city_country(ds.get("hq") or "")
            ):
                if looks_like_city_country(ds.get("hq") or ""):
                    chosen = ds["hq"]
                    source = f"deepseek:{ds.get('confidence')}"
            if "tegucigalpa" in snippets.lower() and not chosen:
                # Only if DeepSeek also medium+ — weak link to Inversiones La Paz
                if ds.get("confidence") in {"high", "medium"} and "tegucigalpa" in (ds.get("hq") or "").lower():
                    chosen = "Tegucigalpa, Honduras"
                    source = "deepseek+inversiones_la_paz_context"

        entry = {
            "company": co,
            "website": website,
            "old_hq": old_hq,
            "deepseek": ds,
            "snippets_len": len(snippets),
            "chosen": chosen,
            "source": source,
        }
        if chosen and looks_like_city_country(chosen) and chosen != old_hq:
            ws.cell(info["row"], hq_i).value = chosen
            updated.append(f"{co}: {old_hq} -> {chosen} ({source})")
            entry["status"] = "updated"
        else:
            entry["status"] = "left_unchanged"
        results.append(entry)
        print(entry["status"], co, "->", chosen or "(none)", flush=True)

    wb.save(XLSX)

    # recount
    def has_city(hq: str) -> bool:
        s = (hq or "").strip()
        if not s or "," not in s:
            return False
        low = s.lower()
        if low in {"singapore, singapore", "hong kong sar, china", "hong kong, china"}:
            return True
        parts = [p.strip() for p in s.split(",") if p.strip()]
        if len(parts) < 2:
            return False
        countries_only = {
            "united states",
            "usa",
            "united kingdom",
            "uk",
            "germany",
            "france",
            "japan",
            "china",
            "india",
            "brazil",
            "canada",
            "australia",
            "singapore",
            "malaysia",
            "south africa",
            "sweden",
            "spain",
            "italy",
            "netherlands",
            "belgium",
            "switzerland",
            "denmark",
            "finland",
            "norway",
            "mexico",
            "korea",
            "south korea",
            "uae",
            "united arab emirates",
            "hong kong",
            "saudi arabia",
            "europe",
            "iraq",
            "honduras",
            "el salvador",
            "bolivia",
            "puerto rico",
        }
        return parts[0].lower() not in countries_only

    city = 0
    still = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, co_i).value or "").strip()
        hq = str(ws.cell(r, hq_i).value or "").strip()
        if has_city(hq):
            city += 1
        else:
            still.append({"company": co, "hq": hq, "website": str(ws.cell(r, web_i).value or "").strip()})

    audit = {
        "backup": bak.name,
        "updated": updated,
        "results": results,
        "city_country_total": city,
        "country_only_remaining": len(still),
        "still_all": still,
        "rows": ws.max_row - 1,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("UPDATED", len(updated))
    for u in updated:
        print(" ", u)
    print("city", city, "country_only", len(still))
    for s in still:
        print(" STILL", s["hq"], "|", s["company"])


if __name__ == "__main__":
    main()
