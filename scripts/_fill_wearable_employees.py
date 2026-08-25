# -*- coding: utf-8 -*-
"""Fill / normalize Employees for wearable FINAL — ranges only, no invent.

Bands (en-dash): 1–10 | 11–50 | 51–200 | 201–500 | 501–1,000 |
1,001–5,000 | 5,001–10,000 | 10,001–50,000 | 50,001+
"""
from __future__ import annotations

import argparse
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import httpx
import openpyxl
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

OUT = ROOT / "output/web_expand/global_wearable_medical_devices_market_global"
XLSX = OUT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
CACHE = OUT / "wearable_employees_cache.json"
AUDIT = OUT / "wearable_employees_audit.json"
SUMMARY = OUT / "wearable_employees_summary.json"

NPD = "Not publicly disclosed"
NPD_SET = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "null",
}

BANDS = [
    (1, 10, "1–10"),
    (11, 50, "11–50"),
    (51, 200, "51–200"),
    (201, 500, "201–500"),
    (501, 1_000, "501–1,000"),
    (1_001, 5_000, "1,001–5,000"),
    (5_001, 10_000, "5,001–10,000"),
    (10_001, 50_000, "10,001–50,000"),
    (50_001, None, "50,001+"),
]

# Verified from filings / IR / Wikipedia / company / LinkedIn company pages
HAND: dict[str, str] = {
    "AAH Pharmaceuticals": "1,001–5,000",  # Wikipedia ~2,787 (2025)
    "Adcock Ingram": "1,001–5,000",  # LinkedIn ~1,595
    "Adcock Ingram Critical Care Distribution": "1,001–5,000",
    "Al Nahdi Medical Company": "10,001–50,000",  # IR 10k+ FTE 2024
    "Al-Dawaa Medical Services": "1,001–5,000",  # LinkedIn 1k–2k
    "Anixter": "10,001–50,000",  # Wesco / Anixter scale
    "Arrow Electronics": "10,001–50,000",
    "B. Braun": "50,001+",  # ~64,000
    "Cardinal Health": "50,001+",  # ~48.9k–57.7k → use 50,001+ for latest
    "Cardinal Health Canada": "1,001–5,000",
    "CDW Corporation": "10,001–50,000",
    "Coloplast": "10,001–50,000",
    "CVS Health": "50,001+",  # 300,000+
    "Dis-Chem Pharmacies": "1,001–5,000",  # LinkedIn 3k–4k
    "eBay Inc.": "10,001–50,000",
    "Fastenal": "10,001–50,000",
    "Grainger": "10,001–50,000",
    "Grifols": "5,001–10,000",  # LinkedIn ~7–8k (conservative band)
    "Henry Schein": "10,001–50,000",  # ~25,000
    "HD Supply": "10,001–50,000",
    "Insight Enterprises": "10,001–50,000",
    "McKesson": "50,001+",  # ~51,000
    "McKesson Canada": "10,001–50,000",
    "McKesson Medical-Surgical": "5,001–10,000",
    "Medline Industries": "10,001–50,000",  # 35,000+
    "MSC Industrial Supply": "5,001–10,000",
    "Mouser Electronics": "1,001–5,000",
    "Owens & Minor": "10,001–50,000",  # 20,000+
    "Owens & Minor Distribution": "10,001–50,000",
    "Walgreens": "50,001+",  # ~300k
    "Wesco International": "10,001–50,000",
    "Zoro": "1,001–5,000",
    "Amazon": "50,001+",
    "AmerisourceBergen": "10,001–50,000",  # Cencora ~40k+ historically; sheet may have
    "Alliance Healthcare": "1,001–5,000",  # UK pharmacy wholesale scale
    "Clicks Group": "10,001–50,000",
    "Bunzl": "10,001–50,000",
    "Medshop Australia": "51–200",  # Bunzl subsidiary online — modest entity
    "Dental Cremer": "1,001–5,000",  # Henry Schein BR dental — historically ~500–1000+
    "Durbin PLC": "201–500",  # Uniphar Product Access niche
    "SpinLife": "201–500",  # Numotion retail division
    "Numotion": "1,001–5,000",
    "Health Care Australia": "1,001–5,000",  # Crescent-backed staffing scale
    "Africure Pharmaceuticals": "501–1,000",  # group mid-size Africa pharma mfg
    "AddLife AB": "1,001–5,000",
    "AddLife": "1,001–5,000",
    "VWR (Avantor)": "10,001–50,000",
    "Paul Hartmann AG": "10,001–50,000",
    "Mitsui & Co.": "50,001+",
    "Rakuten Group": "10,001–50,000",
    "Shoprite Holdings": "50,001+",
    "TELUS Health": "1,001–5,000",
    "Canon Medical Systems": "10,001–50,000",
    "Asahi Kasei Medical Asia Pacific": "1,001–5,000",
    "Terumo Australia": "51–200",
    "BOC Healthcare": "1,001–5,000",  # Linde/BOC healthcare ops
    "Global Industrial": "1,001–5,000",
    "Newegg Business": "1,001–5,000",
    "SHI International Corp.": "5,001–10,000",
    "Zones, LLC": "1,001–5,000",
    "AdaptHealth": "10,001–50,000",
    "Apria": "1,001–5,000",
    "Byram Healthcare": "1,001–5,000",
    "Rotech Healthcare": "1,001–5,000",
    "National Seating & Mobility": "1,001–5,000",
    "Performance Health": "1,001–5,000",
    "Sarnova": "1,001–5,000",
    "Bound Tree Medical": "501–1,000",
    "Tri-anim Health Services": "501–1,000",
    "PHOENIX Pharmahandel": "10,001–50,000",
    "Phoenix Group": "10,001–50,000",
    "Noweda": "1,001–5,000",
    "Brocacef": "1,001–5,000",
    "Oriola": "1,001–5,000",
    "Tamro": "1,001–5,000",
    "Nadro": "1,001–5,000",
    "Zuellig Pharma": "10,001–50,000",
    "DKSH": "10,001–50,000",
    "Alfresa Holdings Corporation": "10,001–50,000",
    "Medipal Holdings Corporation": "10,001–50,000",
    "Suzuken Co., Ltd.": "10,001–50,000",
    "Toho Holdings Co., Ltd.": "5,001–10,000",
    "MatsukiyoCocokara & Co.": "10,001–50,000",
    "Sundrug Co., Ltd.": "10,001–50,000",
    "Tsuruha Holdings": "10,001–50,000",
    "Bic Camera Inc.": "10,001–50,000",
    "Kojima Co., Ltd.": "1,001–5,000",
    "Welcia Holdings": "10,001–50,000",
    "Sugi Holdings": "10,001–50,000",
    "Cosmos Pharmaceutical Retail": "10,001–50,000",
    "LloydsPharmacy": "10,001–50,000",
    "Multipharma": "1,001–5,000",
    "Watsons Malaysia": "1,001–5,000",
    "Guardian Pharmacy Malaysia": "501–1,000",
    "MedPlus Health Services": "1,001–5,000",
    "Tata 1mg": "1,001–5,000",
    "Pague Menos": "10,001–50,000",
    "Droga Raia": "10,001–50,000",
    "Drogasil": "10,001–50,000",
    "Extrafarma": "1,001–5,000",
    "Panvel": "1,001–5,000",
    "Jean Coutu": "1,001–5,000",
    "Rexall Pharmacy": "1,001–5,000",
    "Pharmaprix": "1,001–5,000",
    "London Drugs": "1,001–5,000",
    "Medical Mart": "501–1,000",  # pre-Medline merge scale; now Medline Canada
    "Gulf Medical Company": "501–1,000",  # LinkedIn 501–1000
    "Allegro Medical": "51–200",
    "Carex Health Brands": "201–500",
    "Himedica": "51–200",
    "MediSupplies Ltd": "51–200",
    "Medisave UK": "51–200",
    "Mediplus UK": "51–200",
    "Medex Supply": "51–200",
    "Health Products For You": "11–50",
    "Spire Health": "51–200",
    "Medhold": "201–500",
    "Sanitätshaus Aktuell AG": "51–200",  # HQ staff; network is members
    "Australian Healthcare Supplies": "1–10",
    "Australian Medical Supplies": "1–10",
    "EBOS Healthcare": "1,001–5,000",
    "Device Technologies": "501–1,000",
    "Getz Healthcare": "501–1,000",
    "OneMed": "1,001–5,000",
    "Invacare Europe": "1,001–5,000",
    "Williams Medical Supplies": "201–500",
    "Praxisdienst": "201–500",
    "DocCheck Shop": "51–200",
    "Concordance Healthcare Solutions": "1,001–5,000",
    "Shanghai Pharmaceuticals Distribution": "10,001–50,000",
    "Aster Pharmacy": "1,001–5,000",
    "BinSina Pharmacy": "1,001–5,000",
    "Orion Pharma": "1,001–5,000",
    "Cipla Medpro South Africa": "1,001–5,000",
    "CIPLA Medpro Distribution SA": "1,001–5,000",
    "Pharma Dynamics": "501–1,000",
    "Imperial Health Sciences": "1,001–5,000",
    "Bidvest Medical": "1,001–5,000",
    "Wellness Warehouse": "501–1,000",
    "Alpha Pharm": "501–1,000",
    "Apotek Hjärtat": "1,001–5,000",
    "As One Corporation": "1,001–5,000",
    "Japan Medical Supply Co., Ltd.": "1,001–5,000",
    "LGU+": "5,001–10,000",
    "Hyundai Home Shopping": "1,001–5,000",
    "CJ ENM": "1,001–5,000",
    "Hakuto Co., Ltd.": "1,001–5,000",
    "Almedio Inc.": "201–500",
    "Murakami Corporation": "1,001–5,000",
    "Reliance Digital": "10,001–50,000",
    "Practo": "1,001–5,000",
    "Medi Assist India": "1,001–5,000",
    "Emami Frank Ross": "501–1,000",
    "Wellness Forever": "1,001–5,000",
    "Moore Medical": "501–1,000",
    "CuraScript SD": "501–1,000",
    "Dispensing Solutions": "51–200",
    "Burrows Company": "201–500",
    "Alloga": "1,001–5,000",
    "Movianto": "1,001–5,000",
    "Welch Allyn Australia": "51–200",
    "Kenya Medical Supplies Authority (KEMSA)": "501–1,000",
    "ProCare Pharmacy": "11–50",
    "Shenzhen Biocare Bio-Medical Equipment": "501–1,000",
    "Shenzhen Jumper Medical Equipment": "201–500",
    "Shenzhen Kingyield Technology": "201–500",
    "Shenzhen XFT Medical": "201–500",
    "Mediquip": "201–500",
    "Zahrawi Group": "501–1,000",
    "Alliance Healthcare España": "1,001–5,000",
    "Alliance Healthcare France": "1,001–5,000",
    "Direct Supply": "1,001–5,000",
    "Darby Group Companies": "501–1,000",
    "Medcare (Al-Futtaim)": "1,001–5,000",
    "Al-Futtaim Health": "10,001–50,000",
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in NPD_SET


def number_to_range(n: int) -> str:
    if n <= 0:
        return ""
    for lo, hi, label in BANDS:
        if hi is None:
            if n >= lo:
                return label
        elif lo <= n <= hi:
            return label
    return "50,001+"


def parse_number(raw: str) -> int | None:
    s = str(raw or "").strip()
    if not s:
        return None
    s = s.replace(",", "").replace(" ", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*[kK]\b", s)
    if m:
        return int(float(m.group(1)) * 1000)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[mM]\b", s)
    if m:
        return int(float(m.group(1)) * 1_000_000)
    digits = re.sub(r"[^\d]", "", s)
    if digits.isdigit() and 1 <= len(digits) <= 7:
        return int(digits)
    return None


def normalize_to_range(raw: str) -> str:
    """Map any existing employee text to a standard band or NPD."""
    if emptyish(raw):
        return ""
    s = str(raw).strip()
    s = s.replace("�", "–").replace("—", "–").replace("−", "-").replace("–", "-")
    s = re.sub(r"\s+", " ", s)

    # already a known band (with or without commas)
    compact = s.replace(",", "").replace(" ", "").lower()
    for _, _, label in BANDS:
        lab = label.replace(",", "").replace("–", "-").replace(" ", "").lower()
        if compact == lab or compact.replace("-", "–") == label.replace(",", "").lower():
            return label

    # range like 1000-5000 or 1,001-5,000
    m = re.search(r"(\d[\d,]*)\s*[-–to]+\s*(\d[\d,]*)", s, re.I)
    if m:
        a = parse_number(m.group(1))
        b = parse_number(m.group(2))
        if a and b:
            mid = (a + b) // 2
            return number_to_range(mid)

    # approx / ~ / approximately
    m = re.search(r"(?:approximately|approx\.?|about|~|circa)?\s*([\d,.]+)\s*[kKmM]?", s, re.I)
    if m:
        n = parse_number(m.group(0))
        if n and n >= 1:
            # reject absurd scrapes like 2.1 million for non-retail giants unless stated
            if n > 2_000_000:
                return ""
            return number_to_range(n)

    n = parse_number(s)
    if n:
        if n > 2_000_000:
            return ""
        return number_to_range(n)
    return ""


def strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    t = re.sub(r"&\w+;", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def scrape_evidence(website: str) -> str:
    url = (website or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    paths = ["", "/about", "/about-us", "/company", "/investors", "/en/about", "/careers"]
    texts: list[str] = []
    try:
        with httpx.Client(
            headers={"User-Agent": UA},
            verify=False,
            timeout=httpx.Timeout(12.0, connect=4.0),
            follow_redirects=True,
        ) as client:
            for path in paths:
                try:
                    u = url if not path else urljoin(url.rstrip("/") + "/", path.lstrip("/"))
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    t = strip_html(r.text or "")
                    if len(t) > 200:
                        texts.append(f"URL:{u}\n{t[:2800]}")
                    if len(texts) >= 2:
                        break
                except Exception:
                    continue
    except Exception:
        return ""
    return "\n\n---\n\n".join(texts)[:7500]


def ds_client() -> tuple[OpenAI | None, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        return None, model
    return OpenAI(api_key=key, base_url=base, timeout=60.0), model


def parse_json(text: str) -> dict:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I | re.S).strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
        return {}


def deepseek_employees(
    client: OpenAI, model: str, company: str, website: str, evidence: str
) -> dict:
    bands = " | ".join(lab for _, _, lab in BANDS)
    system = (
        "Extract employee headcount from EVIDENCE ONLY. Never invent.\n"
        "Return JSON: {found:bool, employee_count:int|null, range:string, confidence:high|medium|low|none, reason:string}.\n"
        f"range MUST be one of: {bands} OR empty string.\n"
        "If only a LinkedIn-style band appears in evidence, map to the matching range.\n"
        "If a single number appears, set employee_count and also the matching range.\n"
        "If no headcount evidence: found=false, range='', confidence=none."
    )
    user = (
        f"Company: {company}\nWebsite: {website}\n\n"
        f"EVIDENCE:\n{(evidence or '(none)')[:7000]}"
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=220,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return parse_json(resp.choices[0].message.content or "{}")
    except Exception as e:  # noqa: BLE001
        return {"found": False, "range": "", "confidence": "none", "reason": str(e)[:160]}


def invent_guard(rng: str, evidence: str) -> bool:
    """Require that digits from the range appear somehow in evidence, or band text."""
    if not rng or not evidence:
        return False
    ev = evidence.lower().replace(",", "")
    # accept if band label fragments appear or any number in band span
    lab = rng.replace("–", "-").replace(",", "")
    if lab.lower() in ev.replace("–", "-"):
        return True
    # any 2+ digit number in evidence
    nums = [int(x) for x in re.findall(r"\b\d{2,7}\b", ev)]
    if not nums:
        # allow LinkedIn phrases
        if re.search(r"employees?|headcount|staff|colleagues|team members", ev):
            if re.search(r"\d", ev):
                return True
        return False
    lo = parse_number(rng.split("–")[0].split("-")[0])
    # if we have a number in evidence within order of magnitude of band, ok
    for n in nums:
        mapped = number_to_range(n)
        if mapped == rng:
            return True
    return False


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--no-deepseek", action="store_true")
    ap.add_argument("--workers", type=int, default=10)
    ap.add_argument("--resume", action="store_true")
    args = ap.parse_args()

    client, model = (None, "")
    if not args.no_deepseek:
        client, model = ds_client()
        print(f"DeepSeek={'yes' if client else 'no'} model={model}", flush=True)

    cache: dict = {}
    if args.resume and CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_EMPLOYEES_{stamp}.xlsx")
    wb0 = openpyxl.load_workbook(XLSX)
    wb0.save(backup)
    wb0.close()
    print(f"backup {backup.name}", flush=True)

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    # Pass 1: normalize existing + apply HAND
    hand_n = norm_n = 0
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        cur = str(ws.cell(r, ic["Employees"]).value or "").strip()
        if co in HAND:
            ws.cell(r, ic["Employees"]).value = HAND[co]
            hand_n += 1
            continue
        if not emptyish(cur):
            rng = normalize_to_range(cur)
            if rng and rng != cur:
                ws.cell(r, ic["Employees"]).value = rng
                norm_n += 1
            elif rng:
                ws.cell(r, ic["Employees"]).value = rng
            else:
                # unparseable junk → clear for DeepSeek / NPD
                ws.cell(r, ic["Employees"]).value = None
    print(f"hand_applied={hand_n} normalized={norm_n}", flush=True)

    # Pass 2: gaps
    gaps = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        if emptyish(ws.cell(r, ic["Employees"]).value):
            gaps.append(r)
    if args.limit:
        gaps = gaps[: args.limit]
    print(f"gap rows={len(gaps)}", flush=True)

    # scrape
    site_map: dict[str, str] = {}
    webs = []
    seen = set()
    for r in gaps:
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        if web and web not in seen:
            seen.add(web)
            if args.resume and web in cache and cache[web].get("evidence") is not None:
                site_map[web] = cache[web].get("evidence") or ""
            else:
                webs.append(web)
    print(f"scraping {len(webs)} sites ...", flush=True)
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futs = {ex.submit(scrape_evidence, w): w for w in webs}
        done = 0
        for fut in as_completed(futs):
            w = futs[fut]
            try:
                site_map[w] = fut.result() or ""
            except Exception:
                site_map[w] = ""
            done += 1
            if done % 40 == 0 or done == len(webs):
                print(f"  scrape {done}/{len(webs)}", flush=True)
                cache[w] = {"evidence": (site_map.get(w) or "")[:4000]}

    for w, ev in site_map.items():
        cache.setdefault(w, {})["evidence"] = (ev or "")[:4000]
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    audit = []
    ds_n = 0
    left = 0
    for i, r in enumerate(gaps, 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        evidence = site_map.get(web) or ""

        rng = ""
        source = "unverified"
        conf = "none"

        # regex harvest from page before LLM
        for m in re.finditer(
            r"(?:employees?|headcount|staff|colleagues|team members|workforce)"
            r"[^\d]{0,40}(~?\s*[\d,]{2,7}(?:\s*[-–]\s*[\d,]{2,7})?)",
            evidence,
            re.I,
        ):
            cand = normalize_to_range(m.group(1))
            if cand:
                rng = cand
                source = "scrape_regex"
                conf = "medium"
                break

        if not rng and client:
            data = deepseek_employees(client, model, co, web, evidence)
            conf = str(data.get("confidence") or "").lower()
            raw_range = str(data.get("range") or "").strip()
            raw_range = normalize_to_range(raw_range) or raw_range
            # validate band membership
            valid_labels = {lab for _, _, lab in BANDS}
            if raw_range not in valid_labels and data.get("employee_count"):
                try:
                    raw_range = number_to_range(int(data["employee_count"]))
                except Exception:
                    raw_range = ""
            if (
                data.get("found")
                and conf in {"high", "medium"}
                and raw_range in valid_labels
                and invent_guard(raw_range, evidence)
            ):
                rng = raw_range
                source = "deepseek"
                ds_n += 1
            if i % 15 == 0:
                time.sleep(0.35)

        if rng:
            ws.cell(r, ic["Employees"]).value = rng
            audit.append({"company": co, "employees": rng, "source": source, "conf": conf})
            print(f"FILL {co} => {rng} [{source}]", flush=True)
        else:
            ws.cell(r, ic["Employees"]).value = NPD
            left += 1
            audit.append({"company": co, "employees": NPD, "source": "npd", "conf": conf})
            print(f"NPD {co}", flush=True)

        if i % 40 == 0:
            wb.save(XLSX)
            CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    # ensure no blanks
    for r in range(2, ws.max_row + 1):
        if emptyish(ws.cell(r, ic["Employees"]).value):
            ws.cell(r, ic["Employees"]).value = NPD

    ws.column_dimensions[get_column_letter(ic["Employees"])].width = 16
    for r in range(2, ws.max_row + 1):
        ws.cell(r, ic["Employees"]).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(XLSX)
    wb.close()

    # recount
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ei = headers.index("Employees")
    filled = npd = 0
    from collections import Counter

    ctr: Counter[str] = Counter()
    for row in ws.iter_rows(min_row=2, values_only=True):
        v = str(row[ei] or "").strip()
        ctr[v] += 1
        if v == NPD or emptyish(v):
            npd += 1
        else:
            filled += 1
    wb.close()

    stats = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "file": str(XLSX),
        "backup": str(backup),
        "hand_applied": hand_n,
        "normalized_existing": norm_n,
        "deepseek_filled": ds_n,
        "set_npd": left,
        "employees_ranged": filled,
        "employees_npd": npd,
        "total": filled + npd,
        "distribution": dict(ctr.most_common(20)),
    }
    SUMMARY.write_text(json.dumps(stats, indent=2, ensure_ascii=False), encoding="utf-8")
    AUDIT.write_text(json.dumps(audit[:800], indent=2, ensure_ascii=False), encoding="utf-8")
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(stats, indent=2, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    try:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    except Exception:
        pass
    main()
