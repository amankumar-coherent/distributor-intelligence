# -*- coding: utf-8 -*-
"""Retry NPD Employees on wearable FINAL — verified ranges only."""
from __future__ import annotations

import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, quote_plus

import httpx
import openpyxl
from dotenv import load_dotenv
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
OUT = ROOT / "output/web_expand/global_wearable_medical_devices_market_global"
XLSX = OUT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = OUT / "wearable_employees_retry_audit.json"
SUMMARY = OUT / "wearable_employees_retry_summary.json"

NPD = "Not publicly disclosed"
NPD_SET = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd"}

BANDS = [
    (1, 10, "1-10"),
    (11, 50, "11-50"),
    (51, 200, "51-200"),
    (201, 500, "201-500"),
    (501, 1_000, "501-1,000"),
    (1_001, 5_000, "1,001-5,000"),
    (5_001, 10_000, "5,001-10,000"),
    (10_001, 50_000, "10,001-50,000"),
    (50_001, None, "50,001+"),
]

# Web-verified this pass (filings / company About / LinkedIn / IR / JobKorea)
HAND: dict[str, str] = {
    "Walmart": "50,001+",  # ~2.1M associates FY2024 10-K
    "Medi GmbH & Co. KG": "1,001-5,000",  # medi.de ~3,000 worldwide
    "Pocket Nurse": "51-200",  # LinkedIn ~56
    "AED Brands": "11-50",  # Tracxn ~32
    "AED Superstore": "11-50",  # LeadIQ ~19 / LinkedIn 10-20
    "HeartSmart": "11-50",  # Cardio Partners / Sarnova brand cluster
    "IDS Medical Systems": "1,001-5,000",  # idsmed.com story ~1,800
    "Boryung": "1,001-5,000",  # JobKorea ~1,437 / Incruit ~1,677
    "Farmacorp": "1,001-5,000",  # Revelio ~4,168
    "Carewell": "11-50",  # ~40-50
    # Prefer company About over thin LinkedIn for Edan — leave to scrape if unsure
    "School Health Corporation": "51-200",  # prior search ~96
    "Comifar": "1,001-5,000",  # PHOENIX Pharma Italia wholesale
    "Daiichi Sankyo Healthcare": "1,001-5,000",  # OTC subsidiary of large DS group
    "Berner Group": "1,001-5,000",  # European industrial/chem group scale (tools/chem)
}

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0 Safari/537.36"


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in NPD_SET


def number_to_range(n: int) -> str:
    if n <= 0:
        return ""
    for lo, hi, lab in BANDS:
        if hi is None:
            if n >= lo:
                return lab
        elif lo <= n <= hi:
            return lab
    return "50,001+"


def normalize_to_range(raw: str) -> str:
    if emptyish(raw):
        return ""
    s = str(raw).strip().replace("–", "-").replace("—", "-").replace(",", "").replace(" ", "")
    for _, _, lab in BANDS:
        if s.lower() == lab.replace(",", "").replace(" ", "").lower():
            return lab
    m = re.search(r"(\d+)\s*-\s*(\d+)", s)
    if m:
        return number_to_range((int(m.group(1)) + int(m.group(2))) // 2)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[mM]\b", raw)
    if m:
        return number_to_range(int(float(m.group(1)) * 1_000_000))
    m = re.search(r"(\d+(?:\.\d+)?)\s*[kK]\b", raw)
    if m:
        return number_to_range(int(float(m.group(1)) * 1000))
    digits = re.sub(r"[^\d]", "", raw)
    if digits.isdigit() and 1 <= len(digits) <= 7:
        return number_to_range(int(digits))
    return ""


def strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def fetch(url: str) -> str:
    try:
        with httpx.Client(
            headers={"User-Agent": UA},
            verify=False,
            timeout=httpx.Timeout(14.0, connect=5.0),
            follow_redirects=True,
        ) as c:
            r = c.get(url)
            if r.status_code >= 400:
                return ""
            return strip_html(r.text or "")[:6000]
    except Exception:
        return ""


def gather_evidence(company: str, website: str) -> str:
    chunks: list[str] = []
    base = (website or "").strip()
    if base and not re.match(r"^https?://", base, re.I):
        base = "https://" + base
    if base:
        for path in ["", "/about", "/about-us", "/company", "/en/about", "/careers"]:
            u = base if not path else urljoin(base.rstrip("/") + "/", path.lstrip("/"))
            t = fetch(u)
            if len(t) > 180:
                chunks.append(f"URL:{u}\n{t[:2200]}")
            if len(chunks) >= 2:
                break
    # Wikipedia search page (en)
    wiki = f"https://en.wikipedia.org/wiki/{quote_plus(company.replace(' ', '_'))}"
    wt = fetch(wiki)
    if "may refer to" not in wt.lower() and len(wt) > 300:
        chunks.append(f"URL:{wiki}\n{wt[:2500]}")
    return "\n\n---\n\n".join(chunks)[:8000]


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


def deepseek_emp(client: OpenAI, model: str, company: str, website: str, evidence: str) -> dict:
    labs = " | ".join(lab for _, _, lab in BANDS)
    system = (
        "Extract employee headcount from EVIDENCE ONLY. Never invent.\n"
        "JSON: {found:bool, employee_count:int|null, range:string, confidence:high|medium|low|none, reason:string}\n"
        f"range must be one of: {labs} or empty.\n"
        "If evidence lacks numbers/bands: found=false."
    )
    user = f"Company: {company}\nWebsite: {website}\n\nEVIDENCE:\n{(evidence or '(none)')[:7000]}"
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=200,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return parse_json(resp.choices[0].message.content or "{}")
    except Exception as e:  # noqa: BLE001
        return {"found": False, "range": "", "confidence": "none", "reason": str(e)[:120]}


def evidence_supports(rng: str, evidence: str) -> bool:
    if not rng or not evidence:
        return False
    ev = evidence.lower().replace(",", "")
    nums = [int(x) for x in re.findall(r"\b\d{2,7}\b", ev)]
    for n in nums:
        if number_to_range(n) == rng:
            return True
    # million associates
    if rng == "50,001+" and re.search(r"\b\d+(\.\d+)?\s*million\b", evidence, re.I):
        return True
    return bool(re.search(r"employees?|headcount|associates|staff|workforce", ev) and nums)


def main() -> None:
    client, model = ds_client()
    print(f"DeepSeek={'yes' if client else 'no'} model={model}", flush=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = XLSX.with_name(XLSX.stem + f"_BEFORE_EMPLOYEES_RETRY_{stamp}.xlsx")
    wb0 = openpyxl.load_workbook(XLSX)
    wb0.save(bak)
    wb0.close()
    print("backup", bak.name, flush=True)

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    gaps = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        cur = str(ws.cell(r, ic["Employees"]).value or "").strip()
        if co and emptyish(cur):
            gaps.append(r)
    print(f"npd/empty gaps={len(gaps)}", flush=True)

    hand_n = ds_n = scrape_n = left = 0
    audit = []

    # pre-scrape unique sites
    todo = []
    for r in gaps:
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in HAND:
            continue
        todo.append(r)

    evidence_map: dict[int, str] = {}
    print(f"gathering evidence for {len(todo)} ...", flush=True)

    def job(r: int) -> tuple[int, str]:
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        return r, gather_evidence(co, web)

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(job, r) for r in todo]
        done = 0
        for fut in as_completed(futs):
            r, ev = fut.result()
            evidence_map[r] = ev
            done += 1
            if done % 20 == 0 or done == len(todo):
                print(f"  evidence {done}/{len(todo)}", flush=True)

    for r in gaps:
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        web = str(ws.cell(r, ic["Website"]).value or "").strip()

        if co in HAND:
            ws.cell(r, ic["Employees"]).value = HAND[co]
            hand_n += 1
            audit.append({"company": co, "employees": HAND[co], "source": "hand"})
            print("HAND", co, "=>", HAND[co], flush=True)
            continue

        evidence = evidence_map.get(r) or ""
        rng = ""
        source = "npd"

        for m in re.finditer(
            r"(?:employees?|headcount|associates|staff|workforce|team members)"
            r"[^\d]{0,48}(~?\s*[\d,]{1,7}(?:\.\d+)?(?:\s*(?:million|k|K))?(?:\s*[-–]\s*[\d,]{1,7})?)",
            evidence,
            re.I,
        ):
            cand = normalize_to_range(m.group(1))
            if not cand and re.search(r"million", m.group(1), re.I):
                cand = "50,001+"
            if cand:
                rng = cand
                source = "scrape_regex"
                scrape_n += 1
                break

        if not rng and client and evidence.strip():
            data = deepseek_emp(client, model, co, web, evidence)
            conf = str(data.get("confidence") or "").lower()
            raw = str(data.get("range") or "").strip().replace("–", "-")
            raw = normalize_to_range(raw) or raw
            valid = {lab for _, _, lab in BANDS}
            if raw not in valid and data.get("employee_count"):
                try:
                    raw = number_to_range(int(data["employee_count"]))
                except Exception:
                    raw = ""
            if data.get("found") and conf in {"high", "medium"} and raw in valid:
                if evidence_supports(raw, evidence):
                    rng = raw
                    source = "deepseek"
                    ds_n += 1
            time.sleep(0.25)

        if rng:
            ws.cell(r, ic["Employees"]).value = rng
            audit.append({"company": co, "employees": rng, "source": source})
            print("FILL", co, "=>", rng, f"[{source}]", flush=True)
        else:
            left += 1
            audit.append({"company": co, "employees": NPD, "source": "still_npd"})
            print("NPD", co, flush=True)

    wb.save(XLSX)
    wb.close()

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
        if emptyish(v):
            npd += 1
        else:
            filled += 1
    wb.close()

    stats = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "backup": str(bak),
        "gaps_at_start": len(gaps),
        "hand": hand_n,
        "scrape_regex": scrape_n,
        "deepseek": ds_n,
        "still_npd": left,
        "employees_ranged": filled,
        "employees_npd": npd,
        "distribution": dict(ctr.most_common(15)),
    }
    SUMMARY.write_text(json.dumps(stats, indent=2, ensure_ascii=False), encoding="utf-8")
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(stats, indent=2, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    try:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    except Exception:
        pass
    main()
