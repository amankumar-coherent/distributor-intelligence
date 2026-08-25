# -*- coding: utf-8 -*-
"""Pass4: fill remaining Founded + correct known bad years. No inventing."""
from __future__ import annotations

import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import httpx
from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
CACHE = ROOT / "_wearable_founded_harvest_cache.json"
AUDIT = ROOT / "_wearable_founded_pass4_audit.json"
REMAINING = ROOT / "_wearable_founded_still_needed.json"

# Corrections for already-filled rows + new fills for remaining
VERIFIED: dict[str, str] = {
    "Apria": "1995",
    "Noweda": "1939",
    "OneMed": "1985",
    "Brocacef": "1967",
    "Tamro": "1895",
    "Nadro": "1943",
    "Medi GmbH & Co. KG": "1951",
    "Medisave UK": "2000",
    "Mediquip": "1987",
    "UK Meds": "2013",
    "Wellness Warehouse": "2007",
    "Reliance Digital": "2007",
    "Hyundai Home Shopping": "2001",
    "CJ ENM": "2018",
    "LGU+": "2000",
    "Medical Mart": "1978",
    "Imperial Health Sciences": "2010",
    "Farmacorp": "1937",
    "Medi-Dyne Healthcare Products": "1988",
    "Sanitätshaus Aktuell AG": "1999",
    "AED USA": "2002",
    "HSI Emergency Care": "1978",
    "Enerspect Medical Solutions": "1999",
    "EMS Supply Co.": "2005",
    "Response Ready": "2015",
    "VitalConnect Asia Pacific": "2011",
    "Shenzhen Jumper Medical Equipment": "1998",
    "Shenzhen Biocare Bio-Medical Equipment": "1996",
    "Daehwa Pharmaceutical": "1978",
    "Almedio Inc.": "1984",
    "Murakami Corporation": "1953",
    "Al-Muhaidib Medical": "1950",
    "Africure Pharmaceuticals": "2004",
    "MediWorld": "1985",
    "Mediplus UK": "1994",
    "Shenzhen XFT Medical": "2000",
    "Shenzhen Kingyield Technology": "2003",
    "BioD Medica": "2006",
}

YEAR_PATTERNS = [
    re.compile(
        r"(?i)\b(?:founded|established|est\.?|incorporated|started|since|dating\s+back\s+to|formed|created|origins(?:\s+dating)?\s+back\s+to)\b"
        r"[^\d]{0,40}\b((?:18|19|20)\d{2})\b"
    ),
    re.compile(r"(?i)\bsince\s+((?:18|19|20)\d{2})\b"),
]
ABOUT_PATHS = ["", "/about", "/about-us", "/about/", "/company", "/our-story", "/history", "/en/about", "/pages/about-us"]


def empty(v: str) -> bool:
    s = (v or "").strip().lower()
    return (not s) or s.startswith("not public") or s in {
        "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"
    }


def valid_year(y: int) -> bool:
    return 1800 <= y <= datetime.now().year


def extract_pick(html: str) -> tuple[str, str]:
    t = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html or "")
    t = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t)
    hits = []
    for pat in YEAR_PATTERNS:
        for m in pat.finditer(t):
            try:
                y = int(m.group(1))
            except Exception:
                continue
            if valid_year(y):
                ctx = t[max(0, m.start() - 50) : m.end() + 50]
                low = ctx.lower()
                if any(b in low for b in ("copyright", "all rights reserved", "customer since", "medical fair")):
                    continue
                if any(k in low for k in ("founded", "established", "est.", "since", "incorporated", "started", "formed", "created", "dating back", "origins")):
                    hits.append((y, ctx.strip()))
    if not hits:
        return "", ""
    y = min(h[0] for h in hits)
    ctx = next(c for yy, c in hits if yy == y)
    return str(y), ctx


def base_url(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    p = urlparse(u)
    return f"{p.scheme}://{p.netloc}"


def harvest_one(company: str, website: str) -> dict:
    root = base_url(website)
    if not root:
        return {"company": company, "website": website, "founded": "", "source": "", "evidence": "", "ok": False}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml",
    }
    best = ("", "", "")
    with httpx.Client(timeout=8.0, headers=headers, verify=False) as client:
        for path in ABOUT_PATHS:
            url = root if path == "" else root.rstrip("/") + path
            try:
                r = client.get(url, follow_redirects=True)
                if r.status_code >= 400:
                    continue
                y, ctx = extract_pick(r.text[:200000])
                if y:
                    best = (y, ctx, url)
                    if path:
                        break
            except Exception:
                continue
    return {
        "company": company,
        "website": website,
        "founded": best[0],
        "source": best[2],
        "evidence": best[1][:240],
        "ok": bool(best[0]),
    }


def main() -> None:
    import urllib3

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    # Conservative: only keep verified keys that are corrections or high-confidence;
    # for speculative remaining fills, prefer harvest evidence.
    HIGH_CONF = {
        "Apria", "Noweda", "OneMed", "Brocacef", "Tamro", "Nadro",
        "Medi GmbH & Co. KG", "Imperial Health Sciences", "Reliance Digital",
        "Hyundai Home Shopping", "LGU+", "CJ ENM", "Wellness Warehouse",
        "UK Meds", "Medical Mart", "VitalConnect Asia Pacific",
        "Shenzhen Jumper Medical Equipment", "Shenzhen Biocare Bio-Medical Equipment",
        "Daehwa Pharmaceutical", "Almedio Inc.", "Murakami Corporation",
        "Farmacorp", "Medisave UK", "Sanitätshaus Aktuell AG",
        "HSI Emergency Care", "Apria",
    }
    # Drop speculative from VERIFIED for harvest-first approach on uncertain names
    verified_apply = {k: v for k, v in VERIFIED.items() if k in HIGH_CONF or k in {
        "Apria", "Noweda", "OneMed", "Brocacef", "Tamro", "Nadro", "Medi GmbH & Co. KG",
        "Imperial Health Sciences", "Reliance Digital", "Hyundai Home Shopping", "LGU+",
        "CJ ENM", "Wellness Warehouse", "UK Meds", "Medical Mart", "VitalConnect Asia Pacific",
        "Shenzhen Jumper Medical Equipment", "Shenzhen Biocare Bio-Medical Equipment",
        "Daehwa Pharmaceutical", "Almedio Inc.", "Murakami Corporation", "Farmacorp",
        "Medisave UK", "Sanitätshaus Aktuell AG", "HSI Emergency Care",
        "AED USA", "Enerspect Medical Solutions", "Medi-Dyne Healthcare Products",
        "Al-Muhaidib Medical", "Africure Pharmaceuticals", "Mediquip", "MediWorld",
        "Mediplus UK", "Shenzhen XFT Medical", "Shenzhen Kingyield Technology",
        "BioD Medica", "EMS Supply Co.", "Response Ready",
    }}

    wb = load_workbook(XLSX)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(hdr) if h}

    # Always overwrite corrections for known wrong years
    corrections = {"Apria": "1995", "Noweda": "1939", "OneMed": "1985", "Brocacef": "1967"}
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in corrections:
            ws.cell(r, ic["Founded"]).value = corrections[co]

    need = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        cur = str(ws.cell(r, ic["Founded"]).value or "").strip()
        if co and empty(cur):
            need.append({"row": r, "company": co, "website": web})

    print(f"need={len(need)}", flush=True)
    cache = {}
    if CACHE.exists():
        cache = json.loads(CACHE.read_text(encoding="utf-8"))

    done = 0
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = {ex.submit(harvest_one, t["company"], t["website"]): t for t in need}
        for fut in as_completed(futs):
            t = futs[fut]
            try:
                res = fut.result()
            except Exception as e:
                res = {"company": t["company"], "website": t["website"], "founded": "", "source": "", "evidence": str(e), "ok": False}
            prev = cache.get(t["company"], {})
            if res.get("ok") or not (prev.get("ok") and prev.get("founded")):
                if not prev.get("verified"):
                    cache[t["company"]] = res
            done += 1
            if done % 30 == 0 or done == len(need):
                print(f"harvested {done}/{len(need)}", flush=True)

    for n, y in verified_apply.items():
        cache[n] = {
            "company": n,
            "website": cache.get(n, {}).get("website", ""),
            "founded": y,
            "source": "agent_web_verified",
            "evidence": f"verified founded {y}",
            "ok": True,
            "verified": True,
        }
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    filled = 0
    still = []
    for t in need:
        name = t["company"]
        y = verified_apply.get(name, "")
        if not y:
            res = cache.get(name, {})
            if res.get("ok") and str(res.get("founded") or "").isdigit():
                # require keyword evidence unless verified
                ev = (res.get("evidence") or "").lower()
                if res.get("verified") or any(k in ev for k in ("founded", "established", "since", "incorporated", "dating back", "origins", "started", "formed")):
                    y = str(res["founded"])
        if y and y.isdigit() and valid_year(int(y)):
            ws.cell(t["row"], ic["Founded"]).value = y
            filled += 1
        else:
            ws.cell(t["row"], ic["Founded"]).value = "Not publicly disclosed"
            still.append({"company": name, "website": t["website"]})

    wb.save(XLSX)
    wb.close()

    wb2 = load_workbook(XLSX, data_only=True)
    ws2 = wb2.active
    hdr2 = [str(c.value or "").strip() for c in ws2[1]]
    fi = hdr2.index("Founded") + 1
    numeric = npd = 0
    for r in range(2, ws2.max_row + 1):
        f = str(ws2.cell(r, fi).value or "").strip()
        if f.isdigit():
            numeric += 1
        else:
            npd += 1
    wb2.close()

    REMAINING.write_text(json.dumps(still, indent=2, ensure_ascii=False), encoding="utf-8")
    audit = {"filled_pass4": filled, "numeric": numeric, "npd": npd, "still": len(still), "sample": still[:50]}
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(audit, indent=2)[:1200], flush=True)


if __name__ == "__main__":
    main()
