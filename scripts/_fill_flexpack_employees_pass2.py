#!/usr/bin/env python3
"""Pass-2 Employees: apply verified HAND + DuckDuckGo snippet extract (no invent)."""
from __future__ import annotations

import json
import os
import re
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import httpx
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import load_workbook

load_dotenv(r"D:\vendor-intel\.env")

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
CACHE = PATH.parent / "flexpack_employees_web_cache.json"
AUDIT = PATH.parent / "flexpack_employees_pass2_audit.json"
NPD = "Not publicly disclosed"
WORKERS = 6

RANGES = (
    "1–10",
    "11–50",
    "51–200",
    "201–500",
    "501–1,000",
    "1,001–5,000",
    "5,001–10,000",
    "10,001–50,000",
    "50,001+",
)

# Additional verified from LinkedIn / LeadIQ / company filings (this pass)
HAND2: dict[str, str] = {
    "Acme Paper & Supply": "51–200",
    "Atlantic Packaging Corp": "201–500",
    "Ball & Doggett": "51–200",
    "Blackwoods": "501–1,000",
    "BradyPLUS": "501–1,000",
    "Burch Bottle & Packaging": "1–10",
    "Container and Packaging Supply": "51–200",
    "HD Supply": "1,001–5,000",  # LinkedIn ~4,695
    "Imperial Dade": "1,001–5,000",  # LinkedIn ~1,261
    "Uline": "5,001–10,000",
}

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in {
        "",
        "n/a",
        "na",
        "-",
        "not publicly disclosed",
        "unknown",
        "tbd",
        "null",
    }


def emp_num_to_range(n: int) -> str:
    if n <= 0:
        return ""
    if n <= 10:
        return "1–10"
    if n <= 50:
        return "11–50"
    if n <= 200:
        return "51–200"
    if n <= 500:
        return "201–500"
    if n <= 1000:
        return "501–1,000"
    if n <= 5000:
        return "1,001–5,000"
    if n <= 10000:
        return "5,001–10,000"
    if n <= 50000:
        return "10,001–50,000"
    return "50,001+"


def parse_employees_from_text(text: str) -> int | None:
    if not text:
        return None
    patterns = [
        r"company size[:\s]+([\d,]+)\s*[–\-]\s*([\d,]+)",
        r"([\d,]+)\s*[–\-]\s*([\d,]+)\s+employees",
        r"employs\s+([\d,]+)\s+people",
        r"([\d,]+)\s+employees",
        r"approximately\s+([\d,]+)\s+employees",
        r"has\s+([\d,]+)\s+employees",
        r"employee count[:\s]+([\d,]+)",
        r"(\d{1,3})\s*-\s*(\d{1,3})\s+employees",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if not m:
            continue
        if m.lastindex and m.lastindex >= 2 and m.group(2) and re.search(r"[–\-]", m.group(0)):
            try:
                lo = int(m.group(1).replace(",", ""))
                hi = int(m.group(2).replace(",", ""))
                if 1 <= lo <= hi <= 2_000_000:
                    return (lo + hi) // 2
            except ValueError:
                pass
        try:
            n = int(m.group(1).replace(",", ""))
        except ValueError:
            continue
        if 1 <= n <= 2_000_000 and n not in range(1900, 2031):
            return n
    band_map = [
        (r"1[-–]10\s+employees", 5),
        (r"11[-–]50\s+employees", 30),
        (r"51[-–]200\s+employees", 100),
        (r"201[-–]500\s+employees", 350),
        (r"501[-–]1,?000\s+employees", 750),
        (r"1,?001[-–]5,?000\s+employees", 3000),
        (r"5,?001[-–]10,?000\s+employees", 7500),
        (r"10,?001[-–]50,?000\s+employees", 25000),
        (r"50,?001\+\s+employees", 60000),
    ]
    for pat, mid in band_map:
        if re.search(pat, text, re.I):
            return mid
    return None


def ddg_snippets(query: str) -> str:
    url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(query)
    try:
        with httpx.Client(
            headers={"User-Agent": UA},
            timeout=15.0,
            follow_redirects=True,
        ) as client:
            r = client.get(url)
            if r.status_code >= 400:
                return ""
            # strip tags lightly
            t = re.sub(r"(?is)<script.*?>.*?</script>", " ", r.text or "")
            t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
            # keep result snippets
            parts = re.findall(r'class="result__snippet[^"]*"[^>]*>(.*?)</(?:a|td|div)', t, re.I | re.S)
            if not parts:
                parts = re.findall(r"result__snippet.*?>(.*?)<", t, re.I | re.S)
            plain = []
            for p in parts[:8]:
                p = re.sub(r"<[^>]+>", " ", p)
                p = re.sub(r"\s+", " ", p).strip()
                if p:
                    plain.append(p)
            # also grab title lines mentioning employees
            titles = re.findall(r'class="result__a[^"]*"[^>]*>(.*?)</a>', t, re.I | re.S)
            for tit in titles[:8]:
                tit = re.sub(r"<[^>]+>", " ", tit)
                tit = re.sub(r"\s+", " ", tit).strip()
                if tit:
                    plain.append(tit)
            return "\n".join(plain)[:5000]
    except Exception:
        return ""


def deepseek_extract(client: OpenAI | None, model: str, name: str, web: str, evidence: str) -> str:
    if not client or not evidence or len(evidence) < 40:
        return ""
    system = (
        "Extract employee headcount for THIS company only from EVIDENCE. Never invent. "
        "Ignore other similarly named companies if domain/location clearly differs. "
        f'JSON {{"range":"","confidence":"high|medium|low|none"}}. range one of: {", ".join(RANGES)} or "".'
    )
    user = f"Company: {name}\nWebsite: {web}\n\nEVIDENCE:\n{evidence[:5500]}"
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=120,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = json.loads((resp.choices[0].message.content or "{}").strip())
        rng = str(data.get("range") or "").strip().replace("-", "–")
        conf = str(data.get("confidence") or "").lower()
        for r in RANGES:
            if rng.replace(",", "") == r.replace(",", "") or rng == r:
                rng = r
                break
        else:
            n = parse_employees_from_text(rng)
            rng = emp_num_to_range(n) if n else ""
        if conf in {"high", "medium"} and rng in RANGES:
            return rng
    except Exception:
        return ""
    return ""


def main() -> None:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    client = (
        OpenAI(
            api_key=key,
            base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/"),
            timeout=60.0,
        )
        if key
        else None
    )
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    cache = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    wb = load_workbook(PATH)
    ws = wb.active
    h = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ic = {n: h.index(n) + 1 for n in h}

    # Apply HAND2 first
    hand_n = 0
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, ic["Company"]).value or "").strip()
        if name in HAND2:
            ws.cell(r, ic["Employees"]).value = HAND2[name]
            hand_n += 1

    gaps = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, ic["Company"]).value or "").strip()
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        cur = str(ws.cell(r, ic["Employees"]).value or "").strip()
        if emptyish(cur) or cur == NPD:
            gaps.append((r, name, web))

    print(f"HAND2 applied={hand_n}; gaps={len(gaps)}")

    def work(item):
        r, name, web = item
        ck = f"{name}|{web}".lower()
        if ck in cache and cache[ck].get("range"):
            return r, name, cache[ck]["range"], cache[ck].get("source", "cache")

        q1 = f"{name} LinkedIn employees company size"
        q2 = f"{name} {web.replace('https://','').replace('http://','').split('/')[0]} employees"
        evidence = ddg_snippets(q1)
        if len(evidence) < 60:
            evidence = (evidence + "\n" + ddg_snippets(q2)).strip()
        n = parse_employees_from_text(evidence)
        if n:
            rng = emp_num_to_range(n)
            cache[ck] = {"range": rng, "source": "ddg_regex", "n": n}
            return r, name, rng, "ddg_regex"
        rng = deepseek_extract(client, model, name, web, evidence) if evidence else ""
        if rng:
            cache[ck] = {"range": rng, "source": "ddg_deepseek"}
            time.sleep(0.12)
            return r, name, rng, "ddg_deepseek"
        cache[ck] = {"range": "", "source": "none"}
        return r, name, "", "none"

    filled = 0
    sources = {"ddg_regex": 0, "ddg_deepseek": 0, "none": 0, "cache": 0}
    audit = []

    # unique companies
    uniq = {}
    for r, name, web in gaps:
        uniq[name] = (r, name, web)
    # but multiple rows? unlikely; process unique then map
    results = {}
    items = list({(name, web): (r, name, web) for r, name, web in gaps}.values())

    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = [ex.submit(work, it) for it in items]
        done = 0
        for fut in as_completed(futs):
            r, name, rng, src = fut.result()
            results[name] = (rng, src)
            done += 1
            if done % 30 == 0:
                print(f"  progress {done}/{len(items)}")
                CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    for r, name, web in gaps:
        rng, src = results.get(name, ("", "none"))
        if name in HAND2:
            continue  # already set
        if rng:
            ws.cell(r, ic["Employees"]).value = rng
            filled += 1
            sources[src] = sources.get(src, 0) + 1
        else:
            ws.cell(r, ic["Employees"]).value = NPD
            sources["none"] += 1
        audit.append({"company": name, "employees": rng or NPD, "source": src})

    wb.save(PATH)
    wb.close()

    wb = load_workbook(PATH, data_only=True)
    ws = wb.active
    ei = h.index("Employees")
    named = blank = 0
    dist = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        v = str(row[ei] or "").strip()
        if emptyish(v) or v == NPD:
            blank += 1
        else:
            named += 1
            dist[v] = dist.get(v, 0) + 1
    summary = {
        "pass2_filled": filled,
        "sources": sources,
        "final_named": named,
        "final_npd": blank,
        "distribution": dict(sorted(dist.items(), key=lambda x: -x[1])),
    }
    AUDIT.write_text(json.dumps({"summary": summary, "rows": audit}, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    wb.close()


if __name__ == "__main__":
    main()
