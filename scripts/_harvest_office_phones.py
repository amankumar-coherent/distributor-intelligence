# -*- coding: utf-8 -*-
"""Harvest public HQ/switchboard phones from company websites into Office No.

Never invents numbers. Prefer tel: / JSON-LD telephone. Atomic Excel save.
"""
from __future__ import annotations

import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urljoin, urlparse

import httpx
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel")
CACHE = ROOT / "scripts" / "_office_phone_cache.json"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

PATHS = [
    Path(r"D:\vendor-intel\output\web_expand\global_electrical_distribution_market_global\global_electrical_distribution_market_global_FINAL.xlsx"),
    Path(r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market.xlsx"),
    Path(r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market.xlsx"),
    Path(r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market.xlsx"),
    Path(r"D:\vendor-intel\output\web_expand\global_smartphone_market_global\global_smartphone_market.xlsx"),
    Path(r"D:\vendor-intel\output\web_expand\global_antibody_drug_conjugates_market_global\global_antibody_drug_conjugates_market.xlsx"),
    Path(r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market.xlsx"),
]

NPD = {
    "", "nan", "none", "n/a", "na", "not publicly disclosed",
    "-", "—", "unknown", "tbd", "null",
}

CONTACT_PATHS = (
    "",
    "/contact",
    "/contact-us",
)

TEL_HREF = re.compile(r"""href\s*=\s*["']tel:([^"']+)["']""", re.I)
JSON_TEL = re.compile(r"""["']telephone["']\s*:\s*["']([^"']+)["']""", re.I)
ITEM_TEL = re.compile(r"""itemprop\s*=\s*["']telephone["'][^>]*>([^<]+)""", re.I)
CONTENT_TEL = re.compile(
    r"""(?:itemprop|property)\s*=\s*["'](?:telephone|og:phone_number)["'][^>]*content\s*=\s*["']([^"']+)["']""",
    re.I,
)
PLAIN = re.compile(
    r"(?:(?:Tel|Phone|Telephone|Call|Téléphone|Teléfono|Telefon|Office|HQ|Switchboard|"
    r"Customer Service|Sales)[:\s]+)"
    r"(\+?\d[\d\s()./-]{7,22}\d)",
    re.I,
)
BARE_PLUS = re.compile(r"(\+\d{1,3}[\s().-]?\d[\d\s().-]{6,18}\d)")

JUNK_DIGIT_RUNS = re.compile(r"(?:19|20)\d{2}")  # years often false-positive


def load_cache() -> dict:
    if CACHE.exists():
        try:
            return json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_cache(cache: dict) -> None:
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")


def has_val(s: str) -> bool:
    return bool(s) and s.strip().lower() not in NPD


def digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def normalize_phone(raw: str) -> str:
    if not raw:
        return ""
    s = raw.strip()
    s = s.replace("tel:", "").replace("TEL:", "")
    s = re.sub(r"^/+", "", s)
    s = s.replace("%20", " ").replace("&nbsp;", " ")
    s = re.sub(r"\s+", " ", s).strip(" \t.;,")
    # drop extensions clutter for office switchboard if very long
    d = digits(s)
    if len(d) < 8 or len(d) > 15:
        return ""
    # reject likely years / order ids
    if re.fullmatch(r"20\d{2}", d) or re.fullmatch(r"19\d{2}", d):
        return ""
    if d.startswith("000") or len(set(d)) <= 2:
        return ""
    return s


def pick_best(candidates: list[str]) -> str:
    scored = []
    seen = set()
    for c in candidates:
        p = normalize_phone(c)
        if not p:
            continue
        key = digits(p)
        if key in seen:
            continue
        seen.add(key)
        score = 0
        if p.startswith("+"):
            score += 5
        if 10 <= len(key) <= 13:
            score += 3
        if 8 <= len(key) <= 15:
            score += 1
        scored.append((score, p))
    if not scored:
        return ""
    scored.sort(key=lambda x: (-x[0], len(x[1])))
    return scored[0][1]


def extract_phones(html: str) -> str:
    if not html:
        return ""
    cands: list[str] = []
    for rx in (TEL_HREF, JSON_TEL, ITEM_TEL, CONTENT_TEL):
        cands.extend(rx.findall(html))
    # tel: in javascript
    cands.extend(re.findall(r"tel:([+\d][\d\s().-]{7,22}\d)", html, re.I))
    cands.extend(PLAIN.findall(html))
    # plus-international in footer-ish last 80k chars
    tail = html[-80000:] if len(html) > 80000 else html
    cands.extend(BARE_PLUS.findall(tail))
    return pick_best(cands)


def abs_url(website: str) -> str:
    w = (website or "").strip()
    if not has_val(w):
        return ""
    if w.lower().startswith("mailto:"):
        return ""
    if not re.match(r"^https?://", w, re.I):
        w = "https://" + w.lstrip("/")
    return w.rstrip("/")


def harvest_one(website: str) -> str:
    base = abs_url(website)
    if not base:
        return ""
    parsed = urlparse(base)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    urls = []
    # homepage first
    urls.append(base if parsed.path not in ("", "/") else origin + "/")
    for p in CONTACT_PATHS:
        if not p:
            continue
        urls.append(urljoin(origin + "/", p.lstrip("/")))
    # unique preserve order
    seen = set()
    uniq = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            uniq.append(u)

    headers = {"User-Agent": UA, "Accept": "text/html,application/xhtml+xml"}
    try:
        with httpx.Client(
            follow_redirects=True,
            timeout=httpx.Timeout(10.0, connect=6.0),
            headers=headers,
            verify=True,
        ) as client:
            for url in uniq[:3]:
                try:
                    resp = client.get(url)
                except Exception:
                    continue
                if resp.status_code >= 400:
                    continue
                ctype = (resp.headers.get("content-type") or "").lower()
                if "text" not in ctype and "json" not in ctype and ctype:
                    continue
                phone = extract_phones(resp.text or "")
                if phone:
                    return phone
    except Exception:
        return ""
    return ""


def iter_excel_rows(path: Path):
    wb = load_workbook(path, data_only=True)
    ws = wb["Companies"] if "Companies" in wb.sheetnames else wb.active
    headers = None
    hr = 1
    for r in range(1, 6):
        vals = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=r, max_row=r))]
        if "Company" in vals:
            headers = vals
            hr = r
            break
    if not headers or "Office No." not in headers:
        wb.close()
        raise RuntimeError(f"Office No. missing in {path}")
    co_i = headers.index("Company")
    web_i = headers.index("Website") if "Website" in headers else None
    off_i = headers.index("Office No.")
    rows = []
    for r in range(hr + 1, ws.max_row + 1):
        co = ws.cell(r, co_i + 1).value
        if not co or not str(co).strip():
            continue
        web = str(ws.cell(r, web_i + 1).value or "").strip() if web_i is not None else ""
        off = str(ws.cell(r, off_i + 1).value or "").strip()
        rows.append({"row": r, "company": str(co).strip(), "website": web, "office": off})
    wb.close()
    return headers, hr, rows


def apply_excel(path: Path, fills: dict[str, str]) -> int:
    """fills keyed by exact company name. Only write empty Office No."""
    wb = load_workbook(path)
    ws = wb["Companies"] if "Companies" in wb.sheetnames else wb.active
    headers = None
    hr = 1
    for r in range(1, 6):
        vals = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=r, max_row=r))]
        if "Company" in vals:
            headers = vals
            hr = r
            break
    co_i = headers.index("Company") + 1
    off_i = headers.index("Office No.") + 1
    n = 0
    for r in range(hr + 1, ws.max_row + 1):
        co = str(ws.cell(r, co_i).value or "").strip()
        if not co:
            continue
        cur = str(ws.cell(r, off_i).value or "").strip()
        if has_val(cur):
            continue
        phone = fills.get(co) or ""
        if not phone:
            continue
        ws.cell(r, off_i).value = phone
        n += 1
    tmp = path.with_suffix(".xlsx.tmp")
    wb.save(tmp)
    wb.close()
    tmp.replace(path)
    return n


def main() -> None:
    cache = load_cache()
    # collect unique websites needing phones
    todo: list[tuple[str, str]] = []  # company, website
    company_web: dict[str, str] = {}
    for p in PATHS:
        if not p.exists():
            print("MISSING", p)
            continue
        _, _, rows = iter_excel_rows(p)
        empty = 0
        for row in rows:
            if has_val(row["office"]):
                continue
            empty += 1
            web = row["website"]
            if not abs_url(web):
                continue
            company_web[row["company"]] = web
        print(f"SCAN {p.name} empty_office={empty}")

    # harvest missing cache keys by website
    sites = []
    web_to_cos: dict[str, list[str]] = {}
    for co, web in company_web.items():
        key = abs_url(web).lower()
        web_to_cos.setdefault(key, []).append(co)
        if key not in cache:
            sites.append(web)

    # unique websites
    uniq_sites = []
    seenw = set()
    for w in sites:
        k = abs_url(w).lower()
        if k in seenw:
            continue
        seenw.add(k)
        uniq_sites.append(w)

    print(f"HARVEST unique websites={len(uniq_sites)} cached={sum(1 for v in cache.values() if v)}")
    workers = 32
    done = 0
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(harvest_one, w): w for w in uniq_sites}
        for fut in as_completed(futs):
            w = futs[fut]
            key = abs_url(w).lower()
            try:
                phone = fut.result() or ""
            except Exception:
                phone = ""
            cache[key] = phone
            done += 1
            if done % 40 == 0 or done == len(uniq_sites):
                save_cache(cache)
                hit = sum(1 for v in cache.values() if v)
                print(f"  progress {done}/{len(uniq_sites)} hits={hit} elapsed={time.time()-t0:.0f}s")

    save_cache(cache)

    # map company -> phone
    fills: dict[str, str] = {}
    for co, web in company_web.items():
        phone = cache.get(abs_url(web).lower()) or ""
        if phone:
            fills[co] = phone

    print(f"COMPANY FILLS ready={len(fills)}")
    for p in PATHS:
        if not p.exists():
            continue
        n = apply_excel(p, fills)
        print(f"WRITE {p.name} filled={n}")


if __name__ == "__main__":
    main()
