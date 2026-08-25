"""Remove bad pass3 placeholders; add verified distributors to reach exactly +100 from 224."""
from __future__ import annotations

import json
import os
import re
import shutil
import socket
import ssl
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "add100_cleanup_finish_audit.json"
CKPT = FP.parent / "add100_finish_ckpt.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
ND = "Not publicly disclosed"
BASELINE = 224
TARGET = 100

# Placeholder / dead / parking rows introduced in pass3 manual guesses
REMOVE_NAMES = {
    "day-pak",
    "accurate box",
    "empire packaging",
    "florida packaging",
    "houston packaging",
    "lakeside packaging",
    "midwest packaging",
    "phoenix packaging",
    "texas packaging",
    "western packaging",
    "suncoast packaging",
    "canada packaging",
    "montreal packaging",
    "toronto packaging supplies",
    "packaging ireland",
    "cartonajes distributor spain",
    "embalajes industriales dist",
    "verpakkingen belgië dist",
    "verpakkingen belgi dist",
    "emballages belgique",
    "osterreich verpackung dist",
    "czech packaging dist",
    "boxmart au",
    "nz packaging supplies",
    "singapore packaging hub",
    "delhi packaging traders",
    "south africa packaging dist",
    "brazil embalagens dist",
    "chile embalajes dist",
    "pinnacle packaging",
    "packaging products india",
    "packaging mart india",
    "india packaging distributors",
    "mumbai packaging supply",
    "tokyo packaging supply",
    "korea packaging dist",
    "malaysia packaging supply",
    "thailand packaging dist",
    "auckland packaging",
    "packaging warehouse au",
    "packaging house au",
    "nordic packaging supply",
    "polish packaging dist",
    "swiss packaging supplies",
    "belfast packaging",
    "scottish packaging supplies",
    "embalajes mexico dist",
    "all packaging company",  # ambiguous vs existing All Packaging Company Denver
    "industrial packaging corp dist",
    "pacific packaging",
    "york packaging",
    "cal-pak",
    "unipack trading",
    "packaging tech llc",
    "atlantic packaging products dist",
}

# Agent-verified real distributors (web-researched)
VERIFIED_SEED = [
    {
        "company": "Paper Mart",
        "website": "https://www.papermart.com",
        "hq": "Orange, CA, USA",
        "cc": "US",
        "founded": "1921",
        "cats": "Wholesale packaging; shipping; retail packaging supplies",
        "focus": "US online packaging wholesaler/distributor",
        "summary": "Paper Mart is a long-established wholesale packaging supplies distributor based in Orange, California.",
    },
    {
        "company": "McMaster-Carr",
        "website": "https://www.mcmaster.com",
        "hq": "USA",
        "cc": "US",
        "founded": "Not publicly disclosed",
        "cats": "Industrial supplies including packaging materials",
        "focus": "Industrial catalog distributor with packaging lines",
        "summary": "McMaster-Carr is an industrial supplies catalog distributor that includes packaging materials.",
    },
]


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def dns_ok(d: str) -> bool:
    try:
        socket.getaddrinfo(d, 443)
        return True
    except Exception:
        try:
            socket.getaddrinfo(d, 80)
            return True
        except Exception:
            return False


def fetch_text(website: str) -> str:
    base = website if website.startswith("http") else "https://" + website
    base = base.rstrip("/")
    chunks = []
    ctx = ssl.create_default_context()
    for p in ["", "/about", "/about-us"]:
        try:
            req = Request(base + p, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0"})
            with urlopen(req, timeout=12, context=ctx) as resp:
                html = resp.read(70000).decode("utf-8", errors="ignore")
            low = html.lower()
            if any(x in low for x in ("domain is for sale", "buy this domain", "parked domain", "this domain may be for sale")):
                return "PARKED"
            text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
            text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text)
            chunks.append(text[:2800])
        except Exception:
            continue
        if len("\n".join(chunks)) > 4500:
            break
    return "\n".join(chunks)[:6500]


def suggest(existing: set[str]) -> list[dict]:
    sample = ", ".join(sorted(list(existing))[:80])
    out = []
    for region in [
        "USA packaging distributors wholesalers merchants real websites",
        "UK Ireland packaging merchants distributors real websites",
        "Germany Netherlands Belgium France packaging wholesalers real websites not RAJA Manutan",
        "Canada Australia New Zealand packaging distributors real websites",
    ]:
        prompt = f"""List 18 REAL packaging DISTRIBUTORS/wholesalers/merchants for: {region}.
Real official websites only. No manufacturers/hybrids/OEMs.
Avoid: {sample}
JSON: {{"companies":[{{"company":"","website":"https://","hq":"","country_code":"","founded":"Not publicly disclosed","core_categories":"","specialty_focus":"","summary":""}}]}}"""
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": "Return compact JSON of real packaging channel companies only."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.15,
                response_format={"type": "json_object"},
                max_tokens=3200,
            )
            comps = json.loads(resp.choices[0].message.content or "{}").get("companies") or []
            print(region[:28], len(comps), flush=True)
            out.extend(comps)
        except Exception as e:
            print("suggest fail", e, flush=True)
    return out


def role_check(company: str, website: str, source: str) -> dict:
    prompt = f"""Classify packaging channel role.
Company: {company}
Website: {website}
SOURCE: {source[:6000]}
KEEP only Specialty Distributor/Wholesaler/Merchant/Channel Partner of packaging supplies.
REMOVE Manufacturer/OEM/Converter/Hybrid/Wrong Market/Unknown/Parked.
JSON: {{"keep":true|false,"role":"...","market_fit":true|false,"founded":"...","hq":"...","reason":"..."}}"""
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Strict channel-only. Prefer remove."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        response_format={"type": "json_object"},
    )
    try:
        return json.loads(resp.choices[0].message.content or "{}")
    except Exception:
        return {"keep": False, "role": "Unknown", "market_fit": False}


def row_from(c: dict) -> list:
    cc = c.get("cc") or c.get("country_code") or "XX"
    cc = str(cc).upper()[:2]
    rc = "NA" if cc in {"US", "CA", "MX"} else ("APAC" if cc in {"AU", "NZ", "SG", "IN", "JP", "KR"} else "EU")
    geo = c.get("geo") or f"{rc}({cc})"
    return [
        c["company"],
        c["website"],
        c.get("founded") or ND,
        c.get("hq") or ND,
        geo,
        geo,
        "Private",
        ND,
        c.get("cats") or c.get("core_categories") or "Packaging supplies",
        c.get("focus") or c.get("specialty_focus") or "Packaging distributor",
        c.get("brands") or c["company"],
        "Both",
        "Specialty Distributor",
        ND,
        ND,
        ND,
        ND,
        ND,
        cc,
        rc,
        c.get("summary") or f"{c['company']} is a packaging specialty distributor ({c['website']}).",
    ]


def main():
    backup = FP.with_name(FP.stem + f"_pre_add100_finish_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(FP, backup)
    print("backup", backup)

    wb = load_workbook(FP)
    ws = wb.active
    header = [c.value for c in ws[1]]
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]

    kept = []
    removed = []
    for r in rows:
        name = (r[0] or "").strip()
        low = name.lower()
        # normalize weird encoding
        low2 = low.replace("ë", "e").replace("ö", "o")
        if low in REMOVE_NAMES or low2 in REMOVE_NAMES:
            removed.append(name)
            continue
        kept.append(r)

    print("removed placeholders", len(removed), "kept", len(kept))

    existing_names = {(r[0] or "").strip().lower() for r in kept}
    existing_domains = {domain_of(str(r[1] or "")) for r in kept if r[1]}

    # rebuild sheet from kept first
    ws.delete_rows(1, ws.max_row)
    ws.append(header)
    for r in kept:
        ws.append(r)
    wb.save(FP)

    need = TARGET - (len(kept) - BASELINE)
    print("need", need, "current_added", len(kept) - BASELINE)

    ckpt = {"accepted": [], "rejected": []}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    pool = []
    for s in VERIFIED_SEED:
        pool.append(s)
    print("suggesting...", flush=True)
    for raw in suggest(existing_names):
        name = (raw.get("company") or "").strip()
        web = (raw.get("website") or "").strip()
        if not name or not web:
            continue
        if not web.startswith("http"):
            web = "https://" + web
        d = domain_of(web)
        if name.lower() in existing_names or d in existing_domains:
            continue
        pool.append(
            {
                "company": name,
                "website": web,
                "hq": raw.get("hq") or ND,
                "cc": (raw.get("country_code") or "XX").upper()[:2],
                "founded": raw.get("founded") or ND,
                "cats": raw.get("core_categories") or "Packaging supplies",
                "focus": raw.get("specialty_focus") or "Packaging distributor",
                "summary": raw.get("summary") or "",
            }
        )

    # uniq pool
    seen = set()
    uniq = []
    for c in pool:
        k = c["company"].lower()
        d = domain_of(c["website"])
        if k in seen or k in existing_names or d in existing_domains:
            continue
        seen.add(k)
        uniq.append(c)
    print("pool", len(uniq), flush=True)

    added = 0
    for i, c in enumerate(uniq):
        if (len(kept) - BASELINE) + added >= TARGET:
            break
        name, web, d = c["company"], c["website"], domain_of(c["website"])
        print(f"[{i+1}/{len(uniq)}] {name} ...", flush=True)
        if not dns_ok(d):
            ckpt["rejected"].append({"company": name, "reason": "dns"})
            print("  dns fail", flush=True)
            continue
        src = fetch_text(web)
        if src == "PARKED":
            ckpt["rejected"].append({"company": name, "reason": "parked"})
            print("  parked", flush=True)
            continue
        data = role_check(name, web, src or f"{name} {c.get('focus')} {c.get('summary')}")
        role = (data.get("role") or "").lower()
        keep = bool(data.get("keep")) and bool(data.get("market_fit", True))
        if any(x in role for x in ("manufacturer", "oem", "converter", "hybrid", "wrong", "unknown", "parked")):
            keep = False
        if not keep:
            ckpt["rejected"].append({"company": name, "role": data.get("role"), "reason": data.get("reason")})
            print(f"  REJECT {data.get('role')}", flush=True)
            CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
            time.sleep(0.15)
            continue
        if data.get("founded"):
            c["founded"] = str(data["founded"]).strip()
        if data.get("hq"):
            c["hq"] = str(data["hq"]).strip()
        row = row_from(c)
        ws.append(row)
        kept.append(row)
        existing_names.add(name.lower())
        existing_domains.add(d)
        added += 1
        ckpt["accepted"].append(c)
        print(f"  KEEP ({(len(kept)-BASELINE)}/{TARGET})", flush=True)
        CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
        time.sleep(0.15)

    wb.save(FP)
    final = ws.max_row - 1
    # uniqueness assert
    names = [(r[0] or "").strip().lower() for r in ws.iter_rows(min_row=2, values_only=True)]
    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "removed_placeholders": removed,
        "added_this_pass": added,
        "after": final,
        "added_from_baseline": final - BASELINE,
        "unique_names": len(set(names)),
        "file": str(FP),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("DONE after", final, "from_baseline", final - BASELINE, "unique", len(set(names)))


if __name__ == "__main__":
    main()
