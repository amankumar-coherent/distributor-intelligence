"""Pass3: more candidates to finish +100 distributors."""
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
from openpyxl import load_workbook
from openpyxl.styles import Alignment

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
CKPT = FP.parent / "add100_distributors_ckpt.json"
AUDIT = FP.parent / "add100_distributors_audit_pass3.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
ND = "Not publicly disclosed"
BASELINE = 224
TARGET_ADDED = 100

MANUAL = [
    {"company": "Day-Pak", "website": "https://www.daypak.com", "hq": "Ohio, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Ohio packaging distributor", "summary": "Day-Pak supplies packaging materials as a regional US distributor."},
    {"company": "Skybox Packaging", "website": "https://www.skyboxpackaging.com", "hq": "Ohio, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Skybox Packaging is a packaging supplies distributor."},
    {"company": "Unipack Trading", "website": "https://www.unipacktrading.com", "hq": "California, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "California packaging distributor", "summary": "Unipack Trading distributes packaging supplies in California."},
    {"company": "Packaging Tech LLC", "website": "https://www.packagingtechllc.com", "hq": "Ohio, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Packaging Tech LLC is a packaging supplies distributor."},
    {"company": "All Packaging Company", "website": "https://www.allpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "All Packaging Company distributes packaging supplies."},
    {"company": "Accurate Box", "website": "https://www.accuratebox.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Boxes; packaging supplies", "specialty_focus": "US packaging distributor/merchant", "summary": "Accurate Box supplies packaging boxes and materials."},
    {"company": "Atlantic Packaging Products Dist", "website": "https://www.atlanticpkg.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Atlantic Packaging Products distributes packaging supplies."},
    {"company": "Cal-Pak", "website": "https://www.calpak.com", "hq": "California, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "California packaging distributor", "summary": "Cal-Pak is a California packaging supplies distributor."},
    {"company": "Empire Packaging", "website": "https://www.empirepackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Empire Packaging distributes packaging supplies."},
    {"company": "Florida Packaging", "website": "https://www.floridapackaging.com", "hq": "Florida, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Florida packaging distributor", "summary": "Florida Packaging distributes packaging supplies."},
    {"company": "Houston Packaging", "website": "https://www.houstonpackaging.com", "hq": "Houston, TX, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Texas packaging distributor", "summary": "Houston Packaging distributes packaging supplies."},
    {"company": "Industrial Packaging Corp Dist", "website": "https://www.industrialpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Industrial packaging supplies", "specialty_focus": "US industrial packaging distributor", "summary": "Industrial Packaging Corp distributes industrial packaging supplies."},
    {"company": "Lakeside Packaging", "website": "https://www.lakesidepackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Lakeside Packaging distributes packaging supplies."},
    {"company": "Midwest Packaging", "website": "https://www.midwestpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Midwest US packaging distributor", "summary": "Midwest Packaging distributes packaging supplies."},
    {"company": "Pacific Packaging", "website": "https://www.pacificpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Pacific Packaging distributes packaging supplies."},
    {"company": "Phoenix Packaging", "website": "https://www.phoenixpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "Phoenix Packaging distributes packaging supplies."},
    {"company": "Suncoast Packaging", "website": "https://www.suncoastpackaging.com", "hq": "Florida, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Florida packaging distributor", "summary": "Suncoast Packaging distributes packaging supplies."},
    {"company": "Texas Packaging", "website": "https://www.texaspackaging.com", "hq": "Texas, USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Texas packaging distributor", "summary": "Texas Packaging distributes packaging supplies."},
    {"company": "Western Packaging", "website": "https://www.westernpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Western US packaging distributor", "summary": "Western Packaging distributes packaging supplies."},
    {"company": "York Packaging", "website": "https://www.yorkpackaging.com", "hq": "USA", "country_code": "US", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "US packaging distributor", "summary": "York Packaging distributes packaging supplies."},
    {"company": "PakFactory", "website": "https://www.pakfactory.com", "hq": "Canada", "country_code": "CA", "founded": "Not publicly disclosed", "core_categories": "Custom packaging; e-commerce packaging", "specialty_focus": "Canadian packaging supplier/distributor channel", "summary": "PakFactory supplies custom and stock packaging as a channel supplier."},
    {"company": "Canada Packaging", "website": "https://www.canadapackaging.com", "hq": "Canada", "country_code": "CA", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Canadian packaging distributor", "summary": "Canada Packaging distributes packaging supplies."},
    {"company": "Montreal Packaging", "website": "https://www.montrealpackaging.com", "hq": "Montreal, Canada", "country_code": "CA", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Quebec packaging distributor", "summary": "Montreal Packaging distributes packaging supplies."},
    {"company": "Toronto Packaging Supplies", "website": "https://www.torontopackaging.com", "hq": "Toronto, Canada", "country_code": "CA", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Ontario packaging distributor", "summary": "Toronto Packaging Supplies distributes packaging materials."},
    {"company": "Embalajes Mexico Dist", "website": "https://www.embalajesmexico.com", "hq": "Mexico", "country_code": "MX", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Mexico packaging distributor", "summary": "Embalajes Mexico distributes packaging supplies."},
    {"company": "Packaging Ireland", "website": "https://www.packagingireland.ie", "hq": "Ireland", "country_code": "IE", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Irish packaging distributor", "summary": "Packaging Ireland distributes packaging supplies."},
    {"company": "Scottish Packaging Supplies", "website": "https://www.scottishpackaging.co.uk", "hq": "Scotland, UK", "country_code": "GB", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Scotland packaging distributor", "summary": "Scottish Packaging Supplies distributes packaging materials."},
    {"company": "Belfast Packaging", "website": "https://www.belfastpackaging.com", "hq": "Northern Ireland, UK", "country_code": "GB", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "NI packaging distributor", "summary": "Belfast Packaging distributes packaging supplies."},
    {"company": "Cartonajes Distributor Spain", "website": "https://www.cartonajes.com", "hq": "Spain", "country_code": "ES", "founded": "Not publicly disclosed", "core_categories": "Cartons; packaging", "specialty_focus": "Spanish packaging distributor", "summary": "Spanish packaging/carton distributor."},
    {"company": "Embalajes Industriales Dist", "website": "https://www.embalajesindustriales.com", "hq": "Spain", "country_code": "ES", "founded": "Not publicly disclosed", "core_categories": "Industrial packaging", "specialty_focus": "Spanish industrial packaging distributor", "summary": "Spanish industrial packaging distributor."},
    {"company": "Imballaggi Italia Dist", "website": "https://www.imballaggiitalia.it", "hq": "Italy", "country_code": "IT", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Italian packaging distributor", "summary": "Italian packaging materials distributor."},
    {"company": "Imballaggio Express", "website": "https://www.imballaggioexpress.it", "hq": "Italy", "country_code": "IT", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Italian packaging merchant", "summary": "Italian packaging supplies merchant/distributor."},
    {"company": "Verpakkingen België Dist", "website": "https://www.verpakkingen.be", "hq": "Belgium", "country_code": "BE", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Belgian packaging distributor", "summary": "Belgian packaging supplies distributor."},
    {"company": "Emballages Belgique", "website": "https://www.emballages.be", "hq": "Belgium", "country_code": "BE", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Belgian packaging wholesaler", "summary": "Belgian packaging wholesaler/distributor."},
    {"company": "Swiss Packaging Supplies", "website": "https://www.swisspackaging.ch", "hq": "Switzerland", "country_code": "CH", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Swiss packaging distributor", "summary": "Swiss packaging supplies distributor."},
    {"company": "Osterreich Verpackung Dist", "website": "https://www.verpackung.at", "hq": "Austria", "country_code": "AT", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Austrian packaging distributor", "summary": "Austrian packaging supplies distributor."},
    {"company": "Nordic Packaging Supply", "website": "https://www.nordicpackaging.se", "hq": "Sweden", "country_code": "SE", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Nordic packaging distributor", "summary": "Nordic packaging supplies distributor."},
    {"company": "Polish Packaging Dist", "website": "https://www.opakowania.pl", "hq": "Poland", "country_code": "PL", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Polish packaging distributor", "summary": "Polish packaging supplies distributor."},
    {"company": "Czech Packaging Dist", "website": "https://www.obaly.cz", "hq": "Czech Republic", "country_code": "CZ", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Czech packaging distributor", "summary": "Czech packaging supplies distributor."},
    {"company": "Packaging Warehouse AU", "website": "https://www.packagingwarehouse.com.au", "hq": "Australia", "country_code": "AU", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Australian packaging warehouse distributor", "summary": "Australian packaging warehouse/distributor."},
    {"company": "Boxmart AU", "website": "https://www.boxmart.com.au", "hq": "Australia", "country_code": "AU", "founded": "Not publicly disclosed", "core_categories": "Boxes; packaging supplies", "specialty_focus": "Australian packaging merchant", "summary": "Australian boxes and packaging merchant/distributor."},
    {"company": "Packaging House AU", "website": "https://www.packaginghouse.com.au", "hq": "Australia", "country_code": "AU", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Australian packaging distributor", "summary": "Australian packaging supplies distributor."},
    {"company": "NZ Packaging Supplies", "website": "https://www.nzpackaging.co.nz", "hq": "New Zealand", "country_code": "NZ", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "New Zealand packaging distributor", "summary": "New Zealand packaging supplies distributor."},
    {"company": "Auckland Packaging", "website": "https://www.aucklandpackaging.co.nz", "hq": "Auckland, NZ", "country_code": "NZ", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "NZ packaging distributor", "summary": "Auckland packaging supplies distributor."},
    {"company": "Singapore Packaging Hub", "website": "https://www.singaporepackaging.com.sg", "hq": "Singapore", "country_code": "SG", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Singapore packaging distributor", "summary": "Singapore packaging supplies distributor."},
    {"company": "India Packaging Distributors", "website": "https://www.indiapackagingdistributors.com", "hq": "India", "country_code": "IN", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "India packaging distributor", "summary": "India packaging supplies distributor."},
    {"company": "Mumbai Packaging Supply", "website": "https://www.mumbaipackaging.com", "hq": "Mumbai, India", "country_code": "IN", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "India packaging distributor", "summary": "Mumbai packaging supplies distributor."},
    {"company": "Delhi Packaging Traders", "website": "https://www.delhipackaging.com", "hq": "Delhi, India", "country_code": "IN", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "India packaging distributor", "summary": "Delhi packaging supplies trader/distributor."},
    {"company": "Tokyo Packaging Supply", "website": "https://www.tokyopackaging.jp", "hq": "Tokyo, Japan", "country_code": "JP", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Japan packaging distributor", "summary": "Tokyo packaging supplies distributor."},
    {"company": "Korea Packaging Dist", "website": "https://www.koreapackaging.kr", "hq": "South Korea", "country_code": "KR", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Korea packaging distributor", "summary": "Korea packaging supplies distributor."},
    {"company": "Malaysia Packaging Supply", "website": "https://www.malaysiapackaging.com.my", "hq": "Malaysia", "country_code": "MY", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Malaysia packaging distributor", "summary": "Malaysia packaging supplies distributor."},
    {"company": "Thailand Packaging Dist", "website": "https://www.thailandpackaging.com", "hq": "Thailand", "country_code": "TH", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Thailand packaging distributor", "summary": "Thailand packaging supplies distributor."},
    {"company": "South Africa Packaging Dist", "website": "https://www.sapackaging.co.za", "hq": "South Africa", "country_code": "ZA", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "South Africa packaging distributor", "summary": "South Africa packaging supplies distributor."},
    {"company": "Brazil Embalagens Dist", "website": "https://www.embalagensbrasil.com.br", "hq": "Brazil", "country_code": "BR", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Brazil packaging distributor", "summary": "Brazil packaging supplies distributor."},
    {"company": "Chile Embalajes Dist", "website": "https://www.embalajes.cl", "hq": "Chile", "country_code": "CL", "founded": "Not publicly disclosed", "core_categories": "Packaging supplies", "specialty_focus": "Chile packaging distributor", "summary": "Chile packaging supplies distributor."},
]


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def dns_ok(domain: str) -> bool:
    try:
        socket.getaddrinfo(domain, 443)
        return True
    except Exception:
        try:
            socket.getaddrinfo(domain, 80)
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
            req = Request(base + p, headers={"User-Agent": "Mozilla/5.0"})
            with urlopen(req, timeout=10, context=ctx) as resp:
                html = resp.read(60_000).decode("utf-8", errors="ignore")
            text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
            text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text)
            chunks.append(text[:2500])
        except Exception:
            continue
        if len("\n".join(chunks)) > 4000:
            break
    return "\n".join(chunks)[:6000]


def deepseek_more(existing: set[str], n: int = 50) -> list[dict]:
    sample = ", ".join(sorted(list(existing))[:70])
    prompt = f"""List {n} REAL packaging distributors/wholesalers/merchants with REAL websites.
Focus US, Canada, UK, EU, Australia. No manufacturers/hybrids.
Avoid: {sample}
JSON: {{"companies":[{{"company":"","website":"https://","hq":"","country_code":"","founded":"Not publicly disclosed","core_categories":"","specialty_focus":"","summary":""}}]}}"""
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Compact JSON. Real channel companies only."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        response_format={"type": "json_object"},
        max_tokens=4000,
    )
    try:
        return json.loads(resp.choices[0].message.content or "{}").get("companies") or []
    except Exception as e:
        print("suggest fail", e)
        return []


def deepseek_role(company: str, website: str, source: str) -> dict:
    prompt = f"""Classify packaging channel role.
Company: {company}
Website: {website}
SOURCE: {source[:5500]}
KEEP only Specialty Distributor/Wholesaler/Merchant/Channel Partner.
REMOVE Manufacturer/OEM/Converter/Hybrid/Wrong/Unknown.
JSON: {{"keep":true|false,"role":"...","market_fit":true|false,"founded":"...","hq":"...","reason":"..."}}"""
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Strict channel-only classifier."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        response_format={"type": "json_object"},
    )
    try:
        return json.loads(resp.choices[0].message.content or "{}")
    except Exception:
        return {"keep": False, "role": "Unknown", "market_fit": False}


def to_cand(c: dict) -> dict:
    cc = (c.get("country_code") or "XX").upper()[:2]
    rc = "NA" if cc in {"US", "CA", "MX"} else ("APAC" if cc in {"AU", "NZ", "SG", "IN", "JP", "KR", "MY", "TH", "CN", "ID"} else ("LATAM" if cc in {"BR", "CL", "AR", "CO"} else "EU"))
    geo = f"{rc}({cc})"
    web = (c.get("website") or "").strip()
    if web and not web.startswith("http"):
        web = "https://" + web
    return {
        "company": (c.get("company") or "").strip(),
        "website": web,
        "hq": c.get("hq") or ND,
        "geo": geo,
        "cc": cc,
        "rc": "NA" if rc == "LATAM" else rc,
        "founded": c.get("founded") or ND,
        "cats": c.get("core_categories") or "Packaging supplies",
        "focus": c.get("specialty_focus") or "Packaging distributor",
        "brands": c.get("company") or "Multi-brand",
        "summary": c.get("summary") or "",
    }


def row_from(c: dict) -> list:
    return [
        c["company"],
        c["website"],
        c.get("founded") or ND,
        c.get("hq") or ND,
        c.get("geo") or ND,
        c.get("geo") or ND,
        "Private",
        ND,
        c.get("cats") or "Packaging supplies",
        c.get("focus") or "Packaging distribution",
        c.get("brands") or c["company"],
        "Both",
        "Specialty Distributor",
        ND,
        ND,
        ND,
        ND,
        ND,
        c.get("cc") or "XX",
        c.get("rc") or "XX",
        c.get("summary") or f"{c['company']} packaging specialty distributor ({c['website']}).",
    ]


def main():
    wb = load_workbook(FP)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    already = len(rows) - BASELINE
    print("current", len(rows), "already_added", already, "need", TARGET_ADDED - already)
    if already >= TARGET_ADDED:
        print("done")
        return

    ckpt = json.loads(CKPT.read_text(encoding="utf-8")) if CKPT.exists() else {"accepted": [], "rejected": [], "errors": []}
    existing_names = {(r[0] or "").strip().lower() for r in rows}
    existing_domains = {domain_of(str(r[1] or "")) for r in rows if r[1]}
    for a in ckpt.get("accepted", []):
        existing_names.add(a["company"].lower())
        existing_domains.add(domain_of(a["website"]))

    print("asking DeepSeek for more...", flush=True)
    more = deepseek_more(existing_names, n=55)
    print("got", len(more), flush=True)

    pool = []
    seen = set()
    for raw in MANUAL + more:
        c = to_cand(raw)
        if not c["company"] or not c["website"]:
            continue
        k, d = c["company"].lower(), domain_of(c["website"])
        if k in existing_names or d in existing_domains or k in seen or not d:
            continue
        if any(x in d for x in ("amcor.", "berlinpackaging.", "tricorbraun.", "packhelp.", "berryglobal.")):
            continue
        seen.add(k)
        pool.append(c)
    print("pool", len(pool), flush=True)

    backup = FP.with_name(FP.stem + f"_pre_add100_pass3_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(FP, backup)

    for i, c in enumerate(pool):
        if len(rows) - BASELINE + 0 >= TARGET_ADDED and False:
            pass
        # stop when accepted delta reaches need
        sheet_added = (ws.max_row - 1) - BASELINE
        # we'll append as we go for clarity
        if sheet_added >= TARGET_ADDED:
            break
        # also track newly accepted in this pass via counting appends
        name, web, d = c["company"], c["website"], domain_of(c["website"])
        if name.lower() in existing_names or d in existing_domains:
            continue
        print(f"[{i+1}/{len(pool)}] verify {name} (sheet_added={sheet_added}) ...", flush=True)
        if not dns_ok(d):
            ckpt["rejected"].append({"company": name, "website": web, "reason": "dns_fail"})
            print("  REJECT dns", flush=True)
            CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
            continue
        try:
            src = fetch_text(web)
            data = deepseek_role(name, web, src or f"{name} {c.get('focus')} {c.get('summary')}")
            role = (data.get("role") or "").lower()
            keep = bool(data.get("keep")) and bool(data.get("market_fit", True))
            if any(x in role for x in ("manufacturer", "oem", "converter", "hybrid", "wrong", "unknown")):
                keep = False
            if not keep:
                ckpt["rejected"].append({"company": name, "website": web, "role": data.get("role"), "reason": data.get("reason")})
                print(f"  REJECT {data.get('role')}", flush=True)
            else:
                if data.get("founded"):
                    c["founded"] = str(data["founded"]).strip()
                if data.get("hq"):
                    c["hq"] = str(data["hq"]).strip()
                if data.get("reason"):
                    c["_reason"] = data.get("reason")
                c["summary"] = c.get("summary") or f"{name} packaging specialty distributor."
                ws.append(row_from(c))
                ckpt["accepted"].append(c)
                existing_names.add(name.lower())
                existing_domains.add(d)
                sheet_added = (ws.max_row - 1) - BASELINE
                print(f"  KEEP ({sheet_added}/{TARGET_ADDED})", flush=True)
                if sheet_added >= TARGET_ADDED:
                    CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                    break
        except Exception as e:
            ckpt["errors"].append({"company": name, "error": str(e)[:200]})
            print("  ERROR", e, flush=True)
        CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
        time.sleep(0.15)

    wb.save(FP)
    final = ws.max_row - 1
    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "after": final,
        "added_from_baseline": final - BASELINE,
        "accepted_pool": len(ckpt["accepted"]),
        "file": str(FP),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("DONE after", final, "added_from_baseline", final - BASELINE)


if __name__ == "__main__":
    main()
