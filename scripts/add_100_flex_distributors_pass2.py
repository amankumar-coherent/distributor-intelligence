"""Continue adding verified distributors from extra candidates until +100 total from baseline."""
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
from openpyxl import load_workbook
from openpyxl.styles import Alignment

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
EXTRA = FP.parent / "add100_extra_candidates.json"
CKPT = FP.parent / "add100_distributors_ckpt.json"
AUDIT = FP.parent / "add100_distributors_audit_pass2.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
ND = "Not publicly disclosed"
BASELINE = 224
TARGET_TOTAL_ADDED = 100


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
            req = Request(
                base + p,
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                    )
                },
            )
            with urlopen(req, timeout=12, context=ctx) as resp:
                html = resp.read(70_000).decode("utf-8", errors="ignore")
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


def deepseek_role(company: str, website: str, source: str) -> dict:
    from openai import OpenAI

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    prompt = f"""Classify for Flexible Packaging CHANNEL landscape.
Company: {company}
Website: {website}
SOURCE:
{source[:6000]}

KEEP only Specialty Distributor / Wholesaler / Merchant / Channel Partner of packaging.
REMOVE Manufacturer / OEM / Converter / Hybrid / Wrong Market / Unknown.

JSON: {{"keep":true|false,"role":"...","market_fit":true|false,"founded":"YYYY or Not publicly disclosed","hq":"...","reason":"short"}}"""
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Strict channel-only. Prefer remove hybrids/OEMs."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        response_format={"type": "json_object"},
    )
    try:
        return json.loads(resp.choices[0].message.content or "{}")
    except Exception:
        return {"keep": False, "role": "Unknown", "market_fit": False, "reason": "bad_json"}


def to_cand(c: dict) -> dict:
    cc = (c.get("country_code") or "XX").strip().upper()[:2]
    if cc in {"US", "CA", "MX"}:
        rc, geo = "NA", f"NA({cc})"
    elif cc in {"AU", "NZ", "SG", "IN", "JP", "CN", "KR", "MY", "TH", "ID"}:
        rc, geo = "APAC", f"APAC({cc})"
    else:
        rc, geo = "EU", f"Europe({cc})"
    web = (c.get("website") or "").strip()
    if web and not web.startswith("http"):
        web = "https://" + web
    return {
        "company": (c.get("company") or "").strip(),
        "website": web,
        "hq": (c.get("hq") or ND).strip(),
        "geo": geo,
        "cc": cc,
        "rc": rc,
        "founded": (c.get("founded") or ND).strip(),
        "cats": (c.get("core_categories") or "Packaging supplies").strip(),
        "focus": (c.get("specialty_focus") or "Packaging distributor").strip(),
        "brands": (c.get("company") or "Multi-brand").strip(),
        "summary": (c.get("summary") or "").strip(),
        "_source": "deepseek_batch",
    }


def row_from(c: dict, role: str) -> list:
    summary = c.get("summary") or (
        f"{c['company']} is a packaging {role.lower()} ({c['website']})."
    )
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
        summary,
    ]


def main():
    if not API_KEY:
        raise SystemExit("missing API key")

    wb = load_workbook(FP)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    already_added = max(0, len(rows) - BASELINE)
    # if first pass hasn't written yet, use ckpt accepted
    ckpt = {"accepted": [], "rejected": [], "errors": []}
    if CKPT.exists():
        ckpt = json.loads(CKPT.read_text(encoding="utf-8"))

    need = TARGET_TOTAL_ADDED - max(already_added, len(ckpt.get("accepted", [])))
    # Prefer counting accepted in ckpt if sheet not yet updated
    current_accepted = len(ckpt.get("accepted", []))
    if already_added >= TARGET_TOTAL_ADDED:
        print("already have", already_added, "added in sheet")
        return
    if current_accepted >= TARGET_TOTAL_ADDED and already_added < TARGET_TOTAL_ADDED:
        # just write remaining accepted to sheet
        need_write = TARGET_TOTAL_ADDED - already_added
    else:
        need_write = None

    print("sheet", len(rows), "baseline", BASELINE, "already_added", already_added, "ckpt_accepted", current_accepted)

    existing_names = {(r[0] or "").strip().lower() for r in rows}
    existing_domains = {domain_of(str(r[1] or "")) for r in rows if r[1]}
    for a in ckpt.get("accepted", []):
        existing_names.add(a["company"].lower())
        existing_domains.add(domain_of(a["website"]))

    extras = json.loads(EXTRA.read_text(encoding="utf-8")) if EXTRA.exists() else []
    # also add a few manually researched
    extras.extend(
        [
            {
                "company": "Gillis & Lane",
                "website": "https://gillislane.com",
                "hq": "USA",
                "country_code": "US",
                "founded": "Not publicly disclosed",
                "core_categories": "Packaging and paper products",
                "specialty_focus": "US packaging/paper distributor",
                "summary": "Gillis & Lane distributes packaging and paper products with CA/TX warehouses.",
            },
            {
                "company": "Pro Pac Materials",
                "website": "https://www.propacmaterials.com",
                "hq": "Grand Prairie, TX, USA",
                "country_code": "US",
                "founded": "1971",
                "core_categories": "Packaging materials and machinery",
                "specialty_focus": "Southwest US packaging materials distributor",
                "summary": "Pro Pac Materials is an independent packaging materials and machinery distributor in the Southwest US.",
            },
            {
                "company": "Furodet Emballages",
                "website": "https://furodet-emballages.fr",
                "hq": "France",
                "country_code": "FR",
                "founded": "Not publicly disclosed",
                "core_categories": "Food and retail packaging supplies",
                "specialty_focus": "French packaging merchant/wholesaler",
                "summary": "Furodet Emballages supplies packaging materials to businesses in France/Belgium.",
            },
            {
                "company": "Embageant",
                "website": "https://embageant.fr",
                "hq": "France / Netherlands",
                "country_code": "FR",
                "founded": "Not publicly disclosed",
                "core_categories": "E-commerce packaging materials",
                "specialty_focus": "Online packaging wholesaler",
                "summary": "Embageant is an online packaging materials wholesaler serving France, NL, Belgium.",
            },
            {
                "company": "NNZ",
                "website": "https://www.nnz.com",
                "hq": "Netherlands / Europe",
                "country_code": "NL",
                "founded": "1922",
                "core_categories": "Produce and industrial packaging",
                "specialty_focus": "International packaging distributor network",
                "summary": "NNZ is a multinational packaging specialist/distributor network founded in 1922.",
            },
            {
                "company": "Weller Packaging",
                "website": "https://www.wellerpackaging.co.uk",
                "hq": "UK",
                "country_code": "GB",
                "founded": "Not publicly disclosed",
                "core_categories": "Flexible polythene and paper packaging wholesale",
                "specialty_focus": "UK wholesaler to packaging merchants",
                "summary": "Weller Packaging is a UK importer/stockist/distributor supplying the merchant trade.",
            },
            {
                "company": "Brangs Heinrich",
                "website": "https://www.brangs-heinrich.de",
                "hq": "Germany",
                "country_code": "DE",
                "founded": "Not publicly disclosed",
                "core_categories": "Industrial packaging; tapes; films",
                "specialty_focus": "German packaging wholesaler",
                "summary": "Brangs + Heinrich is a German packaging wholesaler/distributor.",
            },
            {
                "company": "Sankawa Packaging Singapore",
                "website": "https://www.sankawa.com.sg",
                "hq": "Singapore",
                "country_code": "SG",
                "founded": "1989",
                "core_categories": "Bubble; mailers; cartons; tapes",
                "specialty_focus": "Singapore packaging distributor",
                "summary": "Sankawa Packaging is a Singapore packaging specialist distributor.",
            },
        ]
    )

    cands = []
    seen = set()
    for raw in extras:
        c = to_cand(raw)
        if not c["company"] or not c["website"]:
            continue
        k = c["company"].lower()
        d = domain_of(c["website"])
        if k in existing_names or d in existing_domains or k in seen:
            continue
        if any(x in d for x in ("amcor.", "berlinpackaging.", "tricorbraun.", "packhelp.", "berryglobal.", "sealedair.")):
            continue
        seen.add(k)
        cands.append(c)

    print("new candidates to verify", len(cands), "need_more_accepts", max(0, TARGET_TOTAL_ADDED - current_accepted))

    backup = FP.with_name(FP.stem + f"_pre_add100_pass2_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(FP, backup)

    for i, c in enumerate(cands):
        if len(ckpt["accepted"]) >= TARGET_TOTAL_ADDED:
            break
        name, web, d = c["company"], c["website"], domain_of(c["website"])
        print(f"[{i+1}/{len(cands)}] verify {name} ...", flush=True)
        if not dns_ok(d):
            # try www
            if not d.startswith("www.") and dns_ok("www." + d):
                if "://" in web:
                    scheme = web.split("://", 1)[0]
                    path = "/" + "/".join(web.split("://", 1)[1].split("/")[1:]) if "/" in web.split("://", 1)[1] else ""
                    web = f"{scheme}://www.{d}{path if path != '/' else ''}"
                    c["website"] = web
                    d = "www." + d if False else d
                else:
                    pass
            else:
                ckpt["rejected"].append({"company": name, "website": web, "reason": "dns_fail"})
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                print("  REJECT dns", flush=True)
                continue
        try:
            src = fetch_text(web)
            data = deepseek_role(name, web, src or f"{name} {c.get('focus')} {c.get('summary')}")
            role = (data.get("role") or "").lower()
            keep = bool(data.get("keep")) and bool(data.get("market_fit", True))
            if any(x in role for x in ("manufacturer", "oem", "converter", "hybrid", "wrong", "unknown")):
                keep = False
            if not keep:
                ckpt["rejected"].append(
                    {"company": name, "website": web, "role": data.get("role"), "reason": data.get("reason")}
                )
                print(f"  REJECT {data.get('role')}", flush=True)
            else:
                if data.get("founded"):
                    c["founded"] = str(data["founded"]).strip()
                if data.get("hq"):
                    c["hq"] = str(data["hq"]).strip()
                c["_role"] = data.get("role") or "Specialty Distributor"
                c["_reason"] = data.get("reason")
                ckpt["accepted"].append(c)
                existing_names.add(name.lower())
                existing_domains.add(domain_of(web))
                print(f"  KEEP {data.get('role')} ({len(ckpt['accepted'])}/{TARGET_TOTAL_ADDED})", flush=True)
        except Exception as e:
            ckpt["errors"].append({"company": name, "error": str(e)[:240]})
            print("  ERROR", e, flush=True)
        CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
        time.sleep(0.2)

    # Rewrite sheet: keep original baseline rows + first 100 accepted (avoid double-append)
    # Reload original from backup of first pass if needed — safer: rebuild from current unique names
    # Strategy: if current sheet still ~224, append missing accepted. If already appended some, only append delta.

    current = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    cur_names = {(r[0] or "").strip().lower() for r in current}
    added_now = 0
    for c in ckpt["accepted"]:
        if added_now + max(0, len(current) - BASELINE) >= TARGET_TOTAL_ADDED and len(current) >= BASELINE:
            # stop when total added from baseline hits 100
            if len(current) + added_now - BASELINE >= TARGET_TOTAL_ADDED:
                break
        if c["company"].lower() in cur_names:
            continue
        if (len(current) - BASELINE) + added_now >= TARGET_TOTAL_ADDED:
            break
        ws.append(row_from(c, c.get("_role") or "Specialty Distributor"))
        cur_names.add(c["company"].lower())
        added_now += 1

    for row in ws.iter_rows(min_row=max(2, ws.max_row - added_now), max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(FP)
    final = ws.max_row - 1
    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "baseline": BASELINE,
        "before_pass2": len(current),
        "added_this_pass": added_now,
        "after": final,
        "accepted_pool": len(ckpt["accepted"]),
        "rejected_pool": len(ckpt["rejected"]),
        "file": str(FP),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("DONE after", final, "added_this_pass", added_now, "accepted_pool", len(ckpt["accepted"]))


if __name__ == "__main__":
    main()
