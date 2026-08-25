"""Add verified packaging DISTRIBUTORS only to Flexible Packaging FINAL.

Reject manufacturers/OEMs/converters/hybrids.
Fill columns only from website About text + DeepSeek extraction (no invented contacts).
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
CANDIDATES = FP.parent / "add_distributor_candidates.json"
AUDIT = FP.parent / "add_distributor_batch_audit.json"
CKPT = FP.parent / "add_distributor_batch_ckpt.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
TARGET_ADD = int(os.getenv("FLEX_ADD_TARGET") or "200")
LIMIT = int(os.getenv("FLEX_ADD_LIMIT") or "400")
ND = "Not publicly disclosed"

HEADERS = [
    "Company",
    "Website",
    "Founded",
    "Headquarters",
    "Continent / Geography",
    "Operational Presence",
    "Ownership",
    "Employees",
    "Core Categories",
    "Specialty Focus",
    "Key Brands Represented",
    "Retail / E-commerce / Both",
    "Distribution Type",
    "Contact Person",
    "Role",
    "Email",
    "LinkedIn",
    "Office No.",
    "Country Code",
    "Region Code",
    "Summary",
]


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def norm_name(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def fetch_text(website: str, max_pages: int = 5) -> str:
    if not website:
        return ""
    base = website if website.startswith("http") else "https://" + website
    base = base.rstrip("/")
    paths = ["", "/about", "/about-us", "/about-us/", "/company", "/our-story", "/history", "/who-we-are"]
    chunks = []
    for p in paths[:max_pages]:
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
            with urlopen(req, timeout=14) as resp:
                html = resp.read(140_000).decode("utf-8", errors="ignore")
            text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
            text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text)
            chunks.append(text[:4500])
        except Exception:
            continue
    return "\n".join(chunks)[:12000]


def deepseek_verify_and_fill(company: str, website: str, country_hint: str, source: str) -> dict:
    from openai import OpenAI

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    prompt = f"""You verify whether a company is a packaging DISTRIBUTOR / wholesaler / merchant / channel partner
(NOT a primary manufacturer/OEM/converter), and extract landscape fields ONLY from SOURCE TEXT.

Company: {company}
Website: {website}
Country hint: {country_hint}

Return JSON ONLY:
{{
  "accept": true|false,
  "role": "Specialty Distributor"|"Wholesaler"|"Merchant"|"Manufacturer"|"OEM"|"Converter"|"Hybrid Manufacturer"|"Unknown"|"Marketplace",
  "role_evidence": "short quote from SOURCE",
  "founded": "YYYY"|null,
  "headquarters": "city/region/country or address"|null,
  "continent_geography": "e.g. NA(United States)"|null,
  "operational_presence": "countries/regions if stated"|null,
  "ownership": "Private|Public|Subsidiary ..."|null,
  "employees": "band or number if stated"|null,
  "core_categories": "semicolon-separated product categories"|null,
  "specialty_focus": "one sentence"|null,
  "key_brands": "brands they distribute if stated; else company name"|null,
  "retail_ecom": "Retail"|"E-commerce"|"Both"|null,
  "linkedin": "company LinkedIn URL if in SOURCE"|null,
  "country_code": "ISO2"|null,
  "region_code": "NA|EU|APAC|LATAM|MEA"|null,
  "summary": "1-2 factual sentences"|null,
  "reject_reason": "if accept=false"
}}

HARD RULES:
- accept=true ONLY if SOURCE clearly shows they distribute/wholesale/supply packaging products as a merchant/distributor (distributor, wholesaler, merchant, distribution centre, catalogue sales, packaging supplies stockist).
- If primarily manufacturer/converter/OEM of films/pouches/boxes, accept=false.
- Hybrid manufacturer+distributor: accept=false (reject_reason=hybrid).
- founded/hq/employees: ONLY if explicitly in SOURCE. Else null. NEVER invent.
- Never invent emails/phones/contact persons.
- country_code/region_code from HQ/country evidence.

SOURCE TEXT:
===
{source[:10000]}
===
"""
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Extract only evidenced facts. Never invent contacts or years."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        response_format={"type": "json_object"},
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        return json.loads(raw)
    except Exception:
        return {"accept": False, "reject_reason": "bad_json", "raw": raw[:500]}


def year_ok(y, source: str):
    if not y:
        return None
    ys = str(y).strip()
    if not re.fullmatch(r"(18|19|20)\d{2}", ys):
        return None
    if ys not in source:
        return None
    return ys


def load_existing(ws):
    names, domains = set(), set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        names.add(norm_name(str(r[0] or "")))
        domains.add(domain_of(str(r[1] or "")))
    return names, domains


def main():
    if not API_KEY:
        raise SystemExit("OPENAI_API_KEY missing")
    if not CANDIDATES.exists():
        raise SystemExit(f"Missing candidates: {CANDIDATES}")

    cands = json.loads(CANDIDATES.read_text(encoding="utf-8"))
    ckpt = {"added": [], "rejected": [], "errors": [], "processed_ids": [], "pass": "resume"}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    backups = list(FP.parent.glob(FP.stem + "_pre_add_dist_*.xlsx"))
    if not backups:
        backup = FP.with_name(FP.stem + f"_pre_add_dist_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(FP, backup)
        print("backup", backup)

    wb = load_workbook(FP)
    ws = wb.active
    names, domains = load_existing(ws)
    for a in ckpt.get("added", []):
        names.add(norm_name(a.get("Company", "")))
        domains.add(domain_of(a.get("Website", "")))

    added_count = len(ckpt.get("added", []))
    processed = set(ckpt.get("processed_ids", []))

    for i, c in enumerate(cands[:LIMIT]):
        if added_count >= TARGET_ADD:
            break
        cid = c.get("id") or f"{c.get('company')}|{c.get('website')}"
        if cid in processed:
            continue
        company = (c.get("company") or "").strip()
        website = (c.get("website") or "").strip()
        country = (c.get("country") or "").strip()
        if not company or not website:
            processed.add(cid)
            continue
        if norm_name(company) in names or domain_of(website) in domains:
            ckpt.setdefault("rejected", []).append({"company": company, "reason": "duplicate"})
            processed.add(cid)
            continue

        print(f"[{i+1}/{min(LIMIT,len(cands))}] {company} ...", flush=True)
        try:
            source = fetch_text(website)
            if len(source) < 200:
                ckpt.setdefault("rejected", []).append(
                    {"company": company, "reason": "thin_or_unreachable_site", "website": website}
                )
                processed.add(cid)
                ckpt["processed_ids"] = list(processed)
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                continue
            src_l = source.lower()
            channel_kw = (
                "distributor",
                "distribute",
                "distribution",
                "wholesaler",
                "wholesale",
                "merchant",
                "catalogue",
                "catalog",
                "packaging supplies",
                "shipping supplies",
                "stockist",
                "reseller",
                "channel partner",
                "we supply",
                "we stock",
                "in stock",
            )
            mfr_kw = (
                "we manufacture",
                "our manufacturing",
                "film converter",
                "extrusion",
                "we are a manufacturer",
                "oem manufacturer",
            )
            if not any(k in src_l for k in channel_kw):
                ckpt.setdefault("rejected", []).append(
                    {"company": company, "reason": "no_channel_keywords", "website": website}
                )
                processed.add(cid)
                ckpt["processed_ids"] = list(processed)
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                continue
            if any(k in src_l for k in mfr_kw) and not any(
                k in src_l for k in ("distributor", "wholesaler", "wholesale distributor", "packaging distributor")
            ):
                ckpt.setdefault("rejected", []).append(
                    {"company": company, "reason": "manufacturer_keywords", "website": website}
                )
                processed.add(cid)
                ckpt["processed_ids"] = list(processed)
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                continue

            data = deepseek_verify_and_fill(company, website, country, source)
            role = (data.get("role") or "").strip()
            accept = bool(data.get("accept"))
            role_l = role.lower()
            if (not accept) or any(
                x in role_l for x in ("manufacturer", "oem", "converter", "hybrid", "marketplace", "unknown")
            ):
                ckpt.setdefault("rejected", []).append(
                    {
                        "company": company,
                        "website": website,
                        "role": role,
                        "reason": data.get("reject_reason") or "role_reject",
                        "evidence": data.get("role_evidence"),
                    }
                )
                processed.add(cid)
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                time.sleep(0.35)
                continue

            founded = year_ok(data.get("founded"), source)
            row = {
                "Company": company,
                "Website": website if website.startswith("http") else "https://" + website,
                "Founded": founded or ND,
                "Headquarters": data.get("headquarters") or ND,
                "Continent / Geography": data.get("continent_geography") or ND,
                "Operational Presence": data.get("operational_presence") or ND,
                "Ownership": data.get("ownership") or ND,
                "Employees": data.get("employees") or ND,
                "Core Categories": data.get("core_categories") or "Packaging supplies distribution",
                "Specialty Focus": data.get("specialty_focus") or "Distributes packaging materials and supplies",
                "Key Brands Represented": data.get("key_brands") or company,
                "Retail / E-commerce / Both": data.get("retail_ecom") or "Both",
                "Distribution Type": "Specialty Distributor",
                "Contact Person": ND,
                "Role": ND,
                "Email": ND,
                "LinkedIn": data.get("linkedin") or ND,
                "Office No.": ND,
                "Country Code": (data.get("country_code") or country or "XX")[:2].upper(),
                "Region Code": data.get("region_code") or "XX",
                "Summary": data.get("summary")
                or f"{company} packaging distributor/merchant ({domain_of(website)}).",
            }
            if row["Headquarters"] == ND and founded is None and not data.get("role_evidence"):
                ckpt.setdefault("rejected", []).append({"company": company, "reason": "insufficient_fields"})
                processed.add(cid)
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                continue

            ws.append([row[h] for h in HEADERS])
            names.add(norm_name(company))
            domains.add(domain_of(website))
            ckpt.setdefault("added", []).append({**row, "role_evidence": data.get("role_evidence"), "api_role": role})
            added_count += 1
            processed.add(cid)
            print(f"  ADDED ({added_count}/{TARGET_ADD}) role={role} founded={row['Founded']}", flush=True)
            if added_count % 5 == 0:
                wb.save(FP)
                CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
        except Exception as e:
            ckpt.setdefault("errors", []).append({"company": company, "error": str(e)[:300]})
            processed.add(cid)
            print("  ERROR", e, flush=True)
        ckpt["processed_ids"] = list(processed)
        CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
        time.sleep(0.3)

    for col in ws.columns:
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col[0].column_letter].width = 22
    for cell in ws[1]:
        cell.font = Font(bold=True)

    wb.save(FP)
    audit = {
        "ts": datetime.now().isoformat(),
        "target": TARGET_ADD,
        "added": len(ckpt.get("added", [])),
        "rejected": len(ckpt.get("rejected", [])),
        "errors": len(ckpt.get("errors", [])),
        "final_rows": ws.max_row - 1,
        "file": str(FP),
        "detail": ckpt,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
    print("DONE added", audit["added"], "final_rows", audit["final_rows"], "rejected", audit["rejected"])


if __name__ == "__main__":
    main()
