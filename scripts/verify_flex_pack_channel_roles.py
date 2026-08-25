"""Audit Flexible Packaging FINAL: keep packaging distributors only; flag OEMs/hybrids/wrong-market."""
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

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "channel_role_verify_388_audit.json"
CKPT = FP.parent / "channel_role_verify_388_ckpt.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()

# Hard-remove: known OEMs / converters / hybrid packaging manufacturers / wrong market
HARD_REMOVE_DOMAINS = {
    # major packaging OEMs / converters
    "amcor.com",
    "berryglobal.com",
    "sealedair.com",
    "pregis.com",
    "ranpak.com",
    "sonoco.com",
    "westrock.com",
    "smurfitkappa.com",
    "dssmith.com",
    "packagingcorp.com",
    "internationalpaper.com",
    "graphicpkg.com",
    "storopack.com",
    "signode.com",
    "fpintl.com",
    "polybags.co.uk",  # manufacturer
    # lab / wrong market previously removed patterns
    "thermofisher.com",
    "vwr.com",
    "coleparmer.com",
    "fishersci.com",
    # hybrid manufacturer+distributor / custom brand producers that slipped
    "canpaco.com",
    "dcgpac.com",
    "stephengould.com",
    "westernplastics.com",
    "packhelp.com",
    "packhelp.de",
    "packhelp.fr",
    "packhelp.co.uk",
    "noissue.co",
    "packlane.com",
    "ecoenclose.com",
    # hybrid packaging suppliers (not pure channel)
    "berlinpackaging.com",
    "berlinpackaging.ca",
    "berlinpackaging.eu",
    "tricorbraun.com",
    "tricorbraun.ca",
    "tricorbraun.com.au",
    # clear manufacturer / hybrid from sheet summaries
    "actionpkg.com",
    "packagingfirst.co.uk",
    "packaging-first.co.uk",
    "aapolybag.com",
    "southwestpolybag.com",
}

HARD_REMOVE_NAME_SUBSTR = [
    "manufacturer",
    " oem",
    "converter",
    "extrusion",
]

HARD_REMOVE_COMPANY_EXACT = {
    "action packaging systems",
    "packaging first ltd",
    "packaging first",
    "cormack packaging",  # TricorBraun AU affiliate/OEM group
    "berlin packaging",
    "berlin packaging canada",
    "tricorbraun",
    "tricorbraun canada",
    "ecoenclose",
    "noissue",
    "packhelp",
    "packlane",
    "southwest poly bag",
}

# Strong distributor/merchant name signals (prefer keep)
KEEP_HINTS = [
    "distributor",
    "distribution",
    "wholesaler",
    "wholesale",
    "merchant",
    "supplies",
    "supply",
    "catalog",
    "manutan",
    "raja",
    "rajapack",
    "ratioform",
    "lyreco",
    "viking",
    "uline",
    "bunzl",
    "grainger",
    "fastenal",
    "staples",
    "office depot",
]


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def fetch_text(website: str) -> str:
    if not website:
        return ""
    base = website if website.startswith("http") else "https://" + website
    base = base.rstrip("/")
    chunks = []
    for p in ["", "/about", "/about-us", "/company", "/our-story"]:
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
            with urlopen(req, timeout=12) as resp:
                html = resp.read(100_000).decode("utf-8", errors="ignore")
            text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
            text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text)
            chunks.append(text[:3500])
        except Exception:
            continue
        if len("\n".join(chunks)) > 6000:
            break
    return "\n".join(chunks)[:8000]


def deepseek_role(company: str, website: str, summary: str, source: str) -> dict:
    from openai import OpenAI

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    prompt = f"""Classify company role for a Flexible Packaging / packaging supplies CHANNEL landscape.

Company: {company}
Website: {website}
Existing summary: {summary}

Decide ONE role:
- Specialty Distributor / Wholesaler / Merchant / Channel Partner  -> KEEP
- Manufacturer / OEM / Converter / Hybrid Manufacturer  -> REMOVE
- Wrong market (not packaging-related) -> REMOVE
- Unknown -> REMOVE if source weak

Return JSON only:
{{
  "keep": true|false,
  "role": "Specialty Distributor"|"Wholesaler"|"Merchant"|"Channel Partner"|"Manufacturer"|"OEM"|"Converter"|"Hybrid Manufacturer"|"Wrong Market"|"Unknown",
  "market_fit": true|false,
  "evidence": "short quote or reason",
  "reason": "why keep or remove"
}}

Rules:
- KEEP only if primarily distributes/wholesales/supplies packaging products as a channel (catalog merchant, industrial packaging distributor, packaging wholesaler).
- REMOVE if primarily manufactures/converts films, pouches, corrugated, bottles as OEM.
- Hybrid manufacturer+distributor: REMOVE.
- Office/industrial catalog merchants that clearly sell packaging supplies: KEEP.
- Foodservice distributors that sell packaging disposables: KEEP if packaging is a clear product line.
- Pure food distributors with no packaging focus: REMOVE wrong market.

SOURCE TEXT:
===
{source[:7000]}
===
"""
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Classify packaging channel roles strictly. Prefer remove when hybrid/OEM."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        response_format={"type": "json_object"},
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        return json.loads(raw)
    except Exception:
        return {"keep": False, "role": "Unknown", "reason": "bad_json", "market_fit": False}


def keyword_prefilter(source: str, company: str, summary: str) -> str | None:
    """Return 'remove_mfr' | 'likely_keep' | None (needs API)."""
    blob = f"{company} {summary} {source}".lower()
    mfr = (
        "we manufacture",
        "our manufacturing",
        "manufacturer of",
        "film converter",
        "we convert",
        "extrusion",
        "we are a manufacturer",
        "oem manufacturer",
        "producing films",
        "our plants produce",
    )
    dist = (
        "distributor",
        "distributes",
        "distribution",
        "wholesaler",
        "wholesale",
        "merchant",
        "packaging supplies",
        "shipping supplies",
        "catalog",
        "catalogue",
        "we stock",
        "in stock",
        "channel partner",
    )
    has_mfr = any(k in blob for k in mfr)
    has_dist = any(k in blob for k in dist)
    if has_mfr and not has_dist:
        return "remove_mfr"
    if has_dist and not has_mfr:
        return "likely_keep"
    if has_mfr and has_dist:
        return "hybrid_suspect"
    return None


def main():
    if not API_KEY:
        raise SystemExit("OPENAI_API_KEY missing")

    backup = FP.with_name(FP.stem + f"_pre_role_verify_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(FP, backup)
    print("backup", backup)

    wb = load_workbook(FP)
    ws = wb.active
    header = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        rows.append(list(r))

    ckpt = {"checked": [], "remove": [], "keep": [], "errors": []}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass
    done = {x.get("company") for x in ckpt.get("checked", [])}

    decisions = {}  # company -> keep bool + meta
    for item in ckpt.get("keep", []):
        decisions[item["company"]] = True
    for item in ckpt.get("remove", []):
        decisions[item["company"]] = False

    for i, r in enumerate(rows):
        company = str(r[0] or "").strip()
        website = str(r[1] or "").strip()
        summary = str(r[20] or "").strip() if len(r) > 20 else ""
        d = domain_of(website)
        if not company:
            continue
        if company in done and company in decisions:
            continue

        # hard remove
        if (
            d in HARD_REMOVE_DOMAINS
            or company.lower() in HARD_REMOVE_COMPANY_EXACT
            or any(s in company.lower() for s in HARD_REMOVE_NAME_SUBSTR)
            or ("manufacturer" in summary.lower() and "distributor/manufacturer" in summary.lower())
            or re.search(r"\bindependent manufacturer\b", summary, re.I)
        ):
            meta = {
                "company": company,
                "website": website,
                "domain": d,
                "reason": "hard_remove_list",
                "role": "Manufacturer/OEM/Hybrid/Wrong",
            }
            ckpt["remove"].append(meta)
            ckpt["checked"].append({"company": company, "action": "remove"})
            decisions[company] = False
            print(f"[{i+1}/{len(rows)}] REMOVE hard {company}", flush=True)
            CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
            continue

        print(f"[{i+1}/{len(rows)}] check {company} ...", flush=True)
        try:
            source = fetch_text(website)
            # include summary in source for API
            src = (source + "\n" + summary)[:8000]
            pref = keyword_prefilter(source, company, summary)

            # Always API-check hybrid suspects and thin/no-keyword cases
            if pref == "likely_keep" and len(source) > 400:
                # still spot-check with cheap local keep for obvious merchant groups
                name_l = company.lower()
                if any(
                    h in name_l
                    for h in (
                        "raja",
                        "manutan",
                        "ratioform",
                        "lyreco",
                        "viking",
                        "uline",
                        "bunzl",
                        "staples",
                        "office depot",
                        "grainger",
                        "fastenal",
                        "verpacking",
                        "verpackung",
                        "emballage",
                        "packaging supply",
                        "packaging supplies",
                        "packaging distributor",
                    )
                ):
                    ckpt["keep"].append(
                        {
                            "company": company,
                            "website": website,
                            "reason": "merchant_group_keyword",
                            "prefilter": pref,
                        }
                    )
                    ckpt["checked"].append({"company": company, "action": "keep"})
                    decisions[company] = True
                    CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
                    time.sleep(0.05)
                    continue

            data = deepseek_role(company, website, summary, src if src.strip() else summary or company)
            keep = bool(data.get("keep")) and bool(data.get("market_fit", True))
            role = (data.get("role") or "").lower()
            if any(x in role for x in ("manufacturer", "oem", "converter", "hybrid", "wrong", "unknown")):
                keep = False
            if keep:
                ckpt["keep"].append(
                    {
                        "company": company,
                        "website": website,
                        "role": data.get("role"),
                        "evidence": data.get("evidence"),
                        "reason": data.get("reason"),
                    }
                )
                decisions[company] = True
                print(f"  KEEP {data.get('role')}", flush=True)
            else:
                ckpt["remove"].append(
                    {
                        "company": company,
                        "website": website,
                        "role": data.get("role"),
                        "evidence": data.get("evidence"),
                        "reason": data.get("reason") or data.get("reject_reason"),
                    }
                )
                decisions[company] = False
                print(f"  REMOVE {data.get('role')} | {data.get('reason')}", flush=True)
            ckpt["checked"].append({"company": company, "action": "keep" if keep else "remove"})
        except Exception as e:
            ckpt["errors"].append({"company": company, "error": str(e)[:300]})
            # fail-safe: keep if already labeled Specialty Distributor and name looks channel; else remove unknown
            name_l = company.lower()
            if any(h in name_l for h in KEEP_HINTS):
                decisions[company] = True
                ckpt["keep"].append({"company": company, "reason": f"error_failsafe_keep:{e}"[:200]})
            else:
                decisions[company] = False
                ckpt["remove"].append({"company": company, "reason": f"error_failsafe_remove:{e}"[:200]})
            ckpt["checked"].append({"company": company, "action": "error"})
            print("  ERROR", e, flush=True)
        CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
        time.sleep(0.25)

    # rebuild sheet
    kept_rows = []
    removed_rows = []
    for r in rows:
        company = str(r[0] or "").strip()
        if decisions.get(company, False):
            # force Distribution Type
            if len(r) > 12:
                r[12] = "Specialty Distributor"
            kept_rows.append(r)
        else:
            removed_rows.append({"company": company, "website": r[1] if len(r) > 1 else ""})

    ws.delete_rows(1, ws.max_row)
    ws.append(header)
    for r in kept_rows:
        ws.append(r)
    wb.save(FP)

    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "before": len(rows),
        "after": len(kept_rows),
        "removed_count": len(removed_rows),
        "kept_count": len(kept_rows),
        "errors": len(ckpt.get("errors", [])),
        "removed": ckpt.get("remove", []),
        "file": str(FP),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")
    print("DONE before", audit["before"], "after", audit["after"], "removed", audit["removed_count"])


if __name__ == "__main__":
    main()
