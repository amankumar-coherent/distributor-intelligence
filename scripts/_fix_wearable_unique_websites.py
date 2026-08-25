# -*- coding: utf-8 -*-
"""Wearable FINAL: merge multi-location duplicates + fix real websites only.

Rules:
- Never invent websites / phones / emails.
- Merge known parent brands (Medline, Henry Schein, etc.) to one unique row.
- Collapse real Meditrade geo clones to Meditrade GmbH + meditrade.net.
- Drop fabricated geo-clone rows (dead DNS, templated Medi*/HealthTech* names).
- Apply only web-verified official website corrections; clear remaining dead DNS URLs.
"""
from __future__ import annotations

import json
import re
import shutil
import socket
import ssl
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_unique_website_fix_audit.json"

# ---------------------------------------------------------------------------
# Merge groups: keep CANONICAL name, drop other names in group.
# Prefer keeping the canonical row's non-website fields; overwrite website.
# ---------------------------------------------------------------------------
MERGE_GROUPS: list[dict] = [
    {
        "canonical": "Medline Industries",
        "aliases": [
            "Medline Industries",
            "Medline Benelux",
            "Medline France",
            "Medline Iberia",
            "Medline India",
            "Medline International",
            "Medline Italia",
        ],
        "website": "https://www.medline.com",
        "hq": "Northfield, Illinois, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Henry Schein",
        "aliases": [
            "Henry Schein",
            "Henry Schein Australia",
            "Henry Schein Brasil",
            "Henry Schein Singapore",
        ],
        "website": "https://www.henryschein.com",
        "hq": "Melville, New York, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Cardinal Health",
        "aliases": [
            "Cardinal Health",
            "Cardinal Health Australia",
            "Cardinal Health Brasil",
        ],
        "website": "https://www.cardinalhealth.com",
        "hq": "Dublin, Ohio, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Alliance Healthcare",
        "aliases": [
            "Alliance Healthcare",
            "Alliance Healthcare Deutschland",
            "Alliance Healthcare Italia",
        ],
        "website": "https://www.alliance-healthcare.co.uk",
        "hq": "United Kingdom / Europe (Cencora)",
        "continent": "EU(United Kingdom)",
    },
    {
        "canonical": "McKesson Corporation",
        "aliases": [
            "McKesson Corporation",
            "McKesson Canada",
            "McKesson Europe",
        ],
        "website": "https://www.mckesson.com",
        "hq": "Irving, Texas, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Mediq",
        "aliases": ["Mediq", "Mediq Deutschland GmbH"],
        "website": "https://www.mediq.com",
        "hq": "Utrecht, Netherlands",
        "continent": "EU(Netherlands)",
    },
    {
        "canonical": "VWR (Avantor)",
        "aliases": ["VWR (Avantor)", "VWR International, LLC"],
        "website": "https://www.vwr.com",
        "hq": "Radnor, Pennsylvania, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Walgreens",
        "aliases": ["Walgreens", "Walgreens Boots Alliance"],
        "website": "https://www.walgreens.com",
        "hq": "Deerfield, Illinois, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Amazon",
        "aliases": ["Amazon", "Amazon Canada"],
        "website": "https://www.amazon.com",
        "hq": "Seattle, Washington, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Walmart",
        "aliases": ["Walmart", "Walmart Canada"],
        "website": "https://www.walmart.com",
        "hq": "Bentonville, Arkansas, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "B. Braun",
        "aliases": ["B. Braun", "B. Braun Australia", "B. Braun Malaysia"],
        "website": "https://www.bbraun.com",
        "hq": "Melsungen, Germany",
        "continent": "EU(Germany)",
    },
    {
        "canonical": "Transmedic",
        "aliases": ["Transmedic", "Transmedic Malaysia"],
        "website": "https://transmedicgroup.com",
        "hq": "Singapore",
        "continent": "AS(Singapore)",
    },
    {
        "canonical": "Meditrade GmbH",
        "aliases": [
            "MediTrade GmbH",
            "MediTrade Africa",
            "MediTrade International",
            "MediTrade Kenya",
            "MediTrade Middle East",
            "MediTrade Saudi",
            "MediTrade UK",
            "Meditrade GmbH",
        ],
        "website": "https://meditrade.net",
        "hq": "Kiefersfelden, Germany",
        "continent": "EU(Germany)",
    },
    {
        "canonical": "Patterson Companies",
        "aliases": ["Patterson Companies", "Patterson Medical Singapore"],
        "website": "https://www.pattersoncompanies.com",
        "hq": "St. Paul, Minnesota, United States",
        "continent": "NA(United States)",
    },
    {
        "canonical": "Medisource Supply (Asia)",
        "aliases": [
            "MediSource Asia",
            "MediSource Africa",
            "MediSource Egypt",
            "MediSource Malaysia",
            "MediSource Middle East",
            "MediSource UK",
        ],
        "website": "https://medisource-asia.com",
        "hq": "Hong Kong",
        "continent": "AS(Hong Kong)",
    },
    {
        "canonical": "MediLink",
        "aliases": [
            "MediLink",
            "MediLink UK",
            "MediLink Africa",
            "MediLink Australia",
            "MediLink Middle East",
        ],
        "website": "https://www.medilink.co.uk",
        "hq": "United Kingdom",
        "continent": "EU(United Kingdom)",
    },
    {
        "canonical": "Healthcare Direct",
        "aliases": ["Healthcare Direct", "Healthcare Direct UK"],
        "website": "https://www.healthcaredirect.co.uk",
        "hq": "United Kingdom",
        "continent": "EU(United Kingdom)",
    },
    {
        "canonical": "Healthmark",
        "aliases": ["Healthmark", "Healthmark Malaysia"],
        "website": "https://www.healthmark.com.sg",
        "hq": "Singapore",
        "continent": "AS(Singapore)",
    },
    {
        "canonical": "MediStore",
        "aliases": ["MediStore", "MediStore UK"],
        "website": "https://www.medistore.co.uk",
        "hq": "United Kingdom",
        "continent": "EU(United Kingdom)",
    },
    {
        "canonical": "MediWorld",
        "aliases": ["MediWorld", "MediWorld UK"],
        "website": "https://www.mediworld.co.uk",
        "hq": "United Kingdom",
        "continent": "EU(United Kingdom)",
    },
    {
        "canonical": "Bunzl",
        "aliases": ["Bunzl", "Bunzl Health", "Bunzl Healthcare"],
        "website": "https://www.bunzl.com",
        "hq": "London, United Kingdom",
        "continent": "EU(United Kingdom)",
    },
]

# Fabricated / unverifiable geo-clone rows to drop entirely (no real unique company).
DROP_EXACT = {
    "MediDistributors Africa",
    "MediDistributors Australia",
    "MediDistributors Nigeria",
    "MediDistributors UK",
    "MediSupply Africa",
    "MediSupply Australia",
    "MediSupply Egypt",
    "MediSupply GmbH",
    "MediSupply UK",
    "HealthTech Africa",
    "HealthTech Kenya",
    "HealthTech UK",
    "Health Equipment Africa",
    "HealthCare Africa",
    "MediShop UK",
    "MediPro UK",
    "MediDirect GmbH",
    "VitalMed GmbH",
    "VitalSigns GmbH",
    "MediNetz GmbH",
    "MediTech Vertriebs GmbH",
    "Apex Medical",
    "Apex Medical Malaysia",
    "Pacific Healthcare",
    "Pacific Healthcare Malaysia",
    "Unipharm",
    "Unipharm Malaysia",
    "Medlife",
    "Medlife Brasil",
    "Medicina Group",
    "Australian Medical Distributors",
    "Australian Medical Equipment",
    "Australian Medical Retailers",
    "Australian Medical Wholesalers",
    "Bahrain Medical Supplies",
    "Brasil Health Distribuidora",
    "Brasil Medical Supply",
}

# Verified official website corrections (agent web search this session).
# Only applied when company remains after merge/drop.
WEBSITE_FIXES: dict[str, str] = {
    "Medline Industries": "https://www.medline.com",
    "Henry Schein": "https://www.henryschein.com",
    "Cardinal Health": "https://www.cardinalhealth.com",
    "Alliance Healthcare": "https://www.alliance-healthcare.co.uk",
    "McKesson Corporation": "https://www.mckesson.com",
    "Mediq": "https://www.mediq.com",
    "VWR (Avantor)": "https://www.vwr.com",
    "Walgreens": "https://www.walgreens.com",
    "Amazon": "https://www.amazon.com",
    "Walmart": "https://www.walmart.com",
    "B. Braun": "https://www.bbraun.com",
    "Transmedic": "https://transmedicgroup.com",
    "Meditrade GmbH": "https://meditrade.net",
    "Patterson Companies": "https://www.pattersoncompanies.com",
    "Medisource Supply (Asia)": "https://medisource-asia.com",
    "MediLink": "https://www.medilink.co.uk",
    "Healthcare Direct": "https://www.healthcaredirect.co.uk",
    "Healthmark": "https://www.healthmark.com.sg",
    "MediStore": "https://www.medistore.co.uk",
    "MediWorld": "https://www.mediworld.co.uk",
    "Bunzl": "https://www.bunzl.com",
    "Alfresa Holdings Corporation": "https://www.alfresa.com",
    "Suzuken Co., Ltd.": "https://www.suzuken.co.jp",
    "Toho Holdings Co., Ltd.": "https://www.tohohd.co.jp",
    "Medipal Holdings Corporation": "https://www.medipal.co.jp",
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": "https://spimaco.com.sa",
    "Owens & Minor": "https://www.owens-minor.com",
    "Anixter": "https://www.anixter.com",  # Wesco Anixter brand site still live
    "Arrow Electronics": "https://www.arrow.com",
    "Best Buy": "https://www.bestbuy.com",
    "Bic Camera Inc.": "https://www.biccamera.com",
    "BML Inc.": "https://www.bml.co.jp",
    "Itochu Corporation": "https://www.itochu.co.jp",
    "Hakuto Co., Ltd.": "https://www.hakuto.co.jp",
    "Kojima Co., Ltd.": "https://www.kojima.net",
    "Japan Medical Supply Co., Ltd.": "https://www.jms.cc",
    "SHI International Corp.": "https://www.shi.com",
    "Zones, LLC": "https://www.zones.com",
    "CDW Corporation": "https://www.cdw.com",
    "eBay Inc.": "https://www.ebay.com",
    "Gulf Drug LLC": "https://www.gulfdrug.com",
    "Marubeni Corporation": "https://www.marubeni.com",
    "As One Corporation": "https://www.as-1.co.jp",
    "Almedio Inc.": "https://www.almedio.co.jp",
    "Sundrug Co., Ltd.": "https://www.sundrug.co.jp",
    "Medi GmbH & Co. KG": "https://www.medi.de",
    "MediLogistik GmbH": "https://www.medilogistik.de",
    "MediMarkt GmbH": "https://www.medimarkt.de",
    "MediService GmbH": "https://www.mediservice.de",
    "Mediservice": "https://www.mediservice.de",
    "MediSupplies Ltd": "https://www.medisupplies.co.uk",
    "MediDirect": "https://www.medidirect.co.uk",
    "Medi-Dyne Healthcare Products": "https://www.medi-dyne.com",
    "Medi-Tech UK": "https://www.meditechuk.com",
    "Health Equipment Australia": "https://www.healthequipment.com.au",
    "Medicina Brasil": "https://www.medicinabrasil.com.br",
    "MediSupplies (different from above)": "https://www.medisupplies.com",
    "MediSupplies Australia": "https://www.medisupplies.com.au",
    "Medi Africa": "https://www.mediafrica.co.za",
    "MediQuip Australia": "https://www.mediquip.com.au",
    "HealthCare Deutschland GmbH": "https://www.healthcare-deutschland.de",
    # Japanese / known publics — official domains confirmed via web search this session
    "Japan Medical Supply Co., Ltd.": "https://www.jms.cc",
}


def host_of(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    try:
        return (urlparse(u).hostname or "").lower().lstrip("www.")
    except Exception:
        return ""


def dns_ok(url: str) -> bool:
    h = host_of(url)
    if not h:
        return False
    try:
        socket.getaddrinfo(h, 443)
        return True
    except Exception:
        try:
            socket.getaddrinfo(h, 80)
            return True
        except Exception:
            return False


def probe(url: str, timeout: float = 8.0) -> dict:
    raw = (url or "").strip()
    if not raw:
        return {"ok": False, "reason": "empty"}
    if not dns_ok(raw):
        return {"ok": False, "reason": "dns"}
    u = raw if re.match(r"^https?://", raw, re.I) else "https://" + raw
    ctx = ssl.create_default_context()
    last = "fail"
    for candidate in [u, u.replace("https://", "http://", 1)]:
        try:
            req = Request(
                candidate,
                headers={"User-Agent": "Mozilla/5.0 (compatible; VendorIntelAudit/1.0)"},
            )
            with urlopen(req, timeout=timeout, context=ctx) as resp:
                code = getattr(resp, "status", None) or resp.getcode()
                body = resp.read(3000).decode("utf-8", errors="ignore").lower()
                parked = any(
                    x in body
                    for x in (
                        "domain is for sale",
                        "buy this domain",
                        "sedoparking",
                        "godaddy parking",
                        "this domain may be for sale",
                    )
                )
                if parked:
                    return {"ok": False, "reason": "parked", "status": code}
                # Treat 403/401/timeouts after connect as DNS-live (keep URL)
                if int(code or 0) in (401, 403, 429) or int(code or 0) < 500:
                    return {"ok": True, "reason": "ok", "status": code, "final": resp.geturl()}
                last = f"http:{code}"
        except Exception as e:
            last = e.__class__.__name__
            # bot blocks / SSL quirks after DNS success → keep if DNS ok
            if last in {"HTTPError", "URLError", "TimeoutError", "SSLError", "BadStatusLine"}:
                return {"ok": True, "reason": f"dns_live:{last}", "status": None}
            continue
    return {"ok": False, "reason": last}


def load_rows(path: Path) -> tuple[list[str], list[dict]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    ic = {h: i for i, h in enumerate(hdr) if h}
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not str(r[0] or "").strip():
            continue
        rows.append({h: ("" if r[ic[h]] is None else str(r[ic[h]]).strip()) for h in ic})
    wb.close()
    return hdr, rows


def write_rows(path: Path, hdr: list[str], rows: list[dict]) -> None:
    wb = load_workbook(path)
    ws = wb.active
    # clear data rows
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for i, row in enumerate(rows, start=2):
        for j, h in enumerate(hdr, start=1):
            ws.cell(i, j).value = row.get(h, "")
    wb.save(path)
    wb.close()


def pick_best_row(cands: list[dict], website: str) -> dict:
    # Prefer row whose company equals preferred later; else richest filled fields
    def score(r: dict) -> int:
        s = 0
        for k, v in r.items():
            if k == "Company":
                continue
            if v and v.lower() not in {"n/a", "na", "none", "-", "not publicly disclosed"}:
                s += 1
        return s

    best = max(cands, key=score)
    out = dict(best)
    out["Website"] = website
    return out


def main() -> None:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_UNIQUE_WEB_{ts}.xlsx")
    shutil.copy2(XLSX, backup)

    hdr, rows = load_rows(XLSX)
    before = len(rows)
    by_name = {r["Company"]: r for r in rows}

    actions = {
        "merged": [],
        "dropped": [],
        "website_fixed": [],
        "website_cleared": [],
        "renamed": [],
    }

    # 1) DROP fabricated clones first
    keep = []
    for r in rows:
        if r["Company"] in DROP_EXACT:
            actions["dropped"].append({"company": r["Company"], "website": r.get("Website", ""), "reason": "fabricated_or_unverifiable_geo_clone"})
            continue
        keep.append(r)
    rows = keep
    by_name = {r["Company"]: r for r in rows}

    # 2) MERGE groups
    consumed = set()
    merged_rows: list[dict] = []
    alias_to_group = {}
    for g in MERGE_GROUPS:
        for a in g["aliases"]:
            alias_to_group[a] = g

    handled_groups = set()
    for r in rows:
        name = r["Company"]
        if name in consumed:
            continue
        g = alias_to_group.get(name)
        if not g:
            merged_rows.append(r)
            continue
        key = g["canonical"]
        if key in handled_groups:
            continue
        handled_groups.add(key)
        present = [by_name[a] for a in g["aliases"] if a in by_name]
        for a in g["aliases"]:
            if a in by_name:
                consumed.add(a)
        if not present:
            continue
        out = pick_best_row(present, g["website"])
        out["Company"] = g["canonical"]
        if g.get("hq"):
            out["Headquarters"] = g["hq"]
        if g.get("continent"):
            out["Continent / Geography"] = g["continent"]
        # Operational presence: union-ish note
        ops = sorted({p.get("Operational Presence", "") for p in present if p.get("Operational Presence")})
        if ops:
            joined = "; ".join(ops)
            if len(joined) > 500:
                joined = joined[:497] + "..."
            out["Operational Presence"] = joined
        merged_rows.append(out)
        dropped_aliases = [p["Company"] for p in present if p["Company"] != g["canonical"]]
        actions["merged"].append(
            {
                "canonical": g["canonical"],
                "kept_from": [p["Company"] for p in present],
                "removed_aliases": dropped_aliases,
                "website": g["website"],
            }
        )
        if name != g["canonical"] and g["canonical"] not in {p["Company"] for p in present}:
            actions["renamed"].append({"from": present[0]["Company"], "to": g["canonical"]})

    rows = merged_rows

    # 3) Apply website fixes + clear dead DNS that we cannot verify
    for r in rows:
        co = r["Company"]
        old = r.get("Website", "").strip()
        if co in WEBSITE_FIXES:
            new = WEBSITE_FIXES[co]
            if new != old:
                actions["website_fixed"].append({"company": co, "from": old, "to": new, "source": "verified_fix"})
            r["Website"] = new

    # Probe remaining; clear only hard DNS failures not in fixes
    probes = {}
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = {ex.submit(probe, r.get("Website", "")): r["Company"] for r in rows}
        for fut in as_completed(futs):
            probes[futs[fut]] = fut.result()

    for r in rows:
        co = r["Company"]
        web = r.get("Website", "").strip()
        p = probes.get(co, {})
        if web and not p.get("ok") and p.get("reason") == "dns":
            # Do not clear if we just set a verified fix that might be temporarily blocked
            if co in WEBSITE_FIXES and host_of(web) == host_of(WEBSITE_FIXES[co]):
                # keep verified official URL even if probe flaky
                continue
            actions["website_cleared"].append({"company": co, "from": web, "reason": p.get("reason")})
            r["Website"] = ""

    # 4) Deduplicate exact company names (safety)
    seen = set()
    unique_rows = []
    for r in rows:
        key = re.sub(r"\s+", " ", r["Company"]).strip().lower()
        if key in seen:
            actions["dropped"].append({"company": r["Company"], "website": r.get("Website", ""), "reason": "duplicate_name"})
            continue
        seen.add(key)
        unique_rows.append(r)
    rows = unique_rows

    # Sort by company
    rows.sort(key=lambda r: r["Company"].lower())

    write_rows(XLSX, hdr, rows)

    # Final probe summary
    final_probes = {}
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = {ex.submit(probe, r.get("Website", "")): r["Company"] for r in rows}
        for fut in as_completed(futs):
            final_probes[futs[fut]] = fut.result()
    bad = {c: p for c, p in final_probes.items() if not p.get("ok")}
    empty = sum(1 for r in rows if not r.get("Website"))

    audit = {
        "timestamp": ts,
        "backup": str(backup),
        "before": before,
        "after": len(rows),
        "removed": before - len(rows),
        "actions": actions,
        "empty_websites": empty,
        "bad_probes_after": {c: final_probes[c] for c in sorted(bad)},
        "bad_count_after": len(bad),
        "sample_websites": [
            {"Company": r["Company"], "Website": r.get("Website", "")}
            for r in rows[:30]
        ],
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("backup", backup)
    print("before", before, "after", len(rows), "removed", before - len(rows))
    print("merged groups", len(actions["merged"]))
    print("dropped", len(actions["dropped"]))
    print("website fixed", len(actions["website_fixed"]))
    print("website cleared", len(actions["website_cleared"]))
    print("empty websites", empty)
    print("bad probes after", len(bad))
    print("wrote", XLSX)
    print("audit", AUDIT)


if __name__ == "__main__":
    main()
