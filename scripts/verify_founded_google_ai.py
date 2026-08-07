#!/usr/bin/env python3
"""Verify doubtful Founded years via Google AI scraper.

Uses the same search phrasing as manual Google: \"{company} founded year?\"
Requires google-ai-scraper backend on :15551 + Chrome extension connected.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")
os.environ["GOOGLE_AI_SCRAPER_ENABLED"] = "true"

from vendor_intel.clients.google_ai_scraper import (  # noqa: E402
    google_ai_ask,
    google_ai_ping,
)

SLIM_FILES = [
    ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market_FINAL_SLIM.xlsx",
]

OUT_JSON = ROOT / "scripts/_founded_google_ai_fixes.json"
CURRENT_YEAR = 2026


def extract_years(text: str) -> list[int]:
    years = []
    for m in re.finditer(r"\b((?:18|19|20)\d{2})\b", text or ""):
        y = int(m.group(1))
        if 1800 <= y <= CURRENT_YEAR:
            years.append(y)
    return years


def pick_founded_year(markdown: str, company: str) -> tuple[int | None, str]:
    """Prefer 'established/founded in YYYY' over copyright years."""
    text = markdown or ""
    low = text.lower()
    patterns = [
        rf"(?:was\s+)?(?:established|founded|incorporated|started|formed)\s+(?:in\s+)?(\d{{4}})",
        rf"established[:\s]+(\d{{4}})",
        rf"founded[:\s]+(\d{{4}})",
        rf"since\s+(\d{{4}})",
    ]
    for pat in patterns:
        m = re.search(pat, low, re.I)
        if m:
            y = int(m.group(1))
            if 1800 <= y <= CURRENT_YEAR:
                # ignore pure copyright context
                window = low[max(0, m.start() - 20) : m.end() + 20]
                if "©" in window or "copyright" in window:
                    continue
                return y, f"matched:{pat}"
    # fallback: first year that isn't CURRENT_YEAR (copyright contamination)
    years = extract_years(text)
    years = [y for y in years if y != CURRENT_YEAR]
    if years:
        # prefer years mentioned near company name / established words
        return years[0], "fallback_first_year"
    return None, "no_year"


def collect_doubtful() -> list[dict]:
    rows: list[dict] = []
    for f in SLIM_FILES:
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        hdr = None
        hr = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            if row and any(str(c or "").strip().lower() == "company" for c in row):
                hdr = [str(c or "") for c in row]
                hr = i
                break
        assert hdr and hr
        ci, fi = hdr.index("Company"), hdr.index("Founded")
        hi = hdr.index("Headquarters") if "Headquarters" in hdr else None
        wi = hdr.index("Website") if "Website" in hdr else None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or not row[ci]:
                continue
            company = str(row[ci]).strip()
            founded = str(row[fi] or "").strip()
            hq = str(row[hi] or "") if hi is not None else ""
            web = str(row[wi] or "") if wi is not None else ""
            m = re.search(r"(18|19|20)\d{2}", founded)
            y = int(m.group(0)) if m else None
            # Doubtful = recent years often polluted by copyright / scrape year
            if y is not None and y >= 2020:
                rows.append(
                    {
                        "file": str(f),
                        "company": company,
                        "founded": founded,
                        "year": y,
                        "hq": hq,
                        "website": web,
                    }
                )
        wb.close()
    return rows


async def verify_one(item: dict) -> dict:
    company = item["company"]
    # Exact phrasing as user's Google search bar
    query = f"{company} founded year?"
    data = await google_ai_ask(query, llm_clean=True)
    md = str(data.get("markdown") or "").strip()
    raw = str(data.get("markdown_raw") or "").strip()
    text = md or raw
    new_year, how = pick_founded_year(text, company)
    out = {
        "file": item["file"],
        "company": company,
        "old": item["founded"],
        "old_year": item["year"],
        "query": query,
        "query_used": data.get("overview_query") or data.get("query_used"),
        "ai_overview_missing": data.get("ai_overview_missing"),
        "error": data.get("error"),
        "snippet": text[:500],
        "new_year": new_year,
        "extract": how,
        "action": "keep",
        "new": None,
    }
    if new_year is None:
        # if old is current-year copyright pollution and AI gave nothing useful → NPD
        if item["year"] == CURRENT_YEAR:
            out["action"] = "fix"
            out["new"] = "NPD"
            out["extract"] = "2026_invalid_no_ai_year"
        return out
    if new_year != item["year"]:
        out["action"] = "fix"
        out["new"] = str(new_year)
    else:
        out["action"] = "keep"
        out["new"] = str(new_year)
    return out


def apply_fixes(fixes: list[dict]) -> None:
    by_file: dict[str, dict[str, str]] = {}
    for fx in fixes:
        if fx.get("action") != "fix" or not fx.get("new"):
            continue
        key = re.sub(r"\s+", " ", fx["company"].strip().lower())
        by_file.setdefault(fx["file"], {})[key] = str(fx["new"])

    for fpath, mapping in by_file.items():
        wb = load_workbook(fpath)
        ws = wb.active
        hdr = None
        hr = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            if row and any(str(c or "").strip().lower() == "company" for c in row):
                hdr = [str(c or "") for c in row]
                hr = i
                break
        assert hdr and hr
        ci, fi = hdr.index("Company"), hdr.index("Founded")
        n = 0
        for r in range(hr + 1, ws.max_row + 1):
            company = ws.cell(r, ci + 1).value
            if not company:
                continue
            key = re.sub(r"\s+", " ", str(company).strip().lower())
            if key in mapping:
                ws.cell(r, fi + 1).value = mapping[key]
                n += 1
        wb.save(fpath)
        print(f"Applied {n} fixes -> {Path(fpath).name}")


async def main() -> None:
    ok = await google_ai_ping()
    print(f"Google AI scraper health: {ok}")
    if not ok:
        raise SystemExit("Google AI scraper not reachable on :15551 — start backend + extension")

    doubtful = collect_doubtful()
    print(f"Doubtful founded years (>=2020): {len(doubtful)}\n")

    fixes: list[dict] = []
    for i, item in enumerate(doubtful, 1):
        print(f"[{i}/{len(doubtful)}] {item['company']} ({item['founded']}) -> ask: {item['company']} founded year?", flush=True)
        try:
            fx = await verify_one(item)
        except Exception as e:
            fx = {
                "file": item["file"],
                "company": item["company"],
                "old": item["founded"],
                "action": "keep",
                "error": str(e),
                "new": None,
            }
        fixes.append(fx)
        print(
            f"    {fx.get('action')}: {fx.get('old')} => {fx.get('new') or fx.get('old')} "
            f"| year={fx.get('new_year')} | {fx.get('extract')} | missing={fx.get('ai_overview_missing')}",
            flush=True,
        )
        # be gentle with Google / extension queue
        await asyncio.sleep(1.5)

    OUT_JSON.write_text(json.dumps(fixes, indent=2), encoding="utf-8")
    changed = [f for f in fixes if f.get("action") == "fix"]
    print(f"\nChanges: {len(changed)} / {len(fixes)}")
    apply_fixes(fixes)
    print(f"Wrote {OUT_JSON}")


if __name__ == "__main__":
    asyncio.run(main())
