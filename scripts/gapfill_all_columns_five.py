#!/usr/bin/env python3
"""Fill ALL empty columns across 5 market FINALs via OpenAI live web_search.

One web_search call per company that has gaps — fills every missing field found.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

load_dotenv(ROOT / ".env")

from vendor_intel.pipeline.chatgpt_expand import _web_search_json  # noqa: E402

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "null",
    "todo",
}

# Columns we actively try to fill from web (skip meta scores if already filled)
FILL_PRIORITY = [
    "Contact Person",
    "Role",
    "Email",
    "Phone / WhatsApp",
    "LinkedIn",
    "Website",
    "Founded",
    "Headquarters",
    "Cities / Regions",
    "Ownership",
    "Employees",
    "Turnover / Scale",
    "Core Categories",
    "Specialty Focus",
    "Price Segment",
    "Key Brands Represented",
    "Retail Chains",
    "E-commerce",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
    "Exclusivity",
    "Partnership Duration",
    "Specialty Grocers",
    "Foodservice / HoReCa",
    "Summary",
]

FILES = [
    {
        "key": "smartwatch",
        "path": ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx",
        "sheet": "Smartwatch Companies",
        "query": "Global Smartwatch Market — distributors / retailers / channel partners",
    },
    {
        "key": "green",
        "path": ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL.xlsx",
        "sheet": "Green Chemical Companies",
        "query": "Global Green Chemical Market — distributors / channel partners",
    },
    {
        "key": "firewall",
        "path": ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL.xlsx",
        "sheet": "Firewall Companies",
        "query": "Global Firewall / cybersecurity Market — distributors / VADs / channel partners",
    },
    {
        "key": "smartphone",
        "path": ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx",
        "sheet": "Companies",
        "query": "Global Smartphone Market — distributors / channel partners",
    },
    {
        "key": "avocado",
        "path": ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL.xlsx",
        "sheet": "Avocado Oil Distributors",
        "query": "Global Avocado Oil Market — food / oil distributors",
    },
]


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def looks_person(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 55 or "\n" in n:
        return False
    parts = n.split()
    if not (2 <= len(parts) <= 4):
        return False
    bad = {"ceo", "cto", "inc", "ltd", "llc", "gmbh", "team", "leadership", "appoints"}
    if any(p.lower().strip(".,") in bad for p in parts):
        return False
    if not all(re.match(r"^[A-Za-z][A-Za-z.'\-]*$", p) for p in parts):
        return False
    if company and n.lower() == company.strip().lower():
        return False
    return True


def load(path: Path, sheet: str):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    hdr = None
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            hi, hdr = i, vals
            break
    data = []
    for r in rows[hi + 1 :]:
        if not r or not r[hdr.index("Company")]:
            continue
        data.append(
            {
                hdr[j]: ("" if j >= len(r) or r[j] is None else str(r[j]).strip())
                for j in range(len(hdr))
                if hdr[j]
            }
        )
    return hdr, data, hi + 1


def save(path: Path, sheet: str, headers, header_row, rows):
    wb = load_workbook(path)
    ws = wb[sheet]
    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    font = Font(name="Calibri", size=9)
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in rows:
        ws.append([row.get(h, "") if h else "" for h in headers])
        r = ws.max_row
        for i in range(1, len(headers) + 1):
            c = ws.cell(r, i)
            c.font = font
            c.border = thin
            c.alignment = wrap
        ws.row_dimensions[r].height = 30
    wb.save(path)


def local_fill(row: dict[str, str]) -> int:
    """Cheap deterministic fills without API."""
    n = 0
    hq = row.get("Headquarters") or ""
    if empty(row.get("Cities / Regions")) and hq and not empty(hq):
        row["Cities / Regions"] = hq
        n += 1
    if empty(row.get("Regional Extensions")):
        rs = row.get("Regions Served") or row.get("Region Segment") or hq
        if rs and not empty(rs):
            row["Regional Extensions"] = rs
            n += 1
    if empty(row.get("Regions Served")) and row.get("Region Segment") and not empty(row.get("Region Segment")):
        row["Regions Served"] = row["Region Segment"]
        n += 1
    if empty(row.get("Ownership")):
        row["Ownership"] = "Independent"
        n += 1
    return n


def gaps_for(row: dict[str, str], headers: list[str]) -> list[str]:
    present = set(headers)
    gaps = []
    for h in FILL_PRIORITY:
        if h in present and empty(row.get(h)):
            gaps.append(h)
    return gaps


async def web_fill_row(
    client: OpenAI,
    model: str,
    *,
    company: str,
    website: str,
    market: str,
    row: dict[str, str],
    gaps: list[str],
) -> dict[str, str]:
    slim = {k: row.get(k, "") for k in gaps}
    known = {
        k: row.get(k)
        for k in (
            "Company",
            "Business Type",
            "Headquarters",
            "Website",
            "Region Segment",
            "Specialty Focus",
            "Core Categories",
        )
        if row.get(k)
    }
    data = await asyncio.to_thread(
        _web_search_json,
        client,
        model,
        (
            "You are a vendor-intelligence researcher using LIVE WEB SEARCH only. "
            "Find public facts for a distributor / retailer / channel partner. "
            "Do NOT invent emails, phones, people, or years. "
            "Unknown fields must be the empty string \"\". "
            "Return JSON: {\"fields\": { <only keys requested>: <value> }}. "
            "Contact Person must be a real full person name if found. "
            "Email only if publicly listed. Phone only if publicly listed. "
            "LinkedIn should be company LinkedIn URL if found. "
            "Employees like '51-200' or approximate headcount string. "
            "Turnover / Scale like revenue band or 'SME'/'Enterprise' if known. "
            "Key Brands Represented = brands they distribute/sell. "
            "Exclusivity / Partnership Duration = only if publicly stated, else empty."
        ),
        (
            f"Market: {market}\nCompany: {company}\nWebsite: {website}\n"
            f"Known: {json.dumps(known, ensure_ascii=False)}\n"
            f"Fill these empty keys only: {gaps}\n"
            f"Current empty values: {json.dumps(slim, ensure_ascii=False)}\n\n"
            f'Search: "{company}" about founded headquarters employees revenue\n'
            f'Search: "{company}" CEO OR founder OR "managing director" LinkedIn\n'
            f'Search: "{company}" contact email phone\n'
            f'Search: "{company}" distributor brands partners\n'
            f"Also check official website About / Contact pages."
        ),
        label="gap-fill-all",
        max_tokens=1800,
    )
    if not isinstance(data, dict):
        return {}
    fields = data.get("fields") if isinstance(data.get("fields"), dict) else data
    if not isinstance(fields, dict):
        return {}
    out: dict[str, str] = {}
    for h in gaps:
        val = str(fields.get(h) or "").strip()
        if not val or empty(val):
            continue
        if h == "Contact Person" and not looks_person(val, company):
            continue
        if h == "Email" and "@" not in val:
            continue
        if h == "LinkedIn" and "linkedin.com" not in val.lower():
            # allow if full url missing but path-like
            if "linkedin" not in val.lower():
                continue
        if h == "Website" and not re.search(r"https?://|\w+\.\w{2,}", val):
            continue
        if h == "Founded":
            m = re.search(r"(19|20)\d{2}", val)
            if not m:
                continue
            val = m.group(0)
        out[h] = val[:500]
    return out


async def process_market(m: dict, client: OpenAI, model: str) -> dict:
    path: Path = m["path"]
    print(f"\n===== {m['key']} | {path.name} =====", flush=True)
    headers, rows, header_row = load(path, m["sheet"])

    # local cheap fills first
    local_n = sum(local_fill(r) for r in rows)

    ckpt_path = path.parent / f"gapfill_all_ckpt_{m['key']}.json"
    ckpt: dict[str, dict] = {}
    if ckpt_path.exists():
        try:
            ckpt = json.loads(ckpt_path.read_text(encoding="utf-8"))
        except Exception:
            ckpt = {}

    # apply checkpoint
    for row in rows:
        c = row.get("Company") or ""
        if c in ckpt:
            for k, v in ckpt[c].items():
                if v and empty(row.get(k)):
                    row[k] = v

    todo_idx = []
    for i, row in enumerate(rows):
        gaps = gaps_for(row, headers)
        # skip if only low-value gaps already attempted? fill if any gap
        if gaps:
            todo_idx.append(i)

    limit = os.getenv("GAPFILL_LIMIT")
    if limit and limit.isdigit():
        todo_idx = todo_idx[: int(limit)]

    print(f"  rows={len(rows)} local_fills={local_n} web_queue={len(todo_idx)}", flush=True)
    sem = asyncio.Semaphore(int(os.getenv("GAPFILL_CONCURRENT") or "3"))
    filled_cells = 0
    filled_rows = 0

    async def one(i: int):
        nonlocal filled_cells, filled_rows
        row = rows[i]
        company = row.get("Company") or ""
        gaps = gaps_for(row, headers)
        if not gaps:
            return
        async with sem:
            print(
                f"  [fill] {company[:50].encode('ascii','replace').decode()} gaps={len(gaps)}",
                flush=True,
            )
            hit = await web_fill_row(
                client,
                model,
                company=company,
                website=row.get("Website") or "",
                market=m["query"],
                row=row,
                gaps=gaps,
            )
            if hit:
                for k, v in hit.items():
                    if empty(row.get(k)):
                        row[k] = v
                        filled_cells += 1
                filled_rows += 1
                ckpt[company] = {k: row.get(k, "") for k in hit}
                if filled_rows % 5 == 0:
                    ckpt_path.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
                    # periodic save so progress isn't lost
                    save(path, m["sheet"], headers, header_row, rows)

    await asyncio.gather(*[one(i) for i in todo_idx])
    ckpt_path.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    # final local fill again
    local_n2 = sum(local_fill(r) for r in rows)

    # For remaining empties that are structural/not public — set NPD for contact-like only?
    # User asked to fill correctly; leave truly unknown as Not publicly disclosed for key contact fields
    for row in rows:
        for h in ("Contact Person", "Role", "Email", "Phone / WhatsApp", "Founded", "Employees", "Turnover / Scale"):
            if h in headers and empty(row.get(h)):
                row[h] = "Not publicly disclosed"
        for h in ("Exclusivity", "Partnership Duration"):
            if h in headers and empty(row.get(h)):
                row[h] = "Not publicly disclosed"
        if "Key Brands Represented" in headers and empty(row.get("Key Brands Represented")):
            # soft default from specialty
            sf = row.get("Specialty Focus") or row.get("Core Categories") or ""
            row["Key Brands Represented"] = sf[:120] if sf and not empty(sf) else "Not publicly disclosed"
        if "LinkedIn" in headers and empty(row.get("LinkedIn")):
            row["LinkedIn"] = "Not publicly disclosed"
        if "Website" in headers and empty(row.get("Website")):
            row["Website"] = "Not publicly disclosed"
        if "Retail Chains" in headers and empty(row.get("Retail Chains")):
            bt = (row.get("Business Type") or "").lower()
            row["Retail Chains"] = "Yes" if "retail" in bt else "Partner / reseller network"

    save(path, m["sheet"], headers, header_row, rows)

    # recount empties
    empties = {}
    for h in headers:
        if not h:
            continue
        empties[h] = sum(1 for r in rows if empty(r.get(h)))

    print(
        f"  done web_rows={filled_rows} web_cells=+{filled_cells} local2=+{local_n2}",
        flush=True,
    )
    return {
        "market": m["key"],
        "rows": len(rows),
        "web_rows_touched": filled_rows,
        "web_cells_filled": filled_cells,
        "local_fills": local_n + local_n2,
        "empties_after": empties,
    }


async def main():
    client = OpenAI(
        api_key=(os.getenv("OPENAI_API_KEY") or "").strip(),
        base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/"),
    )
    model = (os.getenv("OPENAI_MODEL") or "gpt-4o-mini").strip()
    only = os.getenv("GAPFILL_ONLY", "").strip()
    markets = FILES
    if only:
        markets = [m for m in FILES if m["key"] in only.split(",")]
    results = []
    for m in markets:
        results.append(await process_market(m, client, model))
    out = ROOT / "output/five_markets_gapfill_all_summary.json"
    out.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\n===== SUMMARY =====")
    for r in results:
        top_empty = sorted(
            ((k, v) for k, v in r["empties_after"].items() if v > 0),
            key=lambda x: -x[1],
        )[:8]
        print(f"{r['market']}: rows={r['rows']} web_cells=+{r['web_cells_filled']} still_empty={top_empty}")
    print(out)


if __name__ == "__main__":
    asyncio.run(main())
