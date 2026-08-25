# -*- coding: utf-8 -*-
"""Export 7 market Excels to Desktop/final json.

Heading mentions market name once + Neon markets.id (UUID) or id:null.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import openpyxl
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))

OUT_DIR = Path(r"C:\Users\aman.kumar\Desktop\final json")

MARKETS = [
    {
        "market": "Global Electrical Distribution Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\web_expand\global_electrical_distribution_market_global\global_electrical_distribution_market_global_FINAL.xlsx"
        ),
        "json_name": "global_electrical_distribution_market.json",
    },
    {
        "market": "Global Avocado Oil Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market.xlsx"
        ),
        "json_name": "global_avocado_oil_market.json",
    },
    {
        "market": "Global Firewall Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market.xlsx"
        ),
        "json_name": "global_firewall_market.json",
    },
    {
        "market": "Global Green Chemical Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market.xlsx"
        ),
        "json_name": "global_green_chemical_market.json",
    },
    {
        "market": "Global Smartphone Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\web_expand\global_smartphone_market_global\global_smartphone_market.xlsx"
        ),
        "json_name": "global_smartphone_market.json",
    },
    {
        "market": "Global Antibody Drug Conjugates Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\web_expand\global_antibody_drug_conjugates_market_global\global_antibody_drug_conjugates_market.xlsx"
        ),
        "json_name": "global_antibody_drug_conjugates_market.json",
    },
    {
        "market": "Global Smartwatch Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market.xlsx"
        ),
        "json_name": "global_smartwatch_market.json",
    },
]


def find_header(ws) -> tuple[int, list[str]]:
    for r in range(1, 6):
        vals = [c.value for c in ws[r]]
        labels = [str(v or "").strip() for v in vals]
        if "Company" in labels:
            return r, labels
    raise RuntimeError("Company header not found")


def load_companies(xlsx: Path) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = hdr = None
    headers: list[str] = []
    for name in wb.sheetnames:
        try:
            hdr, headers = find_header(wb[name])
            ws = wb[name]
            break
        except RuntimeError:
            continue
    if ws is None or hdr is None:
        raise RuntimeError(f"No company sheet in {xlsx}")
    co_idx = headers.index("Company") + 1
    rows: list[dict] = []
    for r in range(hdr + 1, ws.max_row + 1):
        company = ws.cell(r, co_idx).value
        if company is None or not str(company).strip():
            continue
        d: dict = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            v = ws.cell(r, i + 1).value
            if v is None:
                d[h] = ""
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                d[h] = v
            else:
                d[h] = str(v).strip()
        rows.append(d)
    wb.close()
    return rows


def neon_market_id(name: str) -> int | None:
    """Integer CMI report id (cmi_reports.newsId). null if market is not in Neon."""
    url = (os.environ.get("DATABASE_URL") or "").strip()
    if not url:
        return None
    candidates = [name]
    if name.lower().startswith("global "):
        candidates.append(name[7:].strip())
    # catalog uses plural for smartwatch
    if "smartwatch" in name.lower() and "smartwatches" not in name.lower():
        candidates.append(name.replace("Smartwatch", "Smartwatches").replace("smartwatch", "smartwatches"))
        if name.lower().startswith("global "):
            candidates.append(name[7:].strip().replace("Smartwatch", "Smartwatches"))
    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            for cand in candidates:
                cur.execute(
                    """
                    SELECT "newsId"
                    FROM cmi_reports
                    WHERE LOWER(COALESCE(keyword,'')) = LOWER(%s)
                    LIMIT 1
                    """,
                    (cand,),
                )
                row = cur.fetchone()
                if row:
                    return int(row[0])
    return None


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for m in MARKETS:
        mid = neon_market_id(m["market"])
        companies = load_companies(m["xlsx"])
        payload = {
            "market": m["market"],
            "id": mid,
            "companies": companies,
        }
        out = OUT_DIR / m["json_name"]
        out.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        office = sum(1 for c in companies if str(c.get("Office No.") or "").strip())
        print(
            f"JSON {m['market']} id={mid} companies={len(companies)} "
            f"office={office} -> {out}"
        )


if __name__ == "__main__":
    main()
