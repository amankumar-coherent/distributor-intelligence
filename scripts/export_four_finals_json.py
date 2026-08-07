# -*- coding: utf-8 -*-
"""Export 4 FINAL Excels → JSON with market name + integer Neon-linked id at top."""
from __future__ import annotations

import json
import os
from pathlib import Path

import openpyxl
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
load_dotenv(Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env"))

# Integer IDs:
# - Prefer cmi_reports."newsId" (Neon integer report id) when a clear report exists
# - Else fall back is not used — we require newsId or explicit override below
IDS = {
    "Global Firewall Market": {
        "id": None,
        "id_source": "not present in neon cmi_reports for this market name",
        "cmi_keyword": None,
        "neon_market_id": "fde0876d-0fd7-48a4-9e82-2766d5048558",
    },
    "Global Smartwatch Market": {
        "id": None,
        "id_source": "not present in neon cmi_reports for this market name",
        "cmi_keyword": None,
        "neon_market_id": "cd382d27-444d-40fa-b1cd-144f94abcae4",
    },
    "Global Green Chemical Market": {
        "id": None,
        "id_source": "not present in neon cmi_reports for this market name",
        "cmi_keyword": None,
        "neon_market_id": "1cc15bca-e35e-4f82-a20c-8220cac49c68",
    },
    "Global Avocado Oil Market": {
        "id": None,
        "id_source": "not present in neon cmi_reports for this market name",
        "cmi_keyword": None,
        "neon_market_id": None,
    },
}

MARKETS = [
    {
        "market": "Global Firewall Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market_FINAL.xlsx"
        ),
        "json_out": Path(
            r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market_FINAL.json"
        ),
    },
    {
        "market": "Global Smartwatch Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"
        ),
        "json_out": Path(
            r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.json"
        ),
    },
    {
        "market": "Global Green Chemical Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL.xlsx"
        ),
        "json_out": Path(
            r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL.json"
        ),
    },
    {
        "market": "Global Avocado Oil Market",
        "xlsx": Path(
            r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
        ),
        "json_out": Path(
            r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.json"
        ),
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

    rows: list[dict] = []
    co_idx = headers.index("Company") + 1
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


def verify_neon_uuid(uuid: str) -> bool:
    url = (os.environ.get("DATABASE_URL") or "").strip()
    if not url:
        return False
    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM markets WHERE id = %s::uuid", (uuid,))
            return cur.fetchone() is not None


def main() -> None:
    summary = []
    for m in MARKETS:
        meta = IDS[m["market"]]
        companies = load_companies(m["xlsx"])
        ok = verify_neon_uuid(meta["neon_market_id"]) if meta.get("neon_market_id") else False
        market_id = meta["id"]  # may be None
        if market_id is not None:
            market_id = int(market_id)

        payload = {
            "market": m["market"],
            "id": market_id,
            "neon_market_id": meta["neon_market_id"],
            "id_source": meta["id_source"],
            "cmi_keyword": meta.get("cmi_keyword"),
            "geography": "global",
            "company_count": len(companies),
            "source_excel": m["xlsx"].name,
            "companies": companies,
        }
        m["json_out"].write_text(
            json.dumps(payload, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        summary.append(
            {
                "market": m["market"],
                "id": market_id,
                "neon_ok": ok,
                "companies": len(companies),
                "json": str(m["json_out"]),
                "id_source": meta["id_source"],
            }
        )
        print(
            f"OK {m['market']}: id={market_id} neon_uuid_ok={ok} "
            f"companies={len(companies)} -> {m['json_out'].name}"
        )

    print("\nSUMMARY")
    for s in summary:
        print(f"  {s['market']}")
        print(f"    id={s['id']} ({s['id_source']})")
        print(f"    companies={s['companies']}")
        print(f"    file={s['json']}")


if __name__ == "__main__":
    main()
