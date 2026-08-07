#!/usr/bin/env python3
"""Second pass: fill remaining Employees/Turnover blanks for ADC FINAL."""
from __future__ import annotations

import asyncio
import json
import os
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import httpx
from dotenv import load_dotenv
from openai import OpenAI

from fill_adc_employees_turnover import (
    load_rows,
    write_xlsx,
    empty,
    norm,
    growjo_lookup,
    wikipedia_lookup,
    apply_values,
    parse_employees,
    parse_revenue,
    MODEL,
)

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")
OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"
SUMMARY = OUT / "adc_fill_employees_turnover_pass2_summary.json"

SYSTEM = """Fill missing Employees and Turnover/Scale for pharmaceutical / life-science distributors.
Use public knowledge and evidence. If precise data unknown, give a reasonable SCALE BAND and mark (est.):
- Employees examples: 1-10, 11-50, 51-200, 201-500, 501-1,000, 1,000-5,000, 5,000+
- Turnover examples: <$5M (est.), $5-25M (est.), $25-100M (est.), $100-500M (est.), $500M-$1B (est.), $1B+ (est.), or parent-group note
Do not invent fake exact headcounts like 137. Prefer bands when unsure.
If truly unknown even as band, use Not publicly disclosed.
Return JSON {"results":[{"company":"...","employees":"...","turnover":"...","note":"..."}]}"""


def deepseek_batch(client: OpenAI, rows: list[dict[str, str]], evidence: list[dict]) -> list[dict]:
    payload = []
    for r, ev in zip(rows, evidence):
        payload.append({
            "company": r["Company"],
            "hq": r.get("Headquarters"),
            "website": r.get("Website"),
            "business_type": r.get("Business Type"),
            "specialty_focus": (r.get("Specialty Focus") or "")[:120],
            "current_employees": r.get("Employees"),
            "current_turnover": r.get("Turnover / Scale"),
            "web_evidence": ev,
        })
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM},
                    {"role": "user", "content": json.dumps({"companies": payload})},
                ],
                temperature=0.15,
                max_tokens=3500,
                response_format={"type": "json_object"},
            )
            data = json.loads(resp.choices[0].message.content or "{}")
            if isinstance(data, dict):
                for k in ("results", "companies", "items"):
                    if isinstance(data.get(k), list):
                        return data[k]
            if isinstance(data, list):
                return data
            return []
        except Exception as e:
            print("err", e)
            time.sleep(1.2 * (attempt + 1))
    return []


async def main() -> int:
    rows = load_rows(XLSX)
    need = [
        i for i, r in enumerate(rows)
        if empty(r.get("Employees")) or empty(r.get("Turnover / Scale"))
        or (r.get("Employees") or "").lower() == "not publicly disclosed"
        or (r.get("Turnover / Scale") or "").lower() == "not publicly disclosed"
    ]
    print(f"need pass2: {len(need)} / {len(rows)}")
    oa = OpenAI(api_key=os.getenv("OPENAI_API_KEY"), base_url=os.getenv("OPENAI_BASE_URL"))
    updated = 0

    async with httpx.AsyncClient(
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0"},
        timeout=httpx.Timeout(10.0, connect=4.0),
    ) as http:
        for i in range(0, len(need), 10):
            idxs = need[i : i + 10]
            chunk = [rows[j] for j in idxs]
            print(f"[{i+1}-{i+len(idxs)}/{len(need)}]", flush=True)
            evids = []
            for r in chunk:
                try:
                    g, w = await asyncio.wait_for(
                        asyncio.gather(growjo_lookup(http, r["Company"]), wikipedia_lookup(http, r["Company"])),
                        timeout=20.0,
                    )
                except asyncio.TimeoutError:
                    g, w = {}, {}
                evids.append({"growjo": g, "wikipedia": w})
            decisions = deepseek_batch(oa, chunk, evids)
            by = {norm(str(d.get("company") or "")): d for d in decisions if d.get("company")}
            for j, r, ev in zip(idxs, chunk, evids):
                d = by.get(norm(r["Company"])) or {}
                emp = str(d.get("employees") or "").strip()
                turn = str(d.get("turnover") or "").strip()
                # prefer web numbers if present
                if ev.get("growjo", {}).get("employees"):
                    emp = ev["growjo"]["employees"]
                if ev.get("growjo", {}).get("turnover"):
                    turn = ev["growjo"]["turnover"]
                if ev.get("wikipedia", {}).get("employees") and (empty(emp) or "est" in emp.lower() or "-" in emp):
                    # wiki exact can override band
                    emp = ev["wikipedia"]["employees"]
                if ev.get("wikipedia", {}).get("turnover"):
                    turn = ev["wikipedia"]["turnover"]

                # only overwrite NPD / empty
                cur_e = r.get("Employees") or ""
                cur_t = r.get("Turnover / Scale") or ""
                new_e = emp if emp and not empty(emp) else cur_e
                new_t = turn if turn and not empty(turn) else cur_t
                if empty(cur_e) or cur_e.lower() == "not publicly disclosed":
                    if emp and not empty(emp):
                        new_e = emp
                if empty(cur_t) or cur_t.lower() == "not publicly disclosed":
                    if turn and not empty(turn):
                        new_t = turn
                if new_e != cur_e or new_t != cur_t:
                    rows[j] = apply_values(r, new_e, new_t, "adc_emp_turnover_pass2")
                    updated += 1
                    print(f"  UP {r['Company'][:36]} | {new_e[:24]} | {new_t[:28]}", flush=True)
                else:
                    # still blank -> keep NPD explicitly
                    rows[j] = apply_values(r, cur_e or "Not publicly disclosed", cur_t or "Not publicly disclosed", "adc_emp_turnover_pass2")
            await asyncio.sleep(0.15)

    rows.sort(key=lambda x: norm(x["Company"]))
    write_xlsx(XLSX, rows)
    emp_f = sum(1 for r in rows if not empty(r.get("Employees")) and r.get("Employees", "").lower() != "not publicly disclosed")
    turn_f = sum(1 for r in rows if not empty(r.get("Turnover / Scale")) and r.get("Turnover / Scale", "").lower() != "not publicly disclosed")
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "updated": updated,
        "employees_filled": emp_f,
        "turnover_filled": turn_f,
        "rows": len(rows),
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"DONE updated={updated} emp_filled={emp_f} turn_filled={turn_f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
