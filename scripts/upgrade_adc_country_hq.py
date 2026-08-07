#!/usr/bin/env python3
"""Upgrade remaining country-only HQ rows via DeepSeek."""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

from fix_adc_hq_cities_founded import (
    COUNTRY_ONLY,
    load_rows,
    write_xlsx,
    is_country_only,
    sanitize_hq,
    sanitize_cities,
)

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")
OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"
MODEL = os.getenv("OPENAI_MODEL", "deepseek-chat")

SYSTEM = """Upgrade country-only headquarters to city-level if publicly known.
Rules:
- Headquarters = ONE city HQ only (City, Country preferred). Never multiple countries.
- Cities / Regions = HQ city + other branches, separated by '; '.
- Founded = YYYY if known.
- If city truly unknown, keep country as Headquarters.
Return JSON: {"results":[{"company":"...","founded":"YYYY","headquarters":"...","cities_regions":"...","note":"..."}]}"""


def main() -> int:
    rows = load_rows(XLSX)
    targets = [r for r in rows if is_country_only(r.get("Headquarters") or "")]
    print(f"country-only targets: {len(targets)}")
    client = OpenAI(
        api_key=os.getenv("OPENAI_API_KEY"),
        base_url=os.getenv("OPENAI_BASE_URL", "https://api.deepseek.com/v1"),
    )
    upgraded = 0
    for i in range(0, len(targets), 10):
        chunk = targets[i : i + 10]
        payload = [
            {
                "company": r["Company"],
                "website": r.get("Website"),
                "founded": r.get("Founded"),
                "hq": r.get("Headquarters"),
                "cities": r.get("Cities / Regions"),
                "regions": r.get("Regions Served"),
            }
            for r in chunk
        ]
        print(f"[{i+1}-{i+len(chunk)}] DeepSeek...", flush=True)
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM},
                {"role": "user", "content": json.dumps({"companies": payload})},
            ],
            temperature=0.05,
            max_tokens=3500,
            response_format={"type": "json_object"},
        )
        data = json.loads(resp.choices[0].message.content or "{}")
        results = []
        if isinstance(data, dict):
            for k in ("results", "companies", "items"):
                if isinstance(data.get(k), list):
                    results = data[k]
                    break
            if not results and "company" in data:
                results = [data]
        elif isinstance(data, list):
            results = data
        dmap = {
            re.sub(r"\s+", " ", (x.get("company") or "").strip().lower()): x
            for x in results
            if x.get("company")
        }
        for r in chunk:
            d = dmap.get(r["Company"].strip().lower())
            if not d:
                continue
            hq = sanitize_hq(str(d.get("headquarters") or ""))
            cities = sanitize_cities(str(d.get("cities_regions") or d.get("cities") or ""), hq)
            founded = str(d.get("founded") or r.get("Founded") or "").strip()
            if not hq:
                continue
            if is_country_only(hq) and hq.lower() == (r.get("Headquarters") or "").lower():
                print(f"  KEEP {r['Company']} ({hq})")
                continue
            if not is_country_only(hq) or ("," in hq):
                r["Headquarters"] = hq
                r["Cities / Regions"] = cities or hq
                if re.match(r"^(18|19|20)\d{2}$", founded):
                    r["Founded"] = founded
                src = r.get("Data Sources") or ""
                if "adc_country_hq_upgrade" not in src:
                    r["Data Sources"] = (src + "; adc_country_hq_upgrade").strip("; ")
                upgraded += 1
                print(f"  UP   {r['Company']} -> {hq} | {cities[:60]}")
            else:
                print(f"  KEEP {r['Company']} ({hq})")
        time.sleep(0.15)

    # sync mutations back — targets are refs into rows? Actually targets are same dict refs if from rows
    # rebuild: targets were slices of rows list elements (same dicts) — OK
    rows.sort(key=lambda x: x["Company"].lower())
    write_xlsx(XLSX, rows)
    left = sum(1 for r in rows if is_country_only(r.get("Headquarters") or ""))
    print(f"DONE upgraded={upgraded} country-only left={left} city-level={len(rows)-left}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
