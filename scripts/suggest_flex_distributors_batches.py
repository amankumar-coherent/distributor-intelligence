"""DeepSeek suggest packaging distributors in small batches; write extra candidates JSON."""
from __future__ import annotations

import json
import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
OUT = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\add100_extra_candidates.json"
)
EXIST = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\existing_names.txt"
)

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

existing = EXIST.read_text(encoding="utf-8").splitlines() if EXIST.exists() else []
avoid = ", ".join(existing[:60])

all_companies = []
regions = [
    "United States regional packaging distributors/wholesalers (not Uline/Veritiv/Imperial Dade/Bunzl/Grainger/SupplyOne/Berlin/TricorBraun)",
    "Canada and Mexico packaging distributors/wholesalers",
    "United Kingdom and Ireland packaging merchants/distributors",
    "Germany Netherlands Belgium France packaging wholesalers (not RAJA/Manutan/Ratioform/Lyreco)",
    "Australia New Zealand Singapore packaging distributors/wholesalers",
    "India Japan South Korea packaging distributors/channel partners",
]

for region in regions:
    prompt = f"""List 20 REAL packaging DISTRIBUTOR/wholesaler/merchant companies in: {region}.
Only channel companies (not manufacturers/OEMs/hybrids). Real official websites only.
Avoid already listed: {avoid}
JSON: {{"companies":[{{"company":"","website":"https://","hq":"","country_code":"","founded":"YYYY or Not publicly disclosed","core_categories":"","specialty_focus":"","summary":""}}]}}"""
    try:
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": "Return compact JSON only. Real distributors with real websites."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
            response_format={"type": "json_object"},
            max_tokens=3500,
        )
        data = json.loads(resp.choices[0].message.content or "{}")
        comps = data.get("companies") or []
        print(region[:40], "->", len(comps), flush=True)
        for c in comps:
            c["_region_query"] = region
            all_companies.append(c)
    except Exception as e:
        print("FAIL", region[:40], e, flush=True)

OUT.write_text(json.dumps(all_companies, indent=2), encoding="utf-8")
print("total", len(all_companies), "wrote", OUT)
