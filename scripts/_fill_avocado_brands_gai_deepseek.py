#!/usr/bin/env python3
"""Apply web-verified Key Brands + Core/Focus, then Google AI + DeepSeek fill remaining."""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

# Windows consoles are often cp1252 — force UTF-8 so GAI client logs don't crash fills.
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
os.environ.setdefault("PYTHONUTF8", "1")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import openpyxl
from dotenv import load_dotenv
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")

os.environ.setdefault("GOOGLE_AI_SCRAPER_ENABLED", "true")
os.environ.setdefault("GOOGLE_AI_SCRAPER_BASE_URL", "http://127.0.0.1:15551")
os.environ.setdefault("GOOGLE_AI_SCRAPER_LLM_CLEAN", "true")
os.environ.setdefault("SEARCH_STACK_OPENAI_RESIDUAL", "true")

XLSX = ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_OFFICE_SLIM.xlsx"
CACHE = ROOT / "output/web_expand/global_avocado_oil_market_global/_avocado_brands_gai_cache.json"
AUDIT = ROOT / "output/web_expand/global_avocado_oil_market_global/_avocado_brands_gai_audit.json"
NPD = "Not publicly disclosed"

# Web-verified brand overrides (agent web search this session)
VERIFIED = {
    "Brakes": {
        "brands": "Sysco Essentials, Sysco Classic, Sysco Premium, La Boulangerie, Portico, Wholesome Farms, House Recipe, Citavo, Earth Plus, Country Choice, others",
    },
    "Brake Bros": {
        "brands": "Sysco Essentials, Sysco Classic, Sysco Premium, La Boulangerie, Portico, Wholesome Farms, House Recipe, Citavo, Earth Plus, Country Choice, others",
    },
    "US Foods": {
        "brands": "Chef's Line, Metro Deli, Rykoff Sexton, Stock Yards, Cross Valley Farms, Monarch, Monogram, Glenview Farms, Harbor Banks, Roseli, Molly's Kitchen, Harvest Value, Hilltop Hearth, others",
    },
    "US Foods CHEF'STORE": {
        "brands": "Chef's Line, Metro Deli, Rykoff Sexton, Stock Yards, Cross Valley Farms, Monarch, Monogram, others",
    },
    "Walmart": {
        "brands": "Great Value, Marketside, Sam's Choice, others",
    },
    "Gordon Food Service": {
        "brands": "Markon First Crop, Ready-Set-Serve, Halperns', GFS, others",
    },
    "Cash & Carry Smart Foodservice": {
        "brands": "Chef's Line, Metro Deli, Rykoff Sexton, Stock Yards, Cross Valley Farms, Monarch, Monogram, others",
    },
    "Trader Joe's": {
        "brands": "Trader Joe's, others",
    },
    "Associated Wholesale Grocers": {
        "brands": "Always Save, Best Choice, Clearly by Best Choice, Superior Selections by Best Choice, Pure Wonder by Best Choice, Hello World by Best Choice, others",
    },
    "Bestway Wholesale": {
        "brands": "Best-in, others",
    },
    "Performance Food Group": {
        "brands": "Peak Fresh Produce, others",
    },
    "PFG Customized Distribution": {
        "brands": "Peak Fresh Produce, others",
    },
    "Ben E. Keith Foods": {
        "brands": "Markon First Crop, Ready-Set-Serve, Markon Essentials, others",
    },
    "Shamrock Foods": {
        "brands": "Markon First Crop, Ready-Set-Serve, Markon Essentials, others",
    },
    "Nicholas and Company": {
        "brands": "Markon First Crop, Ready-Set-Serve, Markon Essentials, others",
    },
    "KeHE Distributors": {
        "brands": "CADIA, MADE WITH, others",
    },
    "KeHE Specialty": {
        "brands": "CADIA, MADE WITH, others",
    },
    "Wakefern Food Corp": {
        "brands": "Bowl & Basket, Wholesome Pantry, Paperbird, others",
    },
    "SpartanNash": {
        "brands": "Our Family, Open Acres, Fresh & Finest by Our Family, others",
    },
    "Metro AG": {
        "brands": "METRO Chef, Aro, METRO Professional, Fine Life, Rioba, METRO Premium, others",
    },
    "Makro": {
        "brands": "METRO Chef, Aro, METRO Professional, Fine Life, Rioba, others",
    },
    "Transgourmet": {
        "brands": "Transgourmet Economy, Transgourmet Quality, Transgourmet Ursprung, Transgourmet Natura, Transgourmet Premium, others",
    },
    "Restaurant Depot": {
        "brands": "Chef's Quality, Sabor Nuestro, Big C, Supremo Italiano, Skyline, James Farm, others",
    },
    "Jetro Restaurant Depot": {
        "brands": "Chef's Quality, Sabor Nuestro, Big C, Supremo Italiano, Skyline, James Farm, others",
    },
    "Sligro Food Group": {
        "brands": "Dinnerland, Meesterschap van het huis, Kern, Sanissimo, others",
    },
    "Associated Food Stores": {
        "brands": "Food Club, Full Circle, TopCare, Buckley Farms, Culinary Tours, BirchBerry, Crav'n Flavor, others",
    },
    "Affiliated Foods": {
        "brands": "IGA, others",
    },
    "Bozzuto's": {
        "brands": "IGA, Hy-Top, Seven Farms, others",
    },
    "C&S Wholesale Grocers": {
        "brands": "Best Yet, Our Family, Fresh & Finest by Our Family, IGA, others",
    },
    "Alex Lee": {
        "brands": "Brown Bag, IGA, others",
    },
    "Chefs' Warehouse": {
        "brands": "BelAria Essentials, Grand Reserve, Allen Brothers, Proper Stock, others",
    },
}


def ensure_others(text: str) -> str:
    t = (text or "").strip().rstrip(",")
    if not t:
        return t
    if re.search(r",\s*others\s*$", t, re.I):
        return re.sub(r",\s*others\s*$", ", others", t, flags=re.I)
    return t + ", others"


def clean_brand_list(raw: str) -> str:
    """Normalize comma/semicolon brand lists; drop junk tokens."""
    if not raw:
        return ""
    parts = re.split(r"[,;/|]+", raw)
    out = []
    seen = set()
    junk = {
        "others",
        "private label",
        "house brands",
        "various",
        "multiple",
        "n/a",
        "none",
        "unknown",
        "brands",
        "brand",
        "principal",
        "suppliers",
        "portfolio",
    }
    for p in parts:
        b = re.sub(r"\s+", " ", p).strip(" .")
        if not b or len(b) < 2:
            continue
        if b.lower() in junk:
            continue
        if re.search(r"principal|private label|house brand|not publicly|inventory|assortment", b, re.I):
            continue
        if len(b) > 60:
            continue
        key = b.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(b)
    if not out:
        return ""
    return ensure_others(", ".join(out[:12]))


def ds_client() -> tuple[OpenAI, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        raise SystemExit("OPENAI_API_KEY missing")
    return OpenAI(api_key=key, base_url=base, timeout=90.0), model


def deepseek_extract_brands(client: OpenAI, model: str, company: str, evidence: str) -> dict:
    system = (
        "Extract ONLY real consumer/trade brand names that this company owns or officially "
        "distributes. Prefer private-label / house brands and named principal brands. "
        "Do NOT invent brands. Do NOT return generic phrases like 'private label oils' or "
        "'principal brands'. If evidence does not clearly name brands, found=false. "
        "Return ONLY JSON."
    )
    user = (
        f"Company: {company}\n\nEVIDENCE:\n{(evidence or '')[:9000]}\n\n"
        'JSON: {"found":true|false,"brands":["Brand A","Brand B"],'
        '"core_categories":"","specialty_focus":"","confidence":"high|medium|low",'
        '"reason":"brief"}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=400,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        raw = resp.choices[0].message.content or "{}"
        if raw.startswith("```"):
            raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I | re.S).strip()
        data = json.loads(raw)
    except Exception as ex:
        return {"found": False, "error": str(ex)[:120]}

    conf = str(data.get("confidence") or "").lower()
    if not data.get("found") or conf not in {"high", "medium"}:
        return {"found": False, "confidence": conf, "raw": data}

    brands = data.get("brands") or []
    if isinstance(brands, str):
        brands = [brands]
    brands = [str(b).strip() for b in brands if str(b).strip()]
    # evidence gate: at least one brand token should appear in evidence (case-insensitive)
    evid_l = (evidence or "").lower()
    kept = []
    for b in brands:
        # allow slight punctuation differences
        token = re.sub(r"[^a-z0-9]", "", b.lower())
        evid_tok = re.sub(r"[^a-z0-9]", "", evid_l)
        if len(token) >= 4 and token in evid_tok:
            kept.append(b)
        elif conf == "high" and len(b) >= 3:
            # high confidence still requires alphanumeric overlap of 4+ chars somewhere
            if any(w.lower() in evid_l for w in re.findall(r"[A-Za-z0-9][A-Za-z0-9&''\-]{2,}", b)):
                kept.append(b)
    if not kept:
        return {"found": False, "confidence": conf, "reason": "brands not evidenced"}

    out = {
        "found": True,
        "brands": clean_brand_list(", ".join(kept)),
        "confidence": conf,
    }
    core = str(data.get("core_categories") or "").strip()
    focus = str(data.get("specialty_focus") or "").strip()
    # only accept core/focus if short and no country names
    country = re.compile(
        r"\b(United States|USA|UK|Germany|France|Canada|Mexico|Australia|China|India)\b",
        re.I,
    )
    if core and len(core) < 160 and not country.search(core):
        out["core"] = ensure_others(core)
    if focus and len(focus) < 180 and not country.search(focus):
        out["focus"] = ensure_others(focus)
    return out


async def fill_one(company: str, website: str, client: OpenAI, model: str) -> dict:
    from vendor_intel.clients.google_ai_scraper import google_ai_ask

    queries = [
        f'key brands represented by "{company}" private label OR house brands OR exclusive brands',
        f'"{company}" own brand names foodservice OR grocery distributor',
    ]
    evidence_parts = []
    for q in queries:
        data = await google_ai_ask(q, llm_clean=True)
        md = str(data.get("markdown") or "").strip()
        if md:
            evidence_parts.append(md[:4500])
        if evidence_parts and len("\n".join(evidence_parts)) > 1200:
            break
        await asyncio.sleep(0.4)

    evidence = "\n---\n".join(evidence_parts)
    if website and website.lower() not in {"not publicly disclosed", "n/a"}:
        evidence = f"Website: {website}\n" + evidence

    if len(evidence) < 80:
        return {"found": False, "reason": "no_evidence"}

    return deepseek_extract_brands(client, model, company, evidence)


async def run_gai(limit: int | None = None) -> None:
    client, model = ds_client()
    cache: dict = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}

    # Phase A: verified overrides
    applied = 0
    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        if co not in VERIFIED:
            continue
        hit = VERIFIED[co]
        if hit.get("brands"):
            ws.cell(row, ic["Key Brands Represented"]).value = hit["brands"]
            applied += 1
            print(f"verified {co}", flush=True)
    wb.save(XLSX)
    print(f"phaseA verified={applied}", flush=True)

    # Phase B: Google AI + DeepSeek for remaining NPD
    gaps = []
    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        web = str(ws.cell(row, ic["Website"]).value or "").strip()
        brands = str(ws.cell(row, ic["Key Brands Represented"]).value or "").strip()
        if brands in ("", NPD):
            gaps.append({"row": row, "company": co, "website": web})
    if limit:
        gaps = gaps[:limit]
    print(f"phaseB gaps={len(gaps)} model={model}", flush=True)

    filled = 0
    for i, g in enumerate(gaps, 1):
        key = f"brands::{g['company']}"
        if key in cache and cache[key].get("done"):
            hit = cache[key]
            if hit.get("brands"):
                ws.cell(g["row"], ic["Key Brands Represented"]).value = hit["brands"]
                if hit.get("core"):
                    ws.cell(g["row"], ic["Core Categories"]).value = hit["core"]
                if hit.get("focus"):
                    ws.cell(g["row"], ic["Specialty Focus"]).value = hit["focus"]
                filled += 1
            print(f"[{i}/{len(gaps)}] cache {g['company'][:40]}", flush=True)
            continue

        print(f"[{i}/{len(gaps)}] GAI+DS {g['company'][:42]}", flush=True)
        try:
            hit = await fill_one(g["company"], g["website"], client, model)
        except Exception as ex:
            hit = {"found": False, "error": str(ex)[:160]}

        cache[key] = {
            "done": True,
            "ts": datetime.now(timezone.utc).isoformat(),
            **hit,
        }
        if hit.get("found") and hit.get("brands"):
            ws.cell(g["row"], ic["Key Brands Represented"]).value = hit["brands"]
            if hit.get("core"):
                ws.cell(g["row"], ic["Core Categories"]).value = hit["core"]
            if hit.get("focus"):
                ws.cell(g["row"], ic["Specialty Focus"]).value = hit["focus"]
            filled += 1
            print(f"  -> {hit['brands'][:100]}", flush=True)
        else:
            print(f"  -> NPD ({hit.get('reason') or hit.get('error') or hit.get('confidence')})", flush=True)

        if i % 10 == 0:
            CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
            wb.save(XLSX)

        await asyncio.sleep(0.6)

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    wb.save(XLSX)
    wb.close()

    # recount
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c or "") for c in rows[1]]
    ic2 = {h: i for i, h in enumerate(hdr) if h}
    real = sum(
        1
        for r in rows[2:]
        if str(r[ic2["Key Brands Represented"]] or "").strip() not in ("", NPD)
    )
    wb.close()
    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "phaseA_verified": applied,
        "phaseB_attempted": len(gaps),
        "phaseB_filled": filled,
        "total_real_brands": real,
        "total_rows": len(rows) - 2,
        "model": model,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(json.dumps(audit, indent=2), flush=True)


def main() -> None:
    limit = None
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        limit = int(sys.argv[1])
    asyncio.run(run_gai(limit=limit))


if __name__ == "__main__":
    main()
