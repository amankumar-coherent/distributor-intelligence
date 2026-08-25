# -*- coding: utf-8 -*-
"""
Phase 2: real web-search (DuckDuckGo via ddgs) pass to fill Key Brands Represented
for flexible-packaging distributors still "Not publicly disclosed" after the
site-scrape-only verification pass (verify_flexpack_core_specialty_brands.py).

Unlike phase 1 (which only fetches each company's own website), this queries a
search engine for third-party mentions of brand partnerships: trade press,
"authorized distributor" listings, LinkedIn/company directories, etc.

Only accepts brand names that literally appear in the search-result snippets
(never invents names). Run AFTER phase 1 finishes (do not run concurrently —
both scripts save the same workbook).
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

WORKDIR = ROOT / "output" / "web_expand" / "global_flexible_packaging_market_global"
XLSX = WORKDIR / "global_flexible_packaging_market_global_FINAL.xlsx"
AUDIT = WORKDIR / "_flexpack_websearch_brands_audit.json"
CKPT = WORKDIR / "_flexpack_websearch_brands_ckpt.json"
NPD = "Not publicly disclosed"


def host_of(website: str) -> str:
    w = (website or "").strip()
    if not w:
        return ""
    if not w.startswith("http"):
        w = "https://" + w
    try:
        h = urlparse(w).netloc
        return h[4:] if h.startswith("www.") else h
    except Exception:
        return ""


def ddg_search(query: str, max_results: int = 6) -> list[dict]:
    try:
        from ddgs import DDGS

        return list(DDGS().text(query, max_results=max_results)) or []
    except Exception:
        return []


def gather_search_evidence(company: str, website: str) -> str:
    host = host_of(website)
    queries = []
    if host:
        queries.append(f'site:{host} brands OR partners OR principals OR manufacturers OR "we carry"')
    queries.append(f'"{company}" authorized distributor brands')
    queries.append(f'"{company}" distributes OR carries OR represents brand packaging')

    parts = []
    for q in queries:
        hits = ddg_search(q, max_results=6)
        time.sleep(0.4)
        if not hits:
            continue
        lines = [
            f"- {h.get('title', '')}: {str(h.get('body', ''))[:280]} ({h.get('href', '')})"
            for h in hits
        ]
        parts.append(f"Query: {q}\n" + "\n".join(lines))
    return "\n\n".join(parts)


def parse_json(text: str) -> dict:
    text = (text or "").strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{[\s\S]*\}", text)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
    return {}


SYSTEM_PROMPT = (
    "You extract manufacturer/brand names a packaging CHANNEL DISTRIBUTOR carries, "
    "represents, or is an authorized reseller of, using ONLY the search-result "
    "snippets given below (title/snippet/url triples from a real web search).\n"
    "RULES:\n"
    "- List a brand ONLY if it is explicitly named in the snippet text as something this "
    "specific company distributes/carries/represents/is authorized for.\n"
    "- NEVER invent or guess brand names. NEVER list the distributor's own company name.\n"
    "- NEVER list generic words (packaging, products, supplies, materials) as brands.\n"
    "- If snippets don't clearly name any brand for this exact company, return empty.\n"
    'Return strict JSON: {"brands_found": "Brand1; Brand2; ...", "confidence": '
    '"high|medium|low|none", "notes": "short reason"}'
)


def deepseek_extract_brands(api_key: str, base: str, model: str, company: str, website: str, evidence: str) -> dict:
    import httpx

    user = (
        f"Company: {company}\nWebsite: {website}\n\n"
        f"Search evidence:\n{evidence[:7000] if evidence else '(no search hits)'}"
    )
    try:
        with httpx.Client(timeout=60.0) as c:
            r = c.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0,
                    "max_tokens": 300,
                    "messages": [
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": user},
                    ],
                    "response_format": {"type": "json_object"},
                },
            )
            r.raise_for_status()
            return parse_json(r.json()["choices"][0]["message"]["content"])
    except Exception as e:
        return {"confidence": "none", "notes": f"error: {e}"}


def brand_in_evidence(brand: str, evidence: str) -> bool:
    b = brand.lower().strip()
    if not b:
        return False
    return b in evidence.lower() or b.split()[0] in evidence.lower()


def clean_brands(raw: str, evidence: str) -> str:
    if not raw:
        return ""
    parts = re.split(r"[;,/|]+", raw)
    out, seen = [], set()
    junk = {"others", "various", "multiple", "n/a", "none", "unknown", "brands", "brand", "packaging"}
    for p in parts:
        b = re.sub(r"\s+", " ", p).strip(" .")
        if not b or len(b) < 2 or b.lower() in junk or len(b) > 60:
            continue
        if not brand_in_evidence(b, evidence):
            continue
        key = b.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(b)
    return "; ".join(out[:10])


def main() -> None:
    api_key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not api_key:
        raise SystemExit("Missing DeepSeek/OpenAI API key in .env")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = XLSX.with_name(XLSX.stem + f"_BEFORE_WEBSEARCH_BRANDS_{stamp}" + XLSX.suffix)
    shutil.copy2(XLSX, bak)
    print(f"backup -> {bak.name}", flush=True)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    ckpt = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    targets = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        cur = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        if not cur or cur.lower() == NPD.lower():
            targets.append(r)
    print(f"targets needing brand web-search={len(targets)}", flush=True)

    filled = 0
    processed = 0
    for i, r in enumerate(targets, 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in ckpt.get("done", {}):
            continue
        website = str(ws.cell(r, ic["Website"]).value or "").strip()

        evidence = gather_search_evidence(co, website)
        ds = deepseek_extract_brands(api_key, base, model, co, website, evidence)
        conf = str(ds.get("confidence") or "").lower()
        brands = clean_brands(str(ds.get("brands_found") or ""), evidence)

        applied = False
        if conf in {"high", "medium"} and brands:
            ws.cell(r, ic["Key Brands Represented"]).value = brands
            filled += 1
            applied = True

        processed += 1
        ckpt.setdefault("done", {})[co] = {"applied": applied, "confidence": conf, "brands": brands}

        if processed % 5 == 0 or i <= 5:
            print(f"[{i}/{len(targets)}] {co} conf={conf} applied={applied} brands={brands[:70]}", flush=True)
        if processed % 10 == 0:
            CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
            wb.save(XLSX)
            print(f"checkpoint saved at processed={processed}", flush=True)

    wb.save(XLSX)
    CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    npd_final = 0
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        if not v or v.lower() == NPD.lower():
            npd_final += 1

    audit = {
        "backup": bak.name,
        "targets": len(targets),
        "processed_this_run": processed,
        "filled_this_run": filled,
        "npd_final": npd_final,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(audit, indent=2, ensure_ascii=False))
    print("DONE", flush=True)


if __name__ == "__main__":
    main()
