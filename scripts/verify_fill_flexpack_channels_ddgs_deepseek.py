# -*- coding: utf-8 -*-
"""Google AI scraper (fast path) + DDGS (fallback) + DeepSeek verify/fill/prune
for the Global Flexible Packaging Market channel-partner workbook.

For every row in
  output/web_expand/global_flexible_packaging_market_global/global_flexible_packaging_market_global_FINAL.xlsx

1. Asks the Google AI Overview scraper (dedicated instance on :15561) one
   combined question per company — much faster than multiple DDGS queries.
   Falls back to DDGS (DuckDuckGo/Bing/Brave/...) only if the scraper is down
   or returns nothing useful for that company.
2. Sends the company's existing Core Categories / Specialty Focus / Key Brands
   Represented + the search evidence to DeepSeek (OPENAI_API_KEY / OPENAI_BASE_URL
   in .env, already pointed at api.deepseek.com) and asks it to verify or correct
   all three fields (fills empties, fixes wrong values) using ONLY the evidence.
3. Asks DeepSeek to decide keep/remove: keep only if the company is genuinely a
   flexible-packaging DISTRIBUTOR / RESELLER / CHANNEL PARTNER / IMPORTER (not a
   manufacturer-only OEM, not an unrelated business), and its Key Brands (if any)
   are actually packaging-industry brands and not from an unrelated field.
4. Rows marked keep=false are removed from the sheet at the end.

No subagents — plain script, run directly:
  .venv\\Scripts\\python.exe scripts\\verify_fill_flexpack_channels_ddgs_deepseek.py
  .venv\\Scripts\\python.exe scripts\\verify_fill_flexpack_channels_ddgs_deepseek.py --limit 5   # smoke test

Requires the Google AI scraper backend + browser extension running and pointed
at GAI_BASE_URL (default http://127.0.0.1:15561):
  cd <google-ai-scraper server dir>
  uv run google-ai-scraper --backend --port 15561

Resumable: a checkpoint file records each company's decision, so re-running skips
already-processed companies. Row removal only happens once every current row has
a decision recorded.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import load_workbook

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")

from vendor_intel.clients.ddgs_engines import normalize_text_backends  # noqa: E402
from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping  # noqa: E402

WORKDIR = ROOT / "output" / "web_expand" / "global_flexible_packaging_market_global"
XLSX = WORKDIR / "global_flexible_packaging_market_global_FINAL.xlsx"
CKPT = WORKDIR / "_flexpack_ddgs_deepseek_verify_ckpt.json"
AUDIT = WORKDIR / "_flexpack_ddgs_deepseek_verify_audit.json"

NPD = "Not publicly disclosed"
NPD_SET = {"", "n/a", "na", "-", "—", "not publicly disclosed", "unknown", "tbd", "none", "null"}

MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
BACKEND_PARAM, BACKEND_NAMES, _DROPPED = normalize_text_backends()

# Dedicated Google AI scraper instance for this project (separate from any other
# project's :15551 instance) — override with GAI_BASE_URL env if needed.
GAI_BASE_URL = (os.getenv("GAI_BASE_URL") or "http://127.0.0.1:15561").rstrip("/")


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in NPD_SET


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

        with DDGS() as d:
            return list(d.text(query, backend=BACKEND_PARAM, max_results=max_results)) or []
    except Exception as e:  # noqa: BLE001
        print(f"    ddgs error ({query[:60]}...): {e}", flush=True)
        return []


def gather_evidence_ddgs(company: str, website: str) -> str:
    host = host_of(website)
    queries = [
        f'"{company}" flexible packaging distributor brands categories',
        f'"{company}" authorized distributor packaging manufacturer brands',
    ]
    if host:
        queries.append(f"site:{host} brands OR partners OR principals OR products OR categories")

    parts: list[str] = []
    for q in queries:
        hits = ddg_search(q, max_results=6)
        time.sleep(1.0)
        if not hits:
            continue
        lines = [
            f"- {h.get('title', '')}: {str(h.get('body', ''))[:280]} ({h.get('href', '')})"
            for h in hits
        ]
        parts.append(f"Query: {q}\n" + "\n".join(lines))
    return "\n\n".join(parts)[:9000]


_GAI_BAD_MARKERS = ("unusual traffic", "did not match", "ai overview is not available", "no ai overview")


async def gather_evidence_gai(company: str, website: str) -> str:
    host = host_of(website)
    q = (
        f"{company} {host} flexible packaging distributor — main product categories, "
        f"specialty focus, and manufacturer brands they distribute or represent"
    ).strip()
    try:
        data = await google_ai_ask(q, base_url=GAI_BASE_URL, llm_clean=True)
    except Exception as e:  # noqa: BLE001
        print(f"    gai error: {e}", flush=True)
        return ""
    md = str(data.get("markdown") or "").strip()
    if not md or any(m in md.lower() for m in _GAI_BAD_MARKERS):
        return ""
    return f"Google AI Overview for: {q}\n{md}"[:9000]


async def gather_evidence(company: str, website: str, gai_ok: bool) -> str:
    """Fast path: Google AI scraper (one call). Fallback: DDGS (multi-query)."""
    if gai_ok:
        gai_evidence = await gather_evidence_gai(company, website)
        if gai_evidence and len(gai_evidence) > 120:
            return gai_evidence
    return gather_evidence_ddgs(company, website)


def parse_json(text: str) -> dict:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I | re.S).strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
    return {}


def clean_brands(company: str, raw: str) -> str:
    if emptyish(raw):
        return NPD
    parts = [p.strip() for p in re.split(r"[;|/]", raw) if p.strip()]
    cname = company.strip().lower()
    c_tokens = set(re.findall(r"[a-z0-9]+", cname))
    cleaned: list[str] = []
    for p in parts:
        pl = p.lower().strip()
        if pl in NPD_SET or pl in {"others", "other", "etc", "etc.", "and more", "various"}:
            continue
        if re.search(r"\bmulti[- ]brand\b|\bprincipals\b|\bassortment\b", pl):
            continue
        if pl == cname or pl.startswith(cname + " "):
            continue
        ptoks = set(re.findall(r"[a-z0-9]+", pl))
        if c_tokens and ptoks and ptoks <= c_tokens:
            continue
        if re.search(r"[A-Za-z]", p) and len(p) >= 2:
            cleaned.append(p)
    out: list[str] = []
    seen = set()
    for b in cleaned:
        k = b.lower()
        if k not in seen:
            seen.add(k)
            out.append(b)
    return "; ".join(out) if out else NPD


SYSTEM = """You are auditing ONE row of a GLOBAL FLEXIBLE PACKAGING MARKET channel-partner \
spreadsheet. A valid row is a genuine DISTRIBUTOR / RESELLER / CHANNEL PARTNER / IMPORTER \
of flexible packaging materials, films, laminates, pouches, bags, labels, or closely \
related packaging supplies/machinery/consumables — NOT a packaging manufacturer-only OEM \
/ brand owner, and NOT an unrelated-industry business.

Using ONLY the company info and the web-search evidence given (real search snippets — \
title/body/url triples), do the following for THIS company:

1. "core_categories": verify or correct the main packaging product categories this \
distributor sells (semicolon-separated short labels, 3-8 items). Keep the existing value \
if the evidence supports it; correct it if the evidence clearly contradicts it. Never \
invent categories unsupported by evidence or the existing value.

2. "specialty_focus": verify or correct one short phrase (<= 22 words) naming the market / \
technology / product / customer segment this distributor is known for. No company-name \
restatement, no trailing country/region.

3. "key_brands": verify or correct named manufacturer/brand names this company \
distributes, represents, or is an authorized reseller of. Use a brand name ONLY if it is \
explicitly evidenced in the search snippets OR already present in the existing value AND \
plausibly a packaging-industry brand. NEVER invent. If the only brands you can find are \
from a clearly unrelated industry (e.g. food brands, electronics, apparel, automotive with \
no packaging connection), do not include them — treat as none evidenced. If nothing is \
evidenced, return "Not publicly disclosed".

4. "keep": true only if this company is genuinely a flexible-packaging DISTRIBUTOR / \
CHANNEL PARTNER / RESELLER / IMPORTER, and (when any key brands exist) those brands are \
actually related to flexible packaging / packaging materials — not from an unrelated \
industry. Set keep=false with a short "reason" if: it is a manufacturer-only OEM, a pure \
logistics/3PL company, an unrelated business, or its only evidenced brands belong to a \
different industry with no packaging distribution role.

Return STRICT JSON only, no prose:
{"core_categories":"...","specialty_focus":"...","key_brands":"...","keep":true,"reason":"short"}"""


def verify_row(
    client: OpenAI, company: str, website: str, core: str, specialty: str, brands: str, evidence: str
) -> dict:
    user = (
        f"Company: {company}\n"
        f"Website: {website or '(none)'}\n"
        f"Existing Core Categories: {core or '(empty)'}\n"
        f"Existing Specialty Focus: {specialty or '(empty)'}\n"
        f"Existing Key Brands Represented: {brands or '(empty)'}\n\n"
        f"SEARCH EVIDENCE:\n{evidence or '(no search hits found)'}"
    )
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                temperature=0,
                max_tokens=350,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": SYSTEM},
                    {"role": "user", "content": user},
                ],
            )
            data = parse_json(resp.choices[0].message.content or "{}")
            if data:
                return data
        except Exception as e:  # noqa: BLE001
            print(f"    deepseek error (attempt {attempt + 1}): {e}", flush=True)
            time.sleep(2.0 * (attempt + 1))
    return {}


def main() -> int:
    return asyncio.run(main_async())


async def main_async() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="Process only first N unprocessed rows (smoke test)")
    args = ap.parse_args()

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    if not api_key:
        print("ERROR: OPENAI_API_KEY missing in .env", file=sys.stderr)
        return 1

    client = OpenAI(api_key=api_key, base_url=base, timeout=75.0)
    smoke = client.chat.completions.create(model=MODEL, messages=[{"role": "user", "content": "OK"}], max_tokens=3)
    print(f"DeepSeek OK: {smoke.choices[0].message.content!r} model={smoke.model}", flush=True)
    print(f"DDGS backends (fallback): {BACKEND_NAMES}", flush=True)

    gai_ok = False
    try:
        gai_ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        gai_ok = False
    print(f"Google AI scraper @ {GAI_BASE_URL}: {'UP (fast path)' if gai_ok else 'DOWN (using DDGS fallback only)'}", flush=True)

    if not XLSX.exists():
        print(f"ERROR: not found: {XLSX}", file=sys.stderr)
        return 1

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = XLSX.with_name(XLSX.stem + f"_BEFORE_DDGS_DEEPSEEK_VERIFY_{stamp}.xlsx")
    if not bak.exists():
        bak.write_bytes(XLSX.read_bytes())
        print(f"backup -> {bak.name}", flush=True)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    # Fixed row/company map captured BEFORE any edits — used for the final removal pass.
    row_company: list[tuple[int, str]] = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co:
            row_company.append((r, co))
    print(f"total rows: {len(row_company)}", flush=True)

    ckpt: dict = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass
    done = ckpt.setdefault("done", {})

    todo = [(r, co) for r, co in row_company if co not in done]
    if args.limit:
        todo = todo[: args.limit]
    print(f"to process this run: {len(todo)} (already done: {len(row_company) - len(todo)})", flush=True)

    filled_brands = corrected_core = corrected_specialty = 0
    for i, (r, co) in enumerate(todo, 1):
        website = str(ws.cell(r, ic["Website"]).value or "").strip()
        core = str(ws.cell(r, ic["Core Categories"]).value or "").strip()
        specialty = str(ws.cell(r, ic["Specialty Focus"]).value or "").strip()
        brands = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()

        print(f"[{i}/{len(todo)}] {co}", flush=True)
        evidence = await gather_evidence(co, website, gai_ok)
        result = verify_row(client, co, website, core, specialty, brands, evidence)
        time.sleep(0.3)

        new_core = str(result.get("core_categories") or "").strip()
        new_specialty = str(result.get("specialty_focus") or "").strip()
        new_brands_raw = str(result.get("key_brands") or "").strip()
        keep = bool(result.get("keep", True))
        reason = str(result.get("reason") or "").strip()

        if new_core and new_core.lower() != core.lower() and len(new_core) > 3:
            ws.cell(r, ic["Core Categories"]).value = new_core
            corrected_core += 1

        if new_specialty and new_specialty.lower() != specialty.lower() and len(new_specialty) >= 8:
            ws.cell(r, ic["Specialty Focus"]).value = new_specialty
            corrected_specialty += 1

        new_brands = clean_brands(co, new_brands_raw) if new_brands_raw else NPD
        if new_brands.lower() != brands.lower():
            ws.cell(r, ic["Key Brands Represented"]).value = new_brands
            if emptyish(brands) and not emptyish(new_brands):
                filled_brands += 1

        done[co] = {
            "keep": keep,
            "reason": reason,
            "core": new_core or core,
            "specialty": new_specialty or specialty,
            "brands": new_brands,
        }

        status = "KEEP" if keep else f"REMOVE ({reason[:60]})"
        print(f"    {status} | core={new_core[:50] or core[:50]!r} brands={new_brands[:50]!r}", flush=True)

        if i % 10 == 0:
            wb.save(XLSX)
            CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
            print(f"  checkpoint saved at {i}/{len(todo)}", flush=True)

    wb.save(XLSX)
    CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    # Removal pass — only once every current row has a recorded decision.
    all_decided = all(co in done for _, co in row_company)
    removed: list[dict] = []
    if all_decided:
        remove_rows = sorted((r for r, co in row_company if not done[co].get("keep", True)), reverse=True)
        for r in remove_rows:
            co = str(ws.cell(r, ic["Company"]).value or "").strip()
            removed.append({"company": co, "reason": done.get(co, {}).get("reason", "")})
            ws.delete_rows(r, 1)
        wb.save(XLSX)
        print(f"Removed {len(removed)} non-channel / unrelated-brand rows.", flush=True)
    else:
        remaining = sum(1 for _, co in row_company if co not in done)
        print(f"Not all rows decided yet ({remaining} left) — skipping removal this run. Re-run to continue.", flush=True)

    audit = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "model": MODEL,
        "ddgs_backends": BACKEND_NAMES,
        "processed_this_run": len(todo),
        "corrected_core": corrected_core,
        "corrected_specialty": corrected_specialty,
        "filled_or_corrected_brands": filled_brands,
        "removal_pass_ran": all_decided,
        "removed_count": len(removed),
        "removed": removed,
        "final_row_count": ws.max_row - 1,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: v for k, v in audit.items() if k != "removed"}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
