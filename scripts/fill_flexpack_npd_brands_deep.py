# -*- coding: utf-8 -*-
"""Second, deeper pass: recover Key Brands Represented for rows still marked
"Not publicly disclosed" after verify_fill_flexpack_channels_ddgs_deepseek.py.

Only targets rows where Key Brands Represented is currently empty/NPD. For each,
gathers MORE evidence than the first pass:
  1. Direct website scrape (About / Brands / Partners / Suppliers / Principals
     pages on the company's own site) — not just search snippets.
  2. DDGS web search (extra brand-focused query variants).
  3. Google AI Overview scraper, with FULL retry count (this pool is much
     smaller than the full 332 rows, so it can afford to be slower/thorough).

Combines all evidence and asks DeepSeek to extract ONLY real, evidenced
manufacturer/brand names (never invents). Does not touch keep/remove or
Core Categories / Specialty Focus — those were already finalized by the first
pass.

No subagents — plain script, run directly (after the first pass has completed
and removed non-channel rows):
  .venv\\Scripts\\python.exe scripts\\fill_flexpack_npd_brands_deep.py
  .venv\\Scripts\\python.exe scripts\\fill_flexpack_npd_brands_deep.py --limit 5   # smoke test
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
from urllib.parse import urljoin, urlparse

import httpx
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
CKPT = WORKDIR / "_flexpack_npd_brands_deep_ckpt.json"
AUDIT = WORKDIR / "_flexpack_npd_brands_deep_audit.json"

NPD = "Not publicly disclosed"
NPD_SET = {"", "n/a", "na", "-", "—", "not publicly disclosed", "unknown", "tbd", "none", "null"}
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
BACKEND_PARAM, BACKEND_NAMES, _DROPPED = normalize_text_backends()
GAI_BASE_URL = (os.getenv("GAI_BASE_URL") or "http://127.0.0.1:15551").rstrip("/")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


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


def strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html or "")
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def scrape_site(website: str) -> str:
    url = (website or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    texts: list[str] = []
    paths = ["", "/brands", "/partners", "/suppliers", "/principals", "/about", "/about-us"]
    try:
        with httpx.Client(
            headers={"User-Agent": UA}, verify=False, timeout=httpx.Timeout(12.0, connect=4.0), follow_redirects=True
        ) as client:
            for path in paths:
                try:
                    u = url if not path else urljoin(url.rstrip("/") + "/", path.lstrip("/"))
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    t = strip_html(r.text)
                    if len(t) > 120:
                        texts.append(t[:3500])
                    if len(texts) >= 3:
                        break
                except Exception:
                    continue
    except Exception:
        return ""
    return "\n---\n".join(texts)[:9000]


def ddg_search(query: str, max_results: int = 6) -> list[dict]:
    try:
        from ddgs import DDGS

        with DDGS() as d:
            return list(d.text(query, backend=BACKEND_PARAM, max_results=max_results)) or []
    except Exception as e:  # noqa: BLE001
        print(f"    ddgs error ({query[:60]}...): {e}", flush=True)
        return []


def gather_ddgs_evidence(company: str, website: str) -> str:
    host = host_of(website)
    queries = [
        f'"{company}" brands OR principals OR "authorized distributor"',
        f'"{company}" "we carry" OR "we represent" OR "our brands" packaging',
        f'"{company}" distributes OR represents manufacturer brand',
        # ownership / acquisition / family-of-brands angle — different phrasing
        # sometimes surfaces parent-company or sister-brand relationships that
        # a plain "brands they carry" query misses.
        f'"{company}" acquired by OR "part of" OR "subsidiary of" OR "family of brands"',
        f'"{company}" rebrand OR "now known as" OR "formerly" packaging',
    ]
    if host:
        queries.append(f"site:{host} brands OR partners OR principals OR suppliers")
    parts: list[str] = []
    for q in queries:
        hits = ddg_search(q, max_results=6)
        time.sleep(0.8)
        if not hits:
            continue
        lines = [
            f"- {h.get('title', '')}: {str(h.get('body', ''))[:280]} ({h.get('href', '')})" for h in hits
        ]
        parts.append(f"Query: {q}\n" + "\n".join(lines))
    return "\n\n".join(parts)[:9000]


_GAI_BAD_MARKERS = ("unusual traffic", "did not match", "ai overview is not available", "no ai overview")


async def gather_gai_evidence(company: str, website: str, gai_ok: bool) -> str:
    if not gai_ok:
        return ""
    host = host_of(website)
    q = f"{company} {host} packaging brands distributed authorized distributor manufacturer partners suppliers"
    try:
        data = await google_ai_ask(q, base_url=GAI_BASE_URL, llm_clean=True)
    except Exception as e:  # noqa: BLE001
        print(f"    gai error: {e}", flush=True)
        return ""
    md = str(data.get("markdown") or "").strip()
    if not md or any(m in md.lower() for m in _GAI_BAD_MARKERS):
        return ""
    return f"Google AI Overview for: {q}\n{md}"[:9000]


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


SYSTEM = """Extract Key Brands Represented for a FLEXIBLE PACKAGING channel distributor.
Definition: named manufacturer/brand names whose products this company officially
distributes, resells, represents, or is an authorized dealer/reseller of.
Rules:
- Use a brand name ONLY if it is explicitly evidenced in the EVIDENCE text below
  (website content, search snippets, or AI overview).
- Named brands only. No distributor self-name. No "others"/"various"/"multi-brand".
- Never invent. If the evidence's only brand-like names are from an unrelated
  industry (not packaging/materials/machinery), do not use them.
- If nothing is clearly evidenced, brands="".
Return STRICT JSON only: {"brands":"Brand A; Brand B","confidence":"high|medium|low|none"}"""


def deepseek_brands(client: OpenAI, company: str, evidence: str) -> str:
    user = f"Company: {company}\n\nEVIDENCE:\n{evidence[:9000] if evidence else '(no evidence found)'}"
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                temperature=0,
                max_tokens=200,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": SYSTEM},
                    {"role": "user", "content": user},
                ],
            )
            data = parse_json(resp.choices[0].message.content or "{}")
            conf = str(data.get("confidence") or "").lower()
            brands = str(data.get("brands") or "").strip()
            if conf in {"high", "medium"} and brands:
                return clean_brands(company, brands)
            return NPD
        except Exception as e:  # noqa: BLE001
            print(f"    deepseek error (attempt {attempt + 1}): {e}", flush=True)
            time.sleep(2.0 * (attempt + 1))
    return NPD


def main() -> int:
    return asyncio.run(main_async())


async def main_async() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="Process only first N NPD rows (smoke test)")
    args = ap.parse_args()

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    if not api_key:
        print("ERROR: OPENAI_API_KEY missing in .env", file=sys.stderr)
        return 1
    client = OpenAI(api_key=api_key, base_url=base, timeout=75.0)
    smoke = client.chat.completions.create(model=MODEL, messages=[{"role": "user", "content": "OK"}], max_tokens=3)
    print(f"DeepSeek OK: {smoke.choices[0].message.content!r} model={smoke.model}", flush=True)

    gai_ok = False
    try:
        gai_ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        gai_ok = False
    print(f"Google AI scraper @ {GAI_BASE_URL}: {'UP' if gai_ok else 'DOWN (site scrape + DDGS only)'}", flush=True)
    print(f"DDGS backends: {BACKEND_NAMES}", flush=True)

    if not XLSX.exists():
        print(f"ERROR: not found: {XLSX}", file=sys.stderr)
        return 1

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = XLSX.with_name(XLSX.stem + f"_BEFORE_NPD_BRANDS_DEEP_{stamp}.xlsx")
    if not bak.exists():
        bak.write_bytes(XLSX.read_bytes())
        print(f"backup -> {bak.name}", flush=True)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    targets: list[tuple[int, str]] = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        if emptyish(ws.cell(r, ic["Key Brands Represented"]).value):
            targets.append((r, co))
    print(f"NPD brand targets: {len(targets)}", flush=True)

    ckpt: dict = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass
    done = ckpt.setdefault("done", {})

    todo = [(r, co) for r, co in targets if co not in done]
    if args.limit:
        todo = todo[: args.limit]
    print(f"to process this run: {len(todo)} (already done: {len(targets) - len(todo)})", flush=True)

    filled = 0
    for i, (r, co) in enumerate(todo, 1):
        website = str(ws.cell(r, ic["Website"]).value or "").strip()
        print(f"[{i}/{len(todo)}] {co}", flush=True)

        site_ev = scrape_site(website)
        gai_ev = await gather_gai_evidence(co, website, gai_ok)
        ddgs_ev = gather_ddgs_evidence(co, website)
        evidence = "\n\n".join(p for p in (gai_ev, site_ev, ddgs_ev) if p)[:12000]

        brands = deepseek_brands(client, co, evidence) if evidence else NPD
        time.sleep(0.3)

        applied = brands != NPD
        if applied:
            ws.cell(r, ic["Key Brands Represented"]).value = brands
            filled += 1

        done[co] = {"applied": applied, "brands": brands}
        print(f"    {'FILLED' if applied else 'still NPD'} -> {brands[:100]}", flush=True)

        if i % 10 == 0:
            wb.save(XLSX)
            CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
            print(f"  checkpoint saved at {i}/{len(todo)}", flush=True)

    wb.save(XLSX)
    CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    still_npd = sum(1 for r, co in targets if done.get(co, {}).get("brands", NPD) == NPD)
    audit = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "targets": len(targets),
        "processed_this_run": len(todo),
        "filled_this_run": filled,
        "still_npd_total": still_npd,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(audit, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
