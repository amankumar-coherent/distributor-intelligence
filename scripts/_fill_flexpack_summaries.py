# -*- coding: utf-8 -*-
"""Rewrite Summary column for Global Flexible Packaging FINAL via DeepSeek.

Rules:
- Use ONLY facts present in the row (and optional scraped About/home text).
- Never invent founded year, employees, brands, phones, or claims.
- 2–4 sentences; channel/distributor focus; plain English.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import httpx
import openpyxl
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

OUT = ROOT / "output/web_expand/global_flexible_packaging_market_global"
XLSX = OUT / "global_flexible_packaging_market_global_FINAL.xlsx"
CACHE = OUT / "flexpack_summary_cache.json"
AUDIT = OUT / "flexpack_summary_audit.json"
SUMMARY = OUT / "flexpack_summary_fill_stats.json"

NPD_SET = {
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
    "not applicable",
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

FACT_COLS = [
    "Company",
    "Website",
    "Founded",
    "Headquarters",
    "Continent / Geography",
    "Operational Presence",
    "Ownership",
    "Employees",
    "Core Categories",
    "Specialty Focus",
    "Key Brands Represented",
    "Online / Offline / Both",
    "Contact Person",
    "Role",
    "Office No.",
    "Office Email",
    "Country Code",
    "Region Code",
]


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in NPD_SET


def norm(v: object) -> str:
    s = str(v or "").strip()
    if emptyish(s):
        return ""
    # fix mojibake dash
    s = s.replace("\ufffd", "—").replace("�", "—")
    return s


def is_weak_summary(company: str, summary: str) -> bool:
    s = (summary or "").strip()
    if emptyish(s):
        return True
    low = s.lower()
    if len(s) < 90:
        return True
    markers = (
        "verified packaging channel",
        "live website",
        "channel/distributor verified",
        "consolidated from",
        "operates as a packaging distribution / channel partner supplying flexible packaging-related products (films, bags, pouches, protective packaging)",
        "todo",
    )
    if any(m in low for m in markers):
        return True
    # ultra-generic stub: "Name — US packaging distributor."
    if re.fullmatch(
        rf"{re.escape(company)}\s*[—\-–]\s*.{{10,70}}\.?",
        s,
        flags=re.I,
    ):
        return True
    return False


def row_facts(row: dict[str, str]) -> dict[str, str]:
    out = {}
    for c in FACT_COLS:
        v = norm(row.get(c))
        if v:
            out[c] = v
    return out


def facts_blob(facts: dict[str, str]) -> str:
    lines = [f"{k}: {v}" for k, v in facts.items()]
    return "\n".join(lines)


def scrape_about_snippet(website: str) -> str:
    url = (website or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    paths = ["", "/about", "/about-us", "/aboutus", "/company", "/en/about"]
    texts: list[str] = []
    try:
        with httpx.Client(
            headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"},
            verify=False,
            timeout=httpx.Timeout(10.0, connect=4.0),
            follow_redirects=True,
        ) as client:
            for path in paths:
                try:
                    u = urljoin(url.rstrip("/") + "/", path.lstrip("/")) if path else url
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    html = r.text or ""
                    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
                    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
                    t = re.sub(r"(?is)<[^>]+>", " ", t)
                    t = re.sub(r"&\w+;", " ", t)
                    t = re.sub(r"\s+", " ", t).strip()
                    if len(t) > 200:
                        texts.append(t[:2500])
                    if texts and path:  # homepage + one about is enough
                        break
                except Exception:
                    continue
    except Exception:
        return ""
    return "\n---\n".join(texts)[:4000]


def ds_client() -> tuple[OpenAI, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        raise SystemExit("OPENAI_API_KEY missing in D:\\vendor-intel\\.env")
    return OpenAI(api_key=key, base_url=base, timeout=90.0), model


def _parse_json(text: str) -> dict:
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


def invent_check(summary: str, facts: dict[str, str], site_text: str) -> list[str]:
    """Light hallucination guards — flag invented founded years / employee bands not in facts."""
    issues: list[str] = []
    blob = (facts_blob(facts) + "\n" + (site_text or "")).lower()
    # founded year in summary must appear in facts
    for y in re.findall(r"\b(19\d{2}|20[0-2]\d)\b", summary):
        if y not in blob and y not in (facts.get("Founded") or ""):
            issues.append(f"year:{y}")
    # reject obvious filler
    low = summary.lower()
    if "leading global leader" in low or "world's largest" in low and "largest" not in blob:
        if "world's largest" in low or "worlds largest" in low:
            issues.append("superlative")
    return issues


def deepseek_summary(
    client: OpenAI,
    model: str,
    company: str,
    facts: dict[str, str],
    site_text: str,
) -> dict:
    system = (
        "You write factual company summaries for a B2B packaging channel landscape database. "
        "Use ONLY the provided FACTS and optional WEBSITE_TEXT. "
        "Do NOT invent founded years, employee counts, brands, geographies, phones, emails, "
        "or claims not explicitly present. "
        "If a field is missing, omit it — do not guess. "
        "Write 2 to 4 clear sentences in English. "
        "Focus on: who they are (distributor/wholesaler/merchant/channel), what packaging they supply "
        "(from Core Categories / Specialty Focus), where they operate (from HQ / geography / presence), "
        "and channel mode (Online/Offline/Both) if present. "
        "Mention Key Brands Represented only if named brands are listed (not 'Not publicly disclosed'). "
        "No marketing fluff. No bullet points. Return ONLY JSON."
    )
    user = (
        f"Company: {company}\n\n"
        f"FACTS:\n{facts_blob(facts)}\n\n"
        f"WEBSITE_TEXT (optional evidence; may be empty):\n{(site_text or '')[:3500]}\n\n"
        "Return JSON:\n"
        '{"ok":true|false,"summary":"2-4 sentences","used_fields":["..."],'
        '"confidence":"high|medium|low"}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0.2,
            max_tokens=450,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception as ex:  # noqa: BLE001
        return {"ok": False, "summary": "", "error": str(ex)[:160]}

    summary = str(data.get("summary") or "").strip()
    summary = re.sub(r"\s+", " ", summary).strip()
    # normalize dashes
    summary = summary.replace(" – ", " — ").replace(" - ", " — ")
    conf = str(data.get("confidence") or "").lower()
    ok = bool(data.get("ok")) and bool(summary) and conf in {"high", "medium"}
    if not summary or len(summary) < 60:
        return {"ok": False, "summary": "", "raw": data}
    issues = invent_check(summary, facts, site_text)
    if issues:
        # drop invented years by regenerating strip — safer to reject
        return {"ok": False, "summary": "", "issues": issues, "raw": data, "draft": summary}
    return {
        "ok": ok or (len(summary) >= 80 and conf in {"high", "medium", "low"}),
        "summary": summary,
        "confidence": conf,
        "used_fields": data.get("used_fields") or [],
    }


def fallback_summary(facts: dict[str, str]) -> str:
    """Deterministic non-hallucinated summary from row fields only."""
    name = facts.get("Company") or "This company"
    bits = []
    hq = facts.get("Headquarters")
    geo = facts.get("Continent / Geography")
    core = facts.get("Core Categories")
    spec = facts.get("Specialty Focus")
    mode = facts.get("Online / Offline / Both")
    founded = facts.get("Founded")
    own = facts.get("Ownership")
    brands = facts.get("Key Brands Represented")
    presence = facts.get("Operational Presence")

    loc = hq or geo or ""
    s1 = f"{name} is a packaging channel/distributor"
    if loc:
        s1 += f" based in {loc}"
    s1 += "."
    bits.append(s1)

    if core or spec:
        focus = spec or core
        bits.append(f"It supplies {focus}.")
    if presence and presence != (hq or ""):
        bits.append(f"Operational presence includes {presence}.")
    extras = []
    if founded:
        extras.append(f"founded {founded}")
    if own:
        extras.append(f"{own.lower()} ownership" if "own" not in own.lower() else own)
    if mode:
        extras.append(f"selling {mode.lower()}")
    if extras:
        bits.append("Profile notes: " + "; ".join(extras) + ".")
    if brands and brands.lower() != "not publicly disclosed":
        bits.append(f"Key brands represented include {brands}.")
    text = " ".join(bits)
    return re.sub(r"\s+", " ", text).strip()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--all", action="store_true", help="Rewrite all rows, not only weak/empty")
    ap.add_argument("--no-scrape", action="store_true")
    ap.add_argument("--workers", type=int, default=8, help="About-page scrape workers")
    ap.add_argument("--resume", action="store_true")
    args = ap.parse_args()

    client, model = ds_client()
    print(f"DeepSeek model={model}", flush=True)

    cache: dict = {}
    if args.resume and CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_SUMMARY_{stamp}.xlsx")
    wb0 = openpyxl.load_workbook(XLSX)
    wb0.save(backup)
    wb0.close()
    print(f"backup {backup.name}", flush=True)

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}
    if "Summary" not in ic:
        raise SystemExit("Summary column missing")

    todos = []
    for r in range(2, ws.max_row + 1):
        row = {h: ws.cell(r, ic[h]).value for h in ic}
        company = norm(row.get("Company"))
        if not company:
            continue
        cur = norm(row.get("Summary"))
        if args.all or is_weak_summary(company, cur):
            todos.append({"row": r, "company": company, "data": {h: norm(row.get(h)) for h in ic}})

    if args.limit:
        todos = todos[: args.limit]
    print(f"todo summaries={len(todos)} / rows={ws.max_row - 1}", flush=True)

    # scrape about snippets in parallel
    site_map: dict[str, str] = {}
    if not args.no_scrape:
        webs = []
        seen = set()
        for t in todos:
            w = t["data"].get("Website") or ""
            if w and w not in seen:
                seen.add(w)
                if args.resume and w in cache and cache[w].get("site_text") is not None:
                    site_map[w] = cache[w].get("site_text") or ""
                else:
                    webs.append(w)
        print(f"scraping about snippets for {len(webs)} sites ...", flush=True)
        with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
            futs = {ex.submit(scrape_about_snippet, w): w for w in webs}
            done = 0
            for fut in as_completed(futs):
                w = futs[fut]
                try:
                    site_map[w] = fut.result() or ""
                except Exception:
                    site_map[w] = ""
                done += 1
                if done % 40 == 0 or done == len(webs):
                    print(f"  scrape {done}/{len(webs)}", flush=True)

    audit = []
    written = 0
    fallback_n = 0
    failed = 0

    for i, t in enumerate(todos, 1):
        company = t["company"]
        facts = row_facts(t["data"])
        web = facts.get("Website") or ""
        site_text = site_map.get(web) or ""
        cache_key = company

        result = None
        if args.resume and cache_key in cache and cache[cache_key].get("summary"):
            result = cache[cache_key]
        else:
            result = deepseek_summary(client, model, company, facts, site_text)
            # brief pacing
            if i % 15 == 0:
                time.sleep(0.35)

        summary = ""
        source = "deepseek"
        if result and result.get("ok") and result.get("summary"):
            summary = result["summary"]
        else:
            summary = fallback_summary(facts)
            source = "fallback"
            fallback_n += 1
            if result and not result.get("ok"):
                failed += 1

        # final sanitize
        summary = re.sub(r"\s+", " ", summary).strip()
        if len(summary) > 600:
            # keep first ~4 sentences
            parts = re.split(r"(?<=[.!?])\s+", summary)
            summary = " ".join(parts[:4]).strip()

        ws.cell(t["row"], ic["Summary"]).value = summary
        ws.cell(t["row"], ic["Summary"]).alignment = Alignment(wrap_text=True, vertical="top")
        written += 1

        entry = {
            "company": company,
            "source": source,
            "summary": summary,
            "confidence": (result or {}).get("confidence"),
            "issues": (result or {}).get("issues"),
            "site_text_len": len(site_text),
        }
        audit.append(entry)
        cache[cache_key] = {
            "summary": summary,
            "ok": source == "deepseek",
            "site_text": site_text[:500] if site_text else "",
            "confidence": entry.get("confidence"),
        }

        if i % 25 == 0 or i == len(todos):
            print(
                f"  apply {i}/{len(todos)} written={written} fallback={fallback_n}",
                flush=True,
            )
            CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
            # intermediate save every 50
            if i % 50 == 0:
                wb.save(XLSX)

    ws.column_dimensions[get_column_letter(ic["Summary"])].width = 56
    wb.save(XLSX)
    wb.close()

    # recount quality
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    si = headers.index("Summary")
    ni = headers.index("Company")
    weak = empty = ok = 0
    samples = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[ni] or "")
        s = str(row[si] or "").strip()
        if emptyish(s):
            empty += 1
        elif is_weak_summary(name, s):
            weak += 1
        else:
            ok += 1
            if len(samples) < 8:
                samples.append({"company": name, "summary": s[:220]})
    wb.close()

    stats = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "file": str(XLSX),
        "backup": str(backup),
        "todos": len(todos),
        "written": written,
        "deepseek_ok_est": written - fallback_n,
        "fallback": fallback_n,
        "failed_ds": failed,
        "quality_after": {"ok": ok, "weak": weak, "empty": empty},
        "samples": samples,
    }
    SUMMARY.write_text(json.dumps(stats, indent=2, ensure_ascii=False), encoding="utf-8")
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(stats, indent=2, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    try:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    except Exception:
        pass
    main()
