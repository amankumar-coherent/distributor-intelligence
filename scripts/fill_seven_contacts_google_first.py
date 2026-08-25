#!/usr/bin/env python3
"""Fill Contact Person / Role / LinkedIn on 7 market Excels.

One approach at a time per company:
  1) Apollo only (sales / procurement). If a person is found, stop person-search.
  2) If Apollo has no result → Google once for name + role.
  3) Then LinkedIn search once for that person's linkedin.com/in/ URL.

Never invent emails/phones. First name + partial last name is OK.

Usage:
  .venv\\Scripts\\python.exe scripts\\fill_seven_contacts_google_first.py
  .venv\\Scripts\\python.exe scripts\\fill_seven_contacts_google_first.py --market electrical --limit 20
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import dotenv_values, load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font

load_dotenv(ROOT / ".env")
_cmi = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env")
if _cmi.exists():
    for k, v in (dotenv_values(_cmi) or {}).items():
        if v and (
            k.startswith(("OPENAI_", "APOLLO_", "LLM_", "GOOGLE_AI_"))
            or k in {"COMPILER_MODEL", "CLASSIFIER_MODEL", "MARKET_MAP_MODEL"}
        ):
            os.environ[k] = v

MARKETS: list[dict[str, Any]] = [
    {
        "key": "electrical",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_electrical_distribution_market_global"
            r"\global_electrical_distribution_market_global_FINAL.xlsx"
        ),
    },
    {
        "key": "adc",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_antibody_drug_conjugates_market_global"
            r"\global_antibody_drug_conjugates_market.xlsx"
        ),
    },
    {
        "key": "avocado",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global"
            r"\global_avocado_oil_market.xlsx"
        ),
    },
    {
        "key": "smartphone",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_smartphone_market_global"
            r"\global_smartphone_market.xlsx"
        ),
    },
    {
        "key": "firewall",
        "path": Path(
            r"D:\vendor-intel\output\run4\global_firewall_market_global"
            r"\global_firewall_market.xlsx"
        ),
    },
    {
        "key": "green",
        "path": Path(
            r"D:\vendor-intel\output\run4\global_green_chemical_market_global"
            r"\global_green_chemical_market.xlsx"
        ),
    },
    {
        "key": "smartwatch",
        "path": Path(
            r"D:\vendor-intel\output\run4\global_smartwatch_market_global"
            r"\global_smartwatch_market.xlsx"
        ),
    },
]

NPD = {
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
}

# Prefer these roles — DeepSeek picks ONE query style per company
ROLE_QUERY_TEMPLATES = [
    "{company} head of sales name role and his linkedin",
    "{company} VP of sales name role and linkedin",
    "{company} sales director name role linkedin",
    "{company} commercial director name role linkedin",
    "{company} head of procurement name role linkedin",
    "{company} procurement manager name role linkedin",
    "{company} VP procurement name role linkedin",
    "{company} chief revenue officer name role linkedin",
]

# Must be sales / procurement (or close commercial / CRO / BD). Not CEO/HR/marketing-only.
OK_ROLE = re.compile(
    r"sales|procurement|purchasing|"
    r"\bcro\b|chief revenue|revenue officer|"
    r"commercial|"
    r"business development|\bbd\b|"
    r"channel sales|partner sales",
    re.I,
)
BAD_EXEC = re.compile(
    r"\b(ceo|cfo|cto|coo|ciso|founder|co-founder|chairman|owner)\b|"
    r"managing director|"
    r"(?<!vice )(?<!vice\s)\bpresident\b",
    re.I,
)
BAD_ROLE = BAD_EXEC  # alias for older call sites
LI_IN = re.compile(r"https?://(?:www\.)?linkedin\.com/in/[A-Za-z0-9\-_%]+/?", re.I)


def _is_allowed_role(role: str) -> bool:
    """Only Head of Sales / VP Sales / Procurement Manager and similar."""
    r = re.sub(r"\s+", " ", (role or "").strip())
    if not r:
        return False
    if not OK_ROLE.search(r):
        return False
    rl = r.lower()
    salesish = bool(
        re.search(r"sales|procurement|purchasing|commercial|revenue|\bcro\b", rl)
    )
    if BAD_EXEC.search(r) and not salesish:
        return False
    if re.search(r"\bpresident\b", rl) and "vice president" not in rl and not salesish:
        return False
    return True


def _empty(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


def _name_ok(name: str) -> bool:
    """First name is enough; partial/obfuscated last name (Mo***s) is OK."""
    s = re.sub(r"\s+", " ", (name or "").strip())
    if not s or s.lower() in NPD:
        return False
    first = s.split()[0].strip(".,")
    letters = re.sub(r"[^A-Za-z]", "", first)
    return len(letters) >= 2


def _needs_contact(row: dict[str, str]) -> bool:
    """True if Contact Person / Role still need a sales|procurement fill.

    First name + Role is enough. Partial last name with * is accepted.
    Do not re-search just to chase LinkedIn.
    """
    cp = str(row.get("Contact Person") or "").strip()
    if not _name_ok(cp):
        return True
    role = str(row.get("Role") or "").strip()
    if not _is_allowed_role(role):
        return True
    return False


def _find_header(ws) -> tuple[list[str], int]:
    for r in range(1, 6):
        vals = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=r, max_row=r))]
        if "Company" in vals and "Contact Person" in vals:
            return vals, r
    raise RuntimeError("No Company/Contact Person header")


def load_xlsx(path: Path) -> tuple[list[str], int, list[dict[str, Any]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb["Companies"] if "Companies" in wb.sheetnames else wb.active
    headers, hr = _find_header(ws)
    rows: list[dict[str, Any]] = []
    for ridx in range(hr + 1, ws.max_row + 1):
        d: dict[str, Any] = {"_row": ridx}
        for j, h in enumerate(headers):
            if not h:
                continue
            d[h] = "" if ws.cell(ridx, j + 1).value is None else str(ws.cell(ridx, j + 1).value).strip()
        if d.get("Company"):
            rows.append(d)
    wb.close()
    return headers, hr, rows


def write_contacts(path: Path, headers: list[str], hr: int, rows: list[dict[str, Any]]) -> None:
    """Write contact columns atomically (temp file + replace) to avoid BadZipFile corruption."""
    wb = load_workbook(path)
    ws = wb["Companies"] if "Companies" in wb.sheetnames else wb.active
    col = {h: i + 1 for i, h in enumerate(headers) if h}
    for r in rows:
        ridx = int(r["_row"])
        for field in ("Contact Person", "Role", "LinkedIn"):
            if field in col and field in r:
                cell = ws.cell(ridx, col[field])
                val = r.get(field) or ""
                cell.value = val
                if field == "LinkedIn" and isinstance(val, str) and "linkedin.com/in/" in val.lower():
                    cell.hyperlink = val
                    cell.font = Font(color="0563C1", underline="single")
    for cell in ws[hr]:
        cell.font = Font(bold=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        wb.save(tmp)
    finally:
        wb.close()
    # Replace only after a valid xlsx is on disk
    tmp.replace(path)


def _ds_client():
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key or "deepseek" not in base.lower():
        raise RuntimeError("DeepSeek OPENAI_* not configured")
    return OpenAI(api_key=key, base_url=base), model


def _parse_json(text: str) -> Any:
    raw = (text or "").strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.I)
    if m:
        raw = m.group(1).strip()
    i = raw.find("{")
    if i < 0:
        return {}
    raw = raw[i:]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        j = raw.rfind("}")
        if j > 0:
            try:
                return json.loads(raw[: j + 1])
            except json.JSONDecodeError:
                return {}
        return {}


def deepseek_pick_query(company: str, website: str, hq: str) -> str:
    """Ask DeepSeek for ONE Google query in the image style."""
    client, model = _ds_client()
    system = (
        "You generate ONE Google search query to find a company's commercial contact. "
        "Prefer Head of Sales, else VP of Sales / Sales Director / Commercial Director, "
        "else Head of Procurement / Procurement Manager / VP Procurement. "
        "Query style MUST be like: "
        "'{Company} head of sales name role and his linkedin' "
        "or '{Company} VP of sales name role and linkedin' "
        "or '{Company} procurement manager name role linkedin'. "
        "Return ONLY JSON: {\"query\":\"...\",\"target_role\":\"Head of Sales|VP of Sales|...\"}"
    )
    user = f"Company: {company}\nWebsite: {website}\nHQ: {hq}\n"
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=200,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
        q = str(data.get("query") or "").strip()
        if q and "linkedin" in q.lower() and len(q) > 20:
            return q
    except Exception:
        pass
    # Fallback template
    return ROLE_QUERY_TEMPLATES[0].format(company=company)


def deepseek_pick_one_contact(company: str, overview: str, query: str) -> dict[str, str] | None:
    """From Google AI Overview text, pick exactly ONE sales/procurement person."""
    if not (overview or "").strip():
        return None
    client, model = _ds_client()
    system = (
        "From Google AI Overview / SERP text, pick EXACTLY ONE real person whose "
        "job is SALES or PROCUREMENT only: Head of Sales, VP of Sales, Sales Director, "
        "Sales Manager, Commercial Director, CRO, Head of Procurement, Procurement Manager, "
        "Purchasing Manager. "
        "If multiple people are listed, pick ONE (prefer Head of Sales, then VP Sales, "
        "then Sales Director, then Procurement). "
        "Do NOT pick CEO, CFO, CTO, Founder, Chairman, Owner, HR, Marketing, Operations, "
        "or Project Director unless the title also clearly includes Sales or Procurement. "
        "Return ONLY JSON: "
        '{"found":true|false,"name":"Full Name","role":"Job Title",'
        '"linkedin":"https://www.linkedin.com/in/... or empty","confidence":"high|medium|low"} '
        "Never invent LinkedIn URLs — only copy if present in the text. "
        "If unsure, found=false."
    )
    user = (
        f"Company: {company}\nSearch query: {query}\n\nGOOGLE TEXT:\n{overview[:7000]}\n"
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
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception:
        return None
    if not data.get("found"):
        return None
    if str(data.get("confidence") or "").lower() == "low":
        return None
    name = re.sub(r"\s+", " ", str(data.get("name") or "").replace("\n", " ")).strip()
    role = re.sub(r"\s+", " ", str(data.get("role") or "").replace("\n", " ")).strip()
    li = str(data.get("linkedin") or "").strip()
    if not _name_ok(name):
        return None
    if not _is_allowed_role(role):
        return None
    # Prefer LinkedIn from DeepSeek; else scrape from overview text
    if not (li and "linkedin.com/in/" in li.lower()):
        m = LI_IN.search(overview or "")
        # Prefer URL near the chosen name if possible
        li = ""
        for m in LI_IN.finditer(overview or ""):
            window = overview[max(0, m.start() - 120) : m.end() + 40]
            if name.split()[0].lower() in window.lower() or name.split()[-1].lower() in window.lower():
                li = m.group(0)
                break
        if not li:
            m0 = LI_IN.search(overview or "")
            if m0:
                li = m0.group(0)
    out = {"contact_person": name, "contact_role": role, "source": "google_ai+deepseek"}
    if li and "linkedin.com/in/" in li.lower() and "/company/" not in li.lower():
        out["linkedin"] = li.split("?")[0].rstrip("/")
    return out


def _domain_from_website(website: str) -> str:
    s = (website or "").strip().lower()
    if not s:
        return ""
    if "://" not in s:
        s = "https://" + s
    try:
        host = urlparse(s).netloc or ""
    except Exception:
        return ""
    return host.removeprefix("www.").split("/")[0]


async def google_ai_search(query: str, sem: asyncio.Semaphore) -> str:
    """One Google AI Overview scrape (concurrency capped by sem). Empty on failure."""
    from vendor_intel.clients.google_ai_scraper import google_ai_ask, heuristic_clean_scrape_markdown

    async with sem:
        print(f"    google: {query[:90]}", flush=True)
        try:
            data = await google_ai_ask(query, llm_clean=False)
        except Exception as e:
            print(f"    google error: {type(e).__name__}: {e}", flush=True)
            return ""
        md = str(data.get("markdown") or "").strip()
        raw = str(data.get("markdown_raw") or "").strip()
        if not md and raw:
            md = heuristic_clean_scrape_markdown(raw)
        return md


def deepseek_direct_contact(company: str, website: str, hq: str) -> dict[str, str] | None:
    """Fallback when Google AI Overview is empty: DeepSeek knowledge (no invented LinkedIn)."""
    client, model = _ds_client()
    system = (
        "You identify ONE real SALES or PROCUREMENT leader for a company. "
        "Prefer: Head of Sales, VP of Sales, Sales Director, Sales Manager, "
        "Commercial Director, CRO, Head of Procurement, Procurement Manager. "
        "Do NOT pick CEO/CFO/Founder/Chairman/HR/Marketing/Operations "
        "unless the title also includes Sales or Procurement. "
        "Only return people you are confident are real for THIS company. "
        "Never invent LinkedIn URLs — leave linkedin empty unless you know a real "
        "https://www.linkedin.com/in/... profile. "
        "Return ONLY JSON: "
        '{"found":true|false,"name":"Full Name","role":"Job Title",'
        '"linkedin":"","confidence":"high|medium|low","reason":"brief"}'
    )
    user = (
        f"Company: {company}\nWebsite: {website or 'unknown'}\nHQ: {hq or 'unknown'}\n"
        "Find one Head of Sales / VP Sales / Procurement Manager (or similar)."
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=350,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception:
        return None
    if not data.get("found"):
        return None
    if str(data.get("confidence") or "").lower() != "high":
        return None
    name = re.sub(r"\s+", " ", str(data.get("name") or "").replace("\n", " ")).strip()
    role = re.sub(r"\s+", " ", str(data.get("role") or "").replace("\n", " ")).strip()
    li = str(data.get("linkedin") or "").strip()
    if not _name_ok(name):
        return None
    if not _is_allowed_role(role):
        return None
    out = {
        "contact_person": name,
        "contact_role": role,
        "source": "deepseek_direct",
    }
    if li and "linkedin.com/in/" in li.lower() and "/company/" not in li.lower():
        out["linkedin"] = li.split("?")[0].rstrip("/")
    return out


async def apollo_fallback(company: str, website: str) -> dict[str, str] | None:
    from vendor_intel.enrichment.apollo_contacts import ApolloContactClient

    if not (os.getenv("APOLLO_API_KEY") or "").strip():
        return None
    # Fail fast: long Apollo 429 cool-downs stall whole waves.
    delay = float(os.getenv("APOLLO_REQUEST_DELAY_SEC") or "5.0")
    os.environ.setdefault("APOLLO_MAX_RETRY_WAIT_SEC", "20")
    try:
        async with ApolloContactClient(verbose=False, request_delay=delay, timeout=30.0) as apollo:
            hit = await asyncio.wait_for(
                apollo.find_contact(
                    {
                        "company": company,
                        "website": website,
                        "domain": _domain_from_website(website),
                    },
                    enrich_match=False,
                ),
                timeout=45,
            )
    except Exception:
        return None
    if not hit or hit.get("error") or not hit.get("contact_person"):
        return None
    name = str(hit.get("contact_person") or "").strip()
    role = str(hit.get("contact_role") or "").strip()
    li = str(hit.get("linkedin") or "").strip()
    if not _name_ok(name):
        return None
    if not _is_allowed_role(role):
        return None
    out = {
        "contact_person": name,
        "contact_role": role or "Sales",
        "source": "apollo",
    }
    if li and "linkedin.com/in/" in li.lower():
        out["linkedin"] = li
    return out


def _paths(xlsx: Path) -> tuple[Path, Path, Path]:
    d = xlsx.parent
    stem = xlsx.stem
    return (
        d / f"{stem}_gfirst_contact_cache.json",
        d / f"{stem}_gfirst_contact_progress.jsonl",
        d / f"{stem}_gfirst_contact_summary.json",
    )


async def _google_linkedin(name: str, company: str, sem: asyncio.Semaphore) -> str:
    first = (name or "").split()[0]
    if not first:
        return ""
    ov = await google_ai_search(f"{first} {company} linkedin", sem)
    if not ov:
        return ""
    m = LI_IN.search(ov)
    if m and "/company/" not in m.group(0).lower():
        return m.group(0).split("?")[0].rstrip("/")
    return ""


async def fill_one(
    r: dict[str, Any],
    *,
    sem: asyncio.Semaphore,
    use_apollo: bool,
    use_google: bool,
) -> dict[str, Any]:
    company = r["Company"]
    website = str(r.get("Website") or "").strip()
    hq = str(r.get("Headquarters") or "").strip()
    query = ""
    hit = None
    source = ""

    # 1) Apollo ONLY — if it returns a sales/procurement person, do not Google for the person
    if use_apollo:
        print(f"    1 apollo: {company[:40]}", flush=True)
        hit = await apollo_fallback(company, website)
        if hit and hit.get("contact_person"):
            query = f"apollo:{company}"
            source = "apollo"

    # 2) Google ONLY if Apollo had no person
    if (not hit or not hit.get("contact_person")) and use_google:
        print(f"    2 google: {company[:40]}", flush=True)
        query = await asyncio.to_thread(deepseek_pick_query, company, website, hq)
        overview = await google_ai_search(query, sem)
        if overview:
            hit = await asyncio.to_thread(deepseek_pick_one_contact, company, overview, query)
            if hit and hit.get("contact_person"):
                source = "google_ai+deepseek"

    if not hit or not hit.get("contact_person"):
        return {"status": "no_match", "company": company, "query": query}

    if not _is_allowed_role(str(hit.get("contact_role") or "")):
        return {"status": "no_match", "company": company, "query": query}

    # 3) LinkedIn search ONLY after we already have the person (one extra Google query)
    if use_google and not hit.get("linkedin"):
        print(f"    3 linkedin: {hit['contact_person'].split()[0]} {company[:28]}", flush=True)
        li = await _google_linkedin(hit["contact_person"], company, sem)
        if li:
            hit["linkedin"] = li

    return {
        "status": "filled",
        "company": company,
        "query": query,
        "contact_person": hit["contact_person"],
        "role": hit.get("contact_role") or "",
        "linkedin": hit.get("linkedin") or "Not publicly disclosed",
        "source": source or hit.get("source") or "google_ai",
    }


async def fill_market(
    *,
    key: str,
    path: Path,
    limit: int | None,
    resume: bool,
    use_apollo: bool,
    concurrency: int,
) -> dict[str, Any]:
    from vendor_intel.clients.google_ai_scraper import google_ai_ping

    if not path.exists():
        raise FileNotFoundError(path)

    use_google = False
    try:
        use_google = bool(await google_ai_ping())
    except Exception:
        use_google = False
    if use_google:
        print(f"  Google AI scraper: ON", flush=True)
    else:
        print(
            "  Google AI scraper: OFF (unreachable) — DeepSeek fallback for all rows",
            flush=True,
        )

    bak = path.with_name(path.stem + "_BEFORE_GFIRST_CONTACTS.xlsx")
    if not bak.exists():
        bak.write_bytes(path.read_bytes())
        print(f"  backup {bak.name}", flush=True)

    headers, hr, rows = load_xlsx(path)
    cache_path, progress_path, summary_path = _paths(path)
    cache: dict[str, Any] = {}
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    done: set[str] = set()
    if resume and progress_path.exists():
        for line in progress_path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            co = str(rec.get("company") or "").strip().lower()
            st = str(rec.get("status") or "").lower()
            # Only skip successful fills. Retry prior no_match so DeepSeek
            # fallback can run when Google returned nothing.
            if co and st == "filled":
                done.add(co)

    todo = [
        r
        for r in rows
        if r["Company"].strip().lower() not in done and _needs_contact(r)
    ]
    if limit is not None:
        todo = todo[:limit]

    print(
        f"\n=== {key}: {len(todo)} todo / {len(rows)} rows "
        f"(concurrency={concurrency}, google={use_google}, apollo={use_apollo}) ===",
        flush=True,
    )

    conc = max(1, min(15, concurrency))  # up to 15 companies at a time
    sem = asyncio.Semaphore(conc)
    filled = 0
    no_match = 0
    errors = 0

    def _append(rec: dict) -> None:
        with progress_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    # Process in waves of up to 15 companies
    for wave_start in range(0, len(todo), conc):
        wave = todo[wave_start : wave_start + conc]
        print(
            f"  wave {wave_start // conc + 1}: "
            + ", ".join(x["Company"][:28] for x in wave),
            flush=True,
        )
        tasks = [
            fill_one(r, sem=sem, use_apollo=use_apollo, use_google=use_google)
            for r in wave
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)

        for r, res in zip(wave, results):
            co = r["Company"]
            if isinstance(res, Exception):
                errors += 1
                print(f"  -> ERROR {co[:40]}: {res}", flush=True)
                _append({"company": co, "status": "error", "error": str(res)[:200]})
                continue
            assert isinstance(res, dict)
            if res.get("status") != "filled":
                no_match += 1
                print(f"  -> no match {co[:40]}", flush=True)
                _append(res)
                cache[co.strip().lower()] = {"skipped": True}
                continue

            person = str(res["contact_person"]).strip()
            role = str(res.get("role") or "").strip()
            li = str(res.get("linkedin") or "").strip()
            if not (li and "linkedin.com/in/" in li.lower()):
                li = "Not publicly disclosed"

            r["Contact Person"] = person
            r["Role"] = role
            r["LinkedIn"] = li
            filled += 1
            line = (
                f"  -> {person} | {role}"
                + (" | LI" if "linkedin.com/in/" in li.lower() else "")
                + f" [{res.get('source')}]"
            )
            try:
                print(line, flush=True)
            except UnicodeEncodeError:
                print(line.encode("ascii", "replace").decode("ascii"), flush=True)
            _append(res)
            cache[co.strip().lower()] = {
                "contact_person": person,
                "contact_role": role,
                "linkedin": li,
                "source": res.get("source"),
                "query": res.get("query"),
            }

        cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
        write_contacts(path, headers, hr, rows)
        print(f"  … checkpoint filled={filled} no_match={no_match}", flush=True)

    still_empty = sum(1 for r in rows if _empty(r.get("Contact Person")))
    summary = {
        "when": datetime.now(timezone.utc).isoformat(),
        "market": key,
        "path": str(path),
        "todo": len(todo),
        "filled": filled,
        "no_match": no_match,
        "errors": errors,
        "still_empty_contact": still_empty,
        "total_rows": len(rows),
        "google_concurrency": conc,
        "google_enabled": use_google,
    }
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(
        f"=== {key} done: filled={filled} no_match={no_match} "
        f"errors={errors} still_empty={still_empty} ===",
        flush=True,
    )
    return summary


async def main_async() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--market", default="", help="Comma keys e.g. electrical,adc")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--no-resume", action="store_true")
    ap.add_argument("--no-apollo", action="store_true")
    ap.add_argument(
        "--concurrency",
        type=int,
        default=15,
        help="Companies at a time (hard-capped at 15)",
    )
    args = ap.parse_args()

    if not (os.getenv("OPENAI_API_KEY") or "").strip():
        raise SystemExit("DeepSeek OPENAI_API_KEY missing")

    keys = [k.strip() for k in args.market.split(",") if k.strip()] if args.market else [
        m["key"] for m in MARKETS
    ]
    selected = [m for m in MARKETS if m["key"] in keys]
    if not selected:
        raise SystemExit(f"No markets matched {keys}")

    conc = max(1, min(15, int(args.concurrency or 15)))
    limit = args.limit or None
    all_sum: list[dict] = []
    # One market at a time (sequential loop)
    for m in selected:
        s = await fill_market(
            key=m["key"],
            path=m["path"],
            limit=limit,
            resume=not args.no_resume,
            use_apollo=not args.no_apollo,
            concurrency=conc,
        )
        all_sum.append(s)

    out = ROOT / "output" / "_gfirst_seven_markets_summary.json"
    out.write_text(json.dumps(all_sum, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nWrote {out}", flush=True)


if __name__ == "__main__":
    asyncio.run(main_async())
