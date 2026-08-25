#!/usr/bin/env python3
"""Fill Contact Person / Role / LinkedIn on 7 market Excels via Apollo (+ optional DeepSeek).

Targets sales / commercial / procurement leadership only (see apollo_contacts.py).
Never invents emails/phones. Gap-fills empty/NPD contacts; resume + cache supported.

Usage:
  .venv\\Scripts\\python.exe scripts\\apollo_fill_seven_markets_contacts.py
  .venv\\Scripts\\python.exe scripts\\apollo_fill_seven_markets_contacts.py --market electrical --limit 30
  .venv\\Scripts\\python.exe scripts\\apollo_fill_seven_markets_contacts.py --enrich
  .venv\\Scripts\\python.exe scripts\\apollo_fill_seven_markets_contacts.py --deepseek-fallback
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
            k.startswith(("OPENAI_", "APOLLO_", "LLM_"))
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

# Roles we refuse to keep as Contact Person (title dumped into name cell)
BAD_NAME = re.compile(
    r"^(ceo|cfo|cto|coo|ciso|president|founder|leadership|board|team|"
    r"contact|sales|procurement|manager|director|vp|head)\b",
    re.I,
)


def _empty(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


def _needs_contact(row: dict[str, str]) -> bool:
    cp = str(row.get("Contact Person") or "").strip()
    if _empty(cp):
        return True
    if BAD_NAME.match(cp):
        return True
    if any(x in cp.lower() for x in ("leadership", "executive team", "not available")):
        return True
    # Need LinkedIn person URL still empty while name exists → still try Apollo for LI
    li = str(row.get("LinkedIn") or "").strip()
    role = str(row.get("Role") or "").strip()
    if _empty(role) or _empty(li) or "/company/" in li.lower():
        return True
    return False


def _domain_from_website(website: str) -> str:
    s = (website or "").strip().lower()
    if not s:
        return ""
    if "://" not in s:
        s = "https://" + s
    try:
        host = urlparse(s).netloc or ""
    except Exception:
        host = ""
    return host.removeprefix("www.").split("/")[0]


def _find_header(ws) -> tuple[list[str], int]:
    for r in range(1, 6):
        vals = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=r, max_row=r))]
        if "Company" in vals and "Contact Person" in vals:
            return vals, r
    raise RuntimeError("No Company/Contact Person header")


def load_xlsx(path: Path) -> tuple[list[str], int, list[dict[str, Any]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
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
    wb = load_workbook(path)
    ws = wb.active
    col = {h: i + 1 for i, h in enumerate(headers) if h}
    for r in rows:
        ridx = int(r["_row"])
        for field in ("Contact Person", "Role", "LinkedIn"):
            if field not in col:
                continue
            if field in r:
                ws.cell(ridx, col[field]).value = r.get(field) or ""
    for cell in ws[hr]:
        cell.font = Font(bold=True)
    wb.save(path)
    wb.close()


def _cache_paths(xlsx: Path) -> tuple[Path, Path, Path]:
    d = xlsx.parent
    stem = xlsx.stem
    return (
        d / f"{stem}_apollo_contact_cache.json",
        d / f"{stem}_apollo_contact_progress.jsonl",
        d / f"{stem}_apollo_contact_summary.json",
    )


async def _deepseek_sales_contact(
    company: str, website: str, hq: str
) -> dict[str, str] | None:
    """Last-resort DeepSeek lookup — only accept sales/procurement titles."""
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key or "deepseek" not in base.lower():
        return None
    client = OpenAI(api_key=key, base_url=base)
    system = (
        "Find the current Head of Sales, VP Sales, Sales Director, Commercial Director, "
        "Head of Procurement, or Procurement Manager for the company. "
        "Return ONLY JSON. If you are not confident of a real public person, set found=false. "
        "Never invent LinkedIn URLs. Prefer linkedin.com/in/ when known from public knowledge. "
        'Schema: {"found":true|false,"name":"","role":"","linkedin":"","confidence":"high|medium|low"}'
    )
    user = f"Company: {company}\nWebsite: {website}\nHQ: {hq}\n"
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
        data = json.loads(raw)
    except Exception:
        return None
    if not data.get("found"):
        return None
    if str(data.get("confidence") or "").lower() == "low":
        return None
    name = str(data.get("name") or "").strip()
    role = str(data.get("role") or "").strip()
    li = str(data.get("linkedin") or "").strip()
    name = re.sub(r"\s+", " ", name.replace("\n", " ")).strip()
    role = re.sub(r"\s+", " ", role.replace("\n", " ")).strip()
    if not name or len(name.split()) < 2:
        return None
    if re.search(r"location|linkedin|http|www\.", name, re.I):
        return None
    # Must look like sales/procurement role
    rl = role.lower()
    if not re.search(
        r"sales|commercial|revenue|procurement|purchasing|business development|channel|partner",
        rl,
    ):
        return None
    if re.search(r"\b(ceo|cfo|cto|coo|founder|chairman|president|managing director)\b", rl) and not re.search(
        r"sales|commercial|revenue|procurement", rl
    ):
        return None
    if role.count("|") >= 2:
        return None
    out = {"contact_person": name, "contact_role": role, "source": "deepseek"}
    if li and "linkedin.com/in/" in li.lower() and "linkedin.com/company/" not in li.lower():
        out["linkedin"] = li.split("?")[0].rstrip("/")
    return out


async def fill_market(
    *,
    key: str,
    path: Path,
    replace_all: bool,
    limit: int | None,
    resume: bool,
    enrich: bool,
    deepseek_fallback: bool,
    deepseek_only: bool = False,
) -> dict[str, Any]:
    from vendor_intel.enrichment.apollo_contacts import ApolloContactClient

    if not path.exists():
        raise FileNotFoundError(path)

    bak = path.with_name(path.stem + "_BEFORE_APOLLO_CONTACTS.xlsx")
    if not bak.exists():
        bak.write_bytes(path.read_bytes())
        print(f"  backup {bak.name}", flush=True)

    headers, hr, rows = load_xlsx(path)
    cache_path, progress_path, summary_path = _cache_paths(path)
    cache: dict[str, Any] = {}
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    done: set[str] = set()
    if resume and progress_path.exists():
        # Only skip if progress says filled AND Excel still has a Contact Person
        still_has = {
            str(r.get("Company") or "").strip().lower()
            for r in rows
            if not _empty(r.get("Contact Person"))
        }
        for line in progress_path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            co = str(rec.get("company") or "").strip().lower()
            if rec.get("status") in {"filled", "deepseek_filled"} and co in still_has:
                done.add(co)
            elif rec.get("status") == "skipped_existing" and co in still_has:
                done.add(co)

    todo = []
    for r in rows:
        co = r["Company"]
        if resume and co.strip().lower() in done:
            continue
        if not replace_all and not _needs_contact(r):
            continue
        todo.append(r)
    if limit is not None:
        todo = todo[:limit]

    print(
        f"\n=== {key}: {len(todo)} todo / {len(rows)} rows "
        f"(enrich={enrich}, deepseek={deepseek_fallback}, deepseek_only={deepseek_only}) ===",
        flush=True,
    )

    filled = 0
    no_match = 0
    errors = 0
    ds_filled = 0
    delay = float(os.getenv("APOLLO_REQUEST_DELAY_SEC") or "2.5")
    per_company_timeout = float(os.getenv("APOLLO_COMPANY_TIMEOUT_SEC") or "2400")
    rate_limit_sleep = float(os.getenv("APOLLO_RATE_LIMIT_SLEEP_SEC") or "1200")
    deepseek_timeout = float(os.getenv("DEEPSEEK_CONTACT_TIMEOUT_SEC") or "60")

    def _append(rec: dict) -> None:
        with progress_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    apollo_cm = None if deepseek_only else ApolloContactClient(
        verbose=True, request_delay=delay, timeout=45.0
    )

    async def _process(apollo: ApolloContactClient | None) -> None:
        nonlocal filled, no_match, errors, ds_filled
        for i, r in enumerate(todo, 1):
            co = r["Company"]
            website = str(r.get("Website") or "").strip()
            cache_key = co.strip().lower()
            hit: dict[str, Any] | None = None
            status = "live"

            if (
                cache_key in cache
                and isinstance(cache[cache_key], dict)
                and cache[cache_key].get("contact_person")
                and not deepseek_only
            ):
                hit = cache[cache_key]
                status = "cached"
            elif (
                deepseek_only
                and cache_key in cache
                and isinstance(cache[cache_key], dict)
                and cache[cache_key].get("contact_person")
                and "*" not in str(cache[cache_key].get("contact_person") or "")
                and cache[cache_key].get("source") == "deepseek"
            ):
                hit = cache[cache_key]
                status = "cached"
            elif not deepseek_only and apollo is not None:
                apollo_row = {
                    "company": co,
                    "website": website,
                    "domain": _domain_from_website(website),
                }
                print(f"  [{i}/{len(todo)}] {co[:44]} … Apollo", flush=True)
                try:
                    hit = await asyncio.wait_for(
                        apollo.find_contact(apollo_row, enrich_match=enrich),
                        timeout=per_company_timeout,
                    )
                except asyncio.TimeoutError:
                    hit = {"error": f"Timeout after {per_company_timeout:.0f}s"}
                    status = "error"
                except Exception as exc:
                    hit = {"error": f"{type(exc).__name__}: {exc}"}
                    status = "error"

                err_s = str((hit or {}).get("error") or "")
                if hit and ("429" in err_s or "rate limit" in err_s.lower()):
                    print(f"  !! rate-limit cool-down {rate_limit_sleep:.0f}s", flush=True)
                    await asyncio.sleep(rate_limit_sleep)
                    try:
                        hit = await asyncio.wait_for(
                            apollo.find_contact(apollo_row, enrich_match=enrich),
                            timeout=per_company_timeout,
                        )
                        status = "live"
                    except Exception as exc:
                        hit = {"error": f"{type(exc).__name__}: {exc}"}
                        status = "error"

            # DeepSeek fallback when Apollo empty OR only partial name without LinkedIn
            need_ds = False
            if deepseek_only:
                need_ds = True
            elif not hit or hit.get("error") or not hit.get("contact_person"):
                need_ds = deepseek_fallback
            elif (
                deepseek_fallback
                and "*" in str(hit.get("contact_person") or "")
                and "linkedin.com/in/" not in str(hit.get("linkedin") or "").lower()
            ):
                need_ds = True

            if need_ds:
                print(f"  [{i}/{len(todo)}] {co[:44]} … DeepSeek", flush=True)
                try:
                    ds = await asyncio.wait_for(
                        _deepseek_sales_contact(
                            co, website, str(r.get("Headquarters") or "")
                        ),
                        timeout=deepseek_timeout,
                    )
                except Exception:
                    ds = None
                if ds and ds.get("contact_person"):
                    if (
                        deepseek_only
                        or not hit
                        or hit.get("error")
                        or not hit.get("contact_person")
                        or (
                            "*" in str(hit.get("contact_person") or "")
                            and "*" not in ds["contact_person"]
                        )
                    ):
                        hit = {**(hit or {}), **ds}
                        status = "deepseek"
                        ds_filled += 1
                    elif ds.get("linkedin") and not hit.get("linkedin"):
                        hit["linkedin"] = ds["linkedin"]
                        status = "apollo+deepseek_li"
                elif deepseek_only:
                    hit = None
            # Clear Apollo timeout errors so DeepSeek result (if any) is used cleanly
            if hit and hit.get("error") and not hit.get("contact_person"):
                if "rate limit" in str(hit.get("error")).lower() or "429" in str(hit.get("error")):
                    # Don't burn DeepSeek on every 429 — cool down account-wide once
                    print(f"  !! account cool-down {rate_limit_sleep:.0f}s", flush=True)
                    await asyncio.sleep(rate_limit_sleep)

            if not hit or hit.get("error") or not hit.get("contact_person"):
                if hit and hit.get("error"):
                    errors += 1
                    print(f"  -> ERROR {str(hit.get('error'))[:100]}", flush=True)
                else:
                    no_match += 1
                    _append({"company": co, "status": "no_match"})
                    cache[cache_key] = {"skipped": True, "reason": "no_match"}
                    print("  -> no match", flush=True)
                if i % 15 == 0:
                    cache_path.write_text(
                        json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8"
                    )
                    write_contacts(path, headers, hr, rows)
                continue

            person = str(hit.get("contact_person") or "").strip()
            role = str(hit.get("contact_role") or "").strip()
            li = str(hit.get("linkedin") or "").strip()
            if role.count("|") >= 2 or len(role) > 120:
                role = role.split("|")[0].strip()
            if len(person.split()) < 2 and "*" not in person:
                no_match += 1
                _append({"company": co, "status": "no_match", "reason": "incomplete_name"})
                print(f"  -> skip incomplete name {person!r}", flush=True)
                continue

            # Apply gap-fill rules (upgrade obfuscated Apollo names when DeepSeek has full name)
            existing_cp = str(r.get("Contact Person") or "").strip()
            if (
                replace_all
                or _empty(existing_cp)
                or BAD_NAME.match(existing_cp)
                or ("*" in existing_cp and "*" not in person)
            ):
                r["Contact Person"] = person
            if role and (replace_all or _empty(r.get("Role"))):
                r["Role"] = role
            if li and "linkedin.com/in/" in li.lower():
                if replace_all or _empty(r.get("LinkedIn")) or "/company/" in str(
                    r.get("LinkedIn") or ""
                ).lower():
                    r["LinkedIn"] = li
            elif _empty(r.get("LinkedIn")):
                r["LinkedIn"] = "Not publicly disclosed"

            filled += 1
            note = " [partial]" if "*" in person else ""
            print(
                f"  -> {person}{note}"
                + (f" | {role}" if role else "")
                + (" | LI" if li and "/in/" in li else ""),
                flush=True,
            )
            _append(
                {
                    "company": co,
                    "status": "deepseek_filled" if status == "deepseek" else "filled",
                    "contact_person": person,
                    "role": role,
                    "linkedin": li,
                    "source": status,
                }
            )
            cache[cache_key] = {
                k: v
                for k, v in hit.items()
                if k
                in {
                    "contact_person",
                    "contact_role",
                    "linkedin",
                    "email",
                    "apollo_person_id",
                    "apollo_name_obfuscated",
                    "source",
                }
            }

            if i % 15 == 0:
                cache_path.write_text(
                    json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8"
                )
                write_contacts(path, headers, hr, rows)
                print(f"  … checkpoint ({filled} filled)", flush=True)

    if apollo_cm is None:
        await _process(None)
    else:
        async with apollo_cm as apollo:
            await _process(apollo)

    cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    write_contacts(path, headers, hr, rows)

    still_empty = sum(1 for r in rows if _empty(r.get("Contact Person")))
    summary = {
        "when": datetime.now(timezone.utc).isoformat(),
        "market": key,
        "path": str(path),
        "todo": len(todo),
        "filled": filled,
        "deepseek_filled": ds_filled,
        "no_match": no_match,
        "errors": errors,
        "still_empty_contact": still_empty,
        "total_rows": len(rows),
        "mode": "deepseek_only" if deepseek_only else "apollo",
    }
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(
        f"=== {key} done: filled={filled} ds={ds_filled} no_match={no_match} "
        f"errors={errors} still_empty={still_empty} ===",
        flush=True,
    )
    return summary


async def main_async() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--market", default="", help="Comma keys: electrical,adc,avocado,...")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--replace-all", action="store_true")
    ap.add_argument("--no-resume", action="store_true")
    ap.add_argument("--enrich", action="store_true", help="Use Apollo people/match credits for full name+LI")
    ap.add_argument("--deepseek-fallback", action="store_true")
    ap.add_argument(
        "--deepseek-only",
        action="store_true",
        help="Skip Apollo; fill empty contacts via DeepSeek only (sales/procurement)",
    )
    args = ap.parse_args()

    if args.deepseek_only:
        args.deepseek_fallback = True
    if not args.deepseek_only and not (os.getenv("APOLLO_API_KEY") or "").strip():
        raise SystemExit("APOLLO_API_KEY missing")

    keys = [k.strip() for k in args.market.split(",") if k.strip()] if args.market else [
        m["key"] for m in MARKETS
    ]
    selected = [m for m in MARKETS if m["key"] in keys]
    if not selected:
        raise SystemExit(f"No markets matched {keys}")

    limit = args.limit or None
    all_sum: list[dict] = []
    for m in selected:
        s = await fill_market(
            key=m["key"],
            path=m["path"],
            replace_all=args.replace_all,
            limit=limit,
            resume=not args.no_resume,
            enrich=args.enrich,
            deepseek_fallback=args.deepseek_fallback or args.deepseek_only,
            deepseek_only=args.deepseek_only,
        )
        all_sum.append(s)

    out = ROOT / "output" / "_apollo_seven_markets_summary.json"
    out.write_text(json.dumps(all_sum, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nWrote {out}", flush=True)


if __name__ == "__main__":
    asyncio.run(main_async())
