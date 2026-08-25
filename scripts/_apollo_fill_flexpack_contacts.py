# -*- coding: utf-8 -*-
"""Fill Contact Person / Role / LinkedIn for Global Flexible Packaging FINAL via Apollo (+ DeepSeek).

Only sales / procurement leadership (Head/VP Sales, Head/VP Procurement, Sales Director, etc.).
Partial/obfuscated Apollo names allowed. Never invent emails. Person LinkedIn only (/in/).
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

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

load_dotenv(ROOT / ".env")

XLSX = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
CACHE = XLSX.parent / "flexpack_apollo_contact_cache.json"
PROGRESS = XLSX.parent / "flexpack_apollo_contact_progress.jsonl"
SUMMARY = XLSX.parent / "flexpack_apollo_contact_summary.json"

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
BAD_NAME = re.compile(
    r"^(ceo|cfo|cto|coo|ciso|president|founder|leadership|board|team|"
    r"contact|sales|procurement|manager|director|vp|head)\b",
    re.I,
)
ROLE_OK = re.compile(
    r"(head of sales|sales head|"
    r"\bvp\b.{0,25}sales|vice president.{0,25}sales|svp.{0,25}sales|"
    r"sales director|director of sales|director.{0,20}sales|"
    r"head of procurement|procurement head|"
    r"\bvp\b.{0,25}procurement|vice president.{0,25}procurement|"
    r"procurement director|director of procurement|director.{0,20}procurement|"
    r"head of purchasing|purchasing director|"
    r"chief (?:revenue|commercial|procurement)|\bcro\b|\bcpo\b|"
    r"commercial director|head of (?:business development|channel|partner)|"
    r"regional (?:vice )?president.{0,20}sales|national sales)",
    re.I,
)
ROLE_BAD = re.compile(
    r"\b(ceo|cfo|cto|coo|ciso|founder|chairman|owner|hr|marketing|accountant)\b",
    re.I,
)


def _empty(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


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


def _role_is_target(role: str) -> bool:
    r = (role or "").strip()
    if not r:
        return False
    if ROLE_BAD.search(r) and not ROLE_OK.search(r):
        return False
    return bool(ROLE_OK.search(r))


def _needs_contact(row: dict[str, str]) -> bool:
    cp = str(row.get("Contact Person") or "").strip()
    role = str(row.get("Role") or "").strip()
    li = str(row.get("LinkedIn") or "").strip()
    # Already have a real sales/procurement contact — only chase LinkedIn if missing
    if (
        not _empty(cp)
        and not BAD_NAME.match(cp)
        and _role_is_target(role)
        and "linkedin.com/in/" in li.lower()
    ):
        return False
    if (
        not _empty(cp)
        and not BAD_NAME.match(cp)
        and _role_is_target(role)
        and ("linkedin.com/in/" in li.lower() or li.lower() == "not publicly disclosed" or _empty(li))
    ):
        # Have name+role; only need LinkedIn upgrade if company URL or empty
        if "/company/" in li.lower() or _empty(li):
            return True
        # NPD LinkedIn is acceptable — skip re-query
        if li.lower() == "not publicly disclosed":
            return False
        return False
    if _empty(cp) or BAD_NAME.match(cp):
        return True
    if _empty(role) or not _role_is_target(role):
        return True
    if _empty(li) or "/company/" in li.lower():
        return True
    return False


def _should_replace_name(existing_cp: str, existing_role: str, person: str) -> bool:
    if _empty(existing_cp) or BAD_NAME.match(existing_cp):
        return True
    if not _role_is_target(existing_role):
        return True
    # Upgrade obfuscated → fuller name only when clearly same or empty last
    if "*" in existing_cp and "*" not in person:
        return True
    # Do NOT replace a solid existing name with a different Apollo candidate
    return False


def _name_ok(name: str) -> bool:
    """At least a usable first name; partial/obfuscated last names OK."""
    s = re.sub(r"\s+", " ", (name or "").strip())
    if not s or BAD_NAME.match(s):
        return False
    first = s.split()[0]
    letters = re.sub(r"[^A-Za-z]", "", first)
    return len(letters) >= 2


def load_rows(path: Path) -> tuple[list[str], int, list[dict[str, Any]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value or "").strip() for c in ws[1]]
    rows: list[dict[str, Any]] = []
    for ridx in range(2, ws.max_row + 1):
        d: dict[str, Any] = {"_row": ridx}
        for j, h in enumerate(headers):
            if not h:
                continue
            d[h] = "" if ws.cell(ridx, j + 1).value is None else str(ws.cell(ridx, j + 1).value).strip()
        if d.get("Company"):
            rows.append(d)
    wb.close()
    return headers, 1, rows


def write_rows(path: Path, headers: list[str], rows: list[dict[str, Any]]) -> None:
    wb = load_workbook(path)
    ws = wb.active
    col = {h: i + 1 for i, h in enumerate(headers) if h}
    for r in rows:
        ridx = int(r["_row"])
        for field in ("Contact Person", "Role", "LinkedIn"):
            if field in col and field in r:
                ws.cell(ridx, col[field]).value = r.get(field) or ""
    # light format
    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    hf = PatternFill("solid", fgColor="1F4E79")
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = hf
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    for r in range(2, ws.max_row + 1):
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
    for name, w in (("Contact Person", 22), ("Role", 28), ("LinkedIn", 40)):
        if name in col:
            ws.column_dimensions[get_column_letter(col[name])].width = w
    wb.save(path)
    wb.close()


async def deepseek_sales_contact(company: str, website: str, hq: str) -> dict[str, str] | None:
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        return None
    client = OpenAI(api_key=key, base_url=base, timeout=60.0)
    system = (
        "Find a CURRENT public Head of Sales, VP of Sales, Sales Director, Commercial Director, "
        "Head of Procurement, VP of Procurement, or Procurement Director for this company. "
        "Return ONLY JSON. If not confident of a real public person, found=false. "
        "Never invent LinkedIn URLs. Partial names OK if that is all that is public. "
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
        data = json.loads(resp.choices[0].message.content or "{}")
    except Exception:
        return None
    if not data.get("found"):
        return None
    if str(data.get("confidence") or "").lower() == "low":
        return None
    name = re.sub(r"\s+", " ", str(data.get("name") or "").strip())
    role = re.sub(r"\s+", " ", str(data.get("role") or "").strip())
    li = str(data.get("linkedin") or "").strip()
    if not _name_ok(name):
        return None
    if not _role_is_target(role):
        return None
    out = {"contact_person": name, "contact_role": role, "source": "deepseek"}
    if li and "linkedin.com/in/" in li.lower() and "/company/" not in li.lower():
        out["linkedin"] = li.split("?")[0].rstrip("/")
    return out


async def run(*, limit: int | None, resume: bool, deepseek_fallback: bool, replace_all: bool) -> dict:
    from vendor_intel.enrichment.apollo_contacts import ApolloContactClient

    if not (os.getenv("APOLLO_API_KEY") or "").strip():
        raise SystemExit("APOLLO_API_KEY missing")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = XLSX.with_name(f"{XLSX.stem}_BEFORE_APOLLO_CONTACTS_{stamp}.xlsx")
    bak.write_bytes(XLSX.read_bytes())
    print("backup", bak.name, flush=True)

    headers, _hr, rows = load_rows(XLSX)
    cache: dict[str, Any] = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    done: set[str] = set()
    if resume and PROGRESS.exists():
        # Skip companies that already have usable sales/procurement contact on sheet
        for r in rows:
            cp = str(r.get("Contact Person") or "").strip()
            role = str(r.get("Role") or "").strip()
            if (
                not _empty(cp)
                and not BAD_NAME.match(cp)
                and _role_is_target(role)
                and str(r.get("Company") or "").strip()
            ):
                done.add(str(r.get("Company") or "").strip().lower())
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            co = str(rec.get("company") or "").strip().lower()
            if rec.get("status") in {"filled", "deepseek_filled", "no_match"}:
                # still retry no_match only if sheet contact empty
                if rec.get("status") == "no_match":
                    continue
                done.add(co)

    todo = []
    for r in rows:
        co = str(r.get("Company") or "").strip()
        if resume and co.lower() in done:
            continue
        if not replace_all and not _needs_contact(r):
            continue
        todo.append(r)
    if limit is not None:
        todo = todo[:limit]

    print(
        f"Flexpack contacts: {len(todo)} todo / {len(rows)} rows "
        f"(deepseek={deepseek_fallback}, resume={resume})",
        flush=True,
    )

    # Clear company LinkedIn URLs to NPD up front for rows we're processing without person LI yet
    for r in todo:
        li = str(r.get("LinkedIn") or "")
        if "/company/" in li.lower():
            r["LinkedIn"] = "Not publicly disclosed"

    filled = no_match = errors = ds_n = 0
    delay = float(os.getenv("APOLLO_REQUEST_DELAY_SEC") or "2.0")
    per_company_timeout = float(os.getenv("APOLLO_COMPANY_TIMEOUT_SEC") or "2400")
    rate_limit_sleep = float(os.getenv("APOLLO_RATE_LIMIT_SLEEP_SEC") or "900")

    def _append(rec: dict) -> None:
        with PROGRESS.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    async with ApolloContactClient(verbose=True, request_delay=delay, timeout=45.0) as apollo:
        for i, r in enumerate(todo, 1):
            co = str(r["Company"]).strip()
            website = str(r.get("Website") or "").strip()
            cache_key = co.lower()
            hit: dict[str, Any] | None = None
            status = "live"

            if cache_key in cache and isinstance(cache[cache_key], dict) and cache[cache_key].get("contact_person"):
                hit = cache[cache_key]
                status = "cached"
            else:
                apollo_row = {
                    "company": co,
                    "website": website,
                    "domain": _domain_from_website(website),
                }
                print(f"  [{i}/{len(todo)}] {co[:44]} … Apollo", flush=True)
                try:
                    hit = await asyncio.wait_for(
                        apollo.find_contact(apollo_row, enrich_match=False),
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
                            apollo.find_contact(apollo_row, enrich_match=False),
                            timeout=per_company_timeout,
                        )
                        status = "live"
                    except Exception as exc:
                        hit = {"error": f"{type(exc).__name__}: {exc}"}
                        status = "error"

            need_ds = False
            if not hit or hit.get("error") or not hit.get("contact_person"):
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
                        deepseek_sales_contact(co, website, str(r.get("Headquarters") or "")),
                        timeout=60.0,
                    )
                except Exception:
                    ds = None
                if ds and ds.get("contact_person"):
                    if (
                        not hit
                        or hit.get("error")
                        or not hit.get("contact_person")
                        or (
                            "*" in str(hit.get("contact_person") or "")
                            and "*" not in ds["contact_person"]
                        )
                    ):
                        hit = {**(hit or {}), **ds}
                        status = "deepseek"
                        ds_n += 1
                    elif ds.get("linkedin") and not hit.get("linkedin"):
                        hit["linkedin"] = ds["linkedin"]
                        status = "apollo+ds_li"

            if not hit or hit.get("error") or not hit.get("contact_person"):
                if hit and hit.get("error"):
                    errors += 1
                    print(f"  -> ERROR {str(hit.get('error'))[:100]}", flush=True)
                    _append({"company": co, "status": "error", "error": hit.get("error")})
                else:
                    no_match += 1
                    _append({"company": co, "status": "no_match"})
                    cache[cache_key] = {"skipped": True, "reason": "no_match"}
                    print("  -> no match", flush=True)
                    # keep company LI cleared
                    if "/company/" in str(r.get("LinkedIn") or "").lower() or _empty(r.get("LinkedIn")):
                        r["LinkedIn"] = "Not publicly disclosed"
                if i % 10 == 0:
                    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
                    write_rows(XLSX, headers, rows)
                continue

            person = str(hit.get("contact_person") or "").strip()
            role = str(hit.get("contact_role") or "").strip()
            li = str(hit.get("linkedin") or "").strip()
            if role.count("|") >= 2 or len(role) > 120:
                role = role.split("|")[0].strip()
            if not _name_ok(person):
                no_match += 1
                _append({"company": co, "status": "no_match", "reason": "bad_name", "name": person})
                print(f"  -> skip bad name {person!r}", flush=True)
                continue
            if role and not _role_is_target(role):
                no_match += 1
                _append({"company": co, "status": "no_match", "reason": "bad_role", "role": role})
                print(f"  -> skip non-sales/procurement role {role!r}", flush=True)
                continue

            existing_cp = str(r.get("Contact Person") or "").strip()
            existing_role = str(r.get("Role") or "").strip()
            applied_name = existing_cp
            if replace_all or _should_replace_name(existing_cp, existing_role, person):
                r["Contact Person"] = person
                applied_name = person
            if role and (
                replace_all
                or _empty(existing_role)
                or not _role_is_target(existing_role)
                or applied_name == person
            ):
                # update role when we set/keep matching Apollo person, or role empty/wrong
                if applied_name == person or _empty(existing_role) or not _role_is_target(existing_role):
                    r["Role"] = role
            if li and "linkedin.com/in/" in li.lower() and "/company/" not in li.lower():
                r["LinkedIn"] = li.split("?")[0].rstrip("/")
            else:
                if _empty(r.get("LinkedIn")) or "/company/" in str(r.get("LinkedIn") or "").lower():
                    r["LinkedIn"] = "Not publicly disclosed"

            filled += 1
            note = " [partial]" if "*" in applied_name else ""
            kept = " (kept existing)" if applied_name != person else ""
            print(
                f"  -> {applied_name}{note}{kept}"
                + (f" | {r.get('Role')}" if r.get("Role") else "")
                + (" | LI" if "linkedin.com/in/" in str(r.get("LinkedIn") or "") else ""),
                flush=True,
            )
            _append(
                {
                    "company": co,
                    "status": "deepseek_filled" if status == "deepseek" else "filled",
                    "contact_person": r.get("Contact Person"),
                    "role": r.get("Role"),
                    "linkedin": r.get("LinkedIn"),
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
                    "apollo_person_id",
                    "apollo_name_obfuscated",
                    "source",
                }
            }
            if i % 10 == 0:
                CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
                write_rows(XLSX, headers, rows)
                print(f"  … checkpoint ({filled} filled)", flush=True)

    # Final: ensure company LinkedIn not left as person column junk across full sheet
    for r in rows:
        li = str(r.get("LinkedIn") or "")
        if "/company/" in li.lower():
            r["LinkedIn"] = "Not publicly disclosed"
        if _empty(r.get("Contact Person")):
            r["Contact Person"] = "Not publicly disclosed"
        if _empty(r.get("Role")):
            r["Role"] = "Not publicly disclosed"
        if _empty(r.get("LinkedIn")):
            r["LinkedIn"] = "Not publicly disclosed"

    write_rows(XLSX, headers, rows)
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    # recount
    named = person_li = 0
    for r in rows:
        if not _empty(r.get("Contact Person")) and str(r.get("Contact Person")).lower() != "not publicly disclosed":
            named += 1
        if "linkedin.com/in/" in str(r.get("LinkedIn") or "").lower():
            person_li += 1

    summary = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "todo": len(todo),
        "filled_this_run": filled,
        "deepseek_fills": ds_n,
        "no_match": no_match,
        "errors": errors,
        "sheet_named_contacts": named,
        "sheet_person_linkedin": person_li,
        "rows": len(rows),
        "backup": bak.name,
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print("DONE", summary, flush=True)
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--resume", action="store_true", default=True)
    ap.add_argument("--no-resume", action="store_true")
    ap.add_argument("--deepseek-fallback", action="store_true", default=True)
    ap.add_argument("--no-deepseek", action="store_true")
    ap.add_argument("--replace-all", action="store_true")
    args = ap.parse_args()
    resume = False if args.no_resume else args.resume
    ds = False if args.no_deepseek else args.deepseek_fallback
    asyncio.run(
        run(
            limit=args.limit,
            resume=resume,
            deepseek_fallback=ds,
            replace_all=args.replace_all,
        )
    )


if __name__ == "__main__":
    main()
