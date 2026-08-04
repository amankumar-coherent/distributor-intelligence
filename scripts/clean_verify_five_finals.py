#!/usr/bin/env python3
"""Clean + verify 5 FINAL landscape Excels: keep distributors/channel only,
fill empty Contact Person (and Role when missing) via live search + LLM extract.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

load_dotenv(ROOT / ".env")

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
    "todo",
    "not applicable",
}

OEM_BY_MARKET: dict[str, set[str]] = {
    "smartwatch": {
        "apple",
        "samsung",
        "garmin",
        "fitbit",
        "amazfit",
        "huawei",
        "xiaomi",
        "google",
        "withings",
        "polar",
        "suunto",
        "oppo",
        "oneplus",
        "honor",
    },
    "smartphone": {
        "apple",
        "samsung",
        "huawei",
        "xiaomi",
        "oppo",
        "vivo",
        "oneplus",
        "google",
        "motorola",
        "nokia",
        "sony",
        "realme",
        "honor",
        "tecno",
        "infinix",
        "nothing",
        "samsunghub",
    },
    "firewall": {
        "cisco",
        "fortinet",
        "palo alto networks",
        "palo alto",
        "check point",
        "sophos",
        "juniper",
        "watchguard",
        "sonicwall",
        "barracuda",
        "forcepoint",
        "zscaler",
        "cloudflare",
        "f5 networks",
    },
    "green": {
        "basf",
        "dow chemical",
        "dow inc",
        "dupont",
        "evonik",
        "sabic",
        "covestro",
        "lyondellbasell",
        "solvay",
        "arkema",
        "clariant",
        "huntsman",
    },
    "avocado": {
        "chosen foods",
        "betterbody foods",
    },
}

MARKET_KW: dict[str, tuple[str, ...]] = {
    "smartwatch": ("smartwatch", "wearable", "fitness tracker", "watch", "band"),
    "smartphone": ("smartphone", "mobile phone", "handset", "cellular", "android", "iphone"),
    "firewall": ("firewall", "ngfw", "network security", "cyber", "utm", "sase", "sd-wan"),
    "green": ("green chemical", "bio-based", "sustainable", "specialty chemical", "solvent", "surfactant"),
    "avocado": ("avocado", "edible oil", "specialty oil", "foodservice", "grocery", "oil"),
}

DIST_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|vad|var|channel|"
    r"partner|importer|import|retailer|retail|dealer|integrator|mssp|broker|"
    r"authorized|value[\s-]?added|cash\s*&\s*carry|foodservice)\b",
    re.I,
)
MFR_RE = re.compile(
    r"\b(manufacturer|manufacturing oem|oem brand|producer only|brand owner|"
    r"device maker|handset maker)\b",
    re.I,
)

MARKETS = [
    {
        "key": "smartwatch",
        "path": ROOT
        / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx",
        "sheet": "Smartwatch Companies",
        "query": "Global Smartwatch Market",
    },
    {
        "key": "green",
        "path": ROOT
        / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL.xlsx",
        "sheet": "Green Chemical Companies",
        "query": "Global Green Chemical Market",
    },
    {
        "key": "firewall",
        "path": ROOT
        / "output/run4/global_firewall_market_global/global_firewall_market_FINAL.xlsx",
        "sheet": "Firewall Companies",
        "query": "Global Firewall Market",
    },
    {
        "key": "smartphone",
        "path": ROOT
        / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx",
        "sheet": "Companies",
        "query": "Global Smartphone Market",
    },
    {
        "key": "avocado",
        "path": ROOT
        / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL.xlsx",
        "sheet": "Avocado Oil Distributors",
        "query": "Global Avocado Oil Market",
    },
]


def _empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def _norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def load_sheet_rows(path: Path, sheet: str) -> tuple[list[str], list[dict[str, str]], int]:
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = None
    headers: list[str] = []
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            header_idx = i
            headers = vals
            break
    if header_idx is None:
        raise RuntimeError(f"No Company header in {path.name}/{sheet}")
    data: list[dict[str, str]] = []
    for r in rows[header_idx + 1 :]:
        d = {
            headers[j]: ("" if j >= len(r) or r[j] is None else str(r[j]).strip())
            for j in range(len(headers))
            if headers[j]
        }
        if d.get("Company"):
            data.append(d)
    return headers, data, header_idx


def is_oem_or_manufacturer(market: str, row: dict[str, str]) -> tuple[bool, str]:
    name = row.get("Company") or ""
    bt = row.get("Business Type") or ""
    focus = row.get("Specialty Focus") or ""
    summary = row.get("Summary") or ""
    nl = _norm_name(name)

    # Keep clear channel wording in the company name
    if DIST_RE.search(name) and not MFR_RE.search(bt):
        pass
    else:
        for oem in OEM_BY_MARKET.get(market, set()):
            if nl == oem or nl.startswith(oem + " ") or f" {oem} " in f" {nl} ":
                # subsidiaries that are pure OEM sales arms
                if not DIST_RE.search(name):
                    return True, f"oem_name:{oem}"

    if MFR_RE.search(bt) and not DIST_RE.search(bt):
        return True, "business_type_manufacturer"

    blob = f"{bt} {focus} {summary}".lower()
    if "manufacturer only" in blob or "oem only" in blob:
        return True, "text_manufacturer_only"
    return False, ""


def looks_off_market(market: str, row: dict[str, str]) -> tuple[bool, str]:
    """Soft relevance: only drop obvious junk, not borderline channel firms."""
    name = (row.get("Company") or "").lower()
    focus = (row.get("Specialty Focus") or "").lower()
    summary = (row.get("Summary") or "").lower()
    cats = (row.get("Core Categories") or "").lower()
    blob = f"{name} {focus} {summary} {cats}"

    junk = (
        "wikipedia",
        "linkedin",
        "glassdoor",
        "indeed.com",
        "market research report",
        "statista",
        "grand view research",
    )
    if any(j in blob for j in junk):
        return True, "junk_source"

    # Smartphone: drop pure smartwatch-only if no phone signal and watch-only
    if market == "smartphone":
        if re.search(r"\b(smartwatch|fitness tracker)\b", blob) and not re.search(
            r"\b(smartphone|mobile|handset|phone|cellular)\b", blob
        ):
            if any(x in name for x in ("garmin", "fitbit", "amazfit")):
                return True, "watch_oem_not_phone"
        if re.search(r"\b(software|saas|erp|crm)\b", name) and not re.search(
            r"\b(phone|mobile|handset|telecom|electronics|distributor)\b", blob
        ):
            return True, "software_not_phone_channel"

    if market == "firewall":
        # Drop consumer IT retail without security signal
        if re.search(r"\b(avocado|grocery|foodservice|cosmetic)\b", blob) and not re.search(
            r"\b(firewall|security|cyber|network)\b", blob
        ):
            return True, "off_vertical"

    if market == "green":
        if re.search(r"\b(smartphone|firewall|gaming console)\b", blob) and not re.search(
            r"\b(chemical|solvent|bio.?based|ingredient)\b", blob
        ):
            return True, "off_vertical"

    return False, ""


def is_garbage_contact(name: str, company: str = "") -> bool:
    """True only for clearly invalid contact strings (for clearing existing cells)."""
    n = (name or "").strip()
    if not n or n.lower() in NPD:
        return True
    if "\n" in n or "\t" in n:
        return True
    if re.search(r"https?://|www\.|@", n, re.I):
        return True
    if re.search(r"\b(appoints|appointed|visionary|experience)\b", n, re.I):
        return True
    if company and _norm_name(n) == _norm_name(company):
        return True
    # title-only masquerading as person
    if re.fullmatch(
        r"(ceo|cto|cfo|founder|president|managing director|sales director)",
        n,
        re.I,
    ):
        return True
    return False


def looks_like_person_name(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 60:
        return False
    if n.lower() in NPD:
        return False
    if re.search(r"[\d@]|https?://|www\.", n, re.I):
        return False
    if "\n" in n or "\t" in n:
        return False
    parts = [p for p in re.split(r"\s+", n) if p]
    if not (2 <= len(parts) <= 5):
        return False
    bad_tokens = {
        "ceo",
        "cto",
        "cfo",
        "appoints",
        "appointed",
        "as",
        "experience",
        "visionary",
        "strategic",
        "executive",
        "team",
        "leadership",
        "company",
        "ltd",
        "llc",
        "inc",
        "gmbh",
        "pvt",
        "founder",
        "president",
        "director",
        "managing",
        "global",
        "solutions",
        "mumbai",
        "delhi",
        "india",
        "usa",
        "uk",
    }
    low = {p.lower().strip(".,") for p in parts}
    if low & bad_tokens:
        return False
    # mostly alphabetic tokens
    alphaish = sum(1 for p in parts if re.match(r"^[A-Za-z][A-Za-z.'\-]+$", p))
    if alphaish < len(parts):
        return False
    if company and _norm_name(n) == _norm_name(company):
        return False
    if company and _norm_name(n) in _norm_name(company):
        return False
    return True


async def resolve_contact(
    *,
    company: str,
    website: str,
    market_query: str,
    router: Any,
    client: Any,
    model: str,
    apollo: Any | None = None,
) -> dict[str, str]:
    """Return Contact Person / Role via OpenAI web_search (primary) or Apollo."""
    from vendor_intel.pipeline.chatgpt_expand import _web_search_json

    # 1) Optional Apollo
    use_apollo = os.getenv("CONTACT_USE_APOLLO", "").lower() in ("1", "true", "yes")
    if use_apollo and (os.getenv("APOLLO_API_KEY") or "").strip():
        try:
            if apollo is not None:
                hit = await apollo.find_contact(
                    {"company": company, "website": website, "domain": ""},
                    enrich_match=True,
                )
                if hit and not hit.get("error"):
                    name = str(hit.get("contact_person") or "").strip()
                    role = str(hit.get("contact_role") or "").strip()
                    if name and looks_like_person_name(name, company) and "*" not in name:
                        out = {"Contact Person": name, "_via": "apollo"}
                        if role:
                            out["Role"] = role
                        if hit.get("email"):
                            out["Email"] = str(hit["email"])
                        return out
        except Exception:
            pass

    # 2) OpenAI Responses API + live web_search (best accuracy)
    try:
        data = await asyncio.to_thread(
            _web_search_json,
            client,
            model,
            (
                "Find the public leadership or sales contact for a distributor / "
                "channel / wholesale company. Use live web search. "
                "Return JSON only: "
                '{"contact_person":"Full Name or empty","role":"Title or empty",'
                '"email":"public email or empty","confidence":"high|medium|low"} '
                "Prefer CEO, Founder, Managing Director, or Sales Director. "
                "Do NOT invent names. If not found, empty strings."
            ),
            (
                f"Market: {market_query}\nCompany: {company}\nWebsite: {website}\n\n"
                f'Search: "{company}" CEO OR founder OR "managing director" OR '
                f'"sales director" LinkedIn\n'
                f'Search: "{company}" leadership team\n'
                f"Search site LinkedIn company people for {company}"
            ),
            label="contact-web",
            max_tokens=500,
        )
    except Exception:
        data = {}

    if isinstance(data, dict):
        person = str(data.get("contact_person") or "").strip()
        role = str(data.get("role") or "").strip()
        email = str(data.get("email") or "").strip()
        if looks_like_person_name(person, company):
            out = {"Contact Person": person, "_via": "openai_web_search"}
            if role and role.lower() not in NPD and len(role) < 80:
                if not re.search(r"\n|appoints|visionary", role, re.I):
                    out["Role"] = role
            if email and "@" in email and email.lower() not in NPD:
                out["Email"] = email
            return out

    # 3) Fallback: local search snippets + plain chat JSON
    from vendor_intel.pipeline.chatgpt_expand import _chat_json

    q = f'"{company}" CEO OR founder OR "managing director" OR "sales director"'
    snippets: list[str] = []
    try:
        hits = await router.search(q)
    except Exception:
        hits = []
    for h in hits or []:
        title = getattr(h, "title", None) or (h.get("title") if isinstance(h, dict) else "") or ""
        snip = getattr(h, "snippet", None) or (h.get("snippet") if isinstance(h, dict) else "") or ""
        link = getattr(h, "link", None) or getattr(h, "url", None) or ""
        if isinstance(h, dict):
            link = link or h.get("link") or h.get("url") or ""
        snippets.append(f"{title}\n{snip}\n{link}")
    text = "\n\n".join(snippets)[:7000]
    if len(text) < 40:
        return {}
    try:
        data2 = _chat_json(
            client,
            model,
            system=(
                "Extract public leadership contact. JSON only: "
                '{"contact_person":"","role":"","email":"","confidence":"low"} '
                "Do NOT invent."
            ),
            user=f"Company: {company}\nWebsite: {website}\n\nSnippets:\n{text}",
            label="contact-fill",
            max_tokens=400,
        )
    except Exception:
        return {}
    if not isinstance(data2, dict):
        return {}
    person = str(data2.get("contact_person") or "").strip()
    role = str(data2.get("role") or "").strip()
    email = str(data2.get("email") or "").strip()
    if not looks_like_person_name(person, company):
        return {}
    out = {"Contact Person": person, "_via": "search_llm"}
    if role and role.lower() not in NPD and len(role) < 80:
        if not re.search(r"\n|appoints|visionary", role, re.I):
            out["Role"] = role
    if email and "@" in email and email.lower() not in NPD:
        out["Email"] = email
    return out


def write_company_sheet(
    path_out: Path,
    *,
    sheet_name: str,
    headers: list[str],
    rows: list[dict[str, str]],
    audit: dict[str, Any],
    title_banner: str = "",
) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    cell_font = Font(name="Calibri", size=9)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")
    start_row = 1
    if title_banner:
        ws.append([title_banner])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(8, len(headers) or 8))
        start_row = 2
        ws.append(headers)
    else:
        ws.append(headers)

    hdr_r = start_row if title_banner else 1
    for i, h in enumerate(headers, 1):
        cell = ws.cell(hdr_r, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
        ws.column_dimensions[get_column_letter(i)].width = 16
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
        r = ws.max_row
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = cell_font
            cell.border = thin
            cell.alignment = wrap
        ws.row_dimensions[r].height = 32
    ws.freeze_panes = f"A{hdr_r + 1}"
    last = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A{hdr_r}:{last}{ws.max_row}"

    aws = wb.create_sheet("Clean Audit")
    aws.append(["Field", "Value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])
    path_out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path_out)


async def process_market(m: dict[str, Any], *, fill_contacts: bool, limit_contacts: int | None) -> dict[str, Any]:
    from openai import OpenAI

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings

    key = m["key"]
    path: Path = m["path"]
    cleaned_path = path.with_name(path.stem + "_CLEANED.xlsx")
    # Prefer prior cleaned workbook if present (resume-friendly)
    src = cleaned_path if cleaned_path.exists() else path
    print(f"\n===== {key} | {src.name} =====", flush=True)
    headers, rows, _ = load_sheet_rows(src, m["sheet"])
    before = len(rows)

    kept: list[dict[str, str]] = []
    removed: list[dict[str, str]] = []
    for row in rows:
        oem, why = is_oem_or_manufacturer(key, row)
        off, why2 = looks_off_market(key, row)
        if oem or off:
            removed.append({**row, "_remove_reason": why or why2})
            continue
        bt = row.get("Business Type") or ""
        if MFR_RE.search(bt) and not DIST_RE.search(bt):
            removed.append({**row, "_remove_reason": "manufacturer_bt"})
            continue
        kept.append(row)

    empty_contact_before = sum(1 for r in kept if _empty(r.get("Contact Person")))
    filled_n = 0
    ckpt_path = path.parent / f"contact_fill_checkpoint_{key}.json"
    ckpt: dict[str, dict[str, str]] = {}
    if ckpt_path.exists():
        try:
            ckpt = json.loads(ckpt_path.read_text(encoding="utf-8"))
        except Exception:
            ckpt = {}
    # Apply checkpoint to kept rows
    for row in kept:
        cname = row.get("Company") or ""
        if cname in ckpt and _empty(row.get("Contact Person")) and ckpt[cname].get("Contact Person"):
            row.update({k: v for k, v in ckpt[cname].items() if k in ("Contact Person", "Role", "Email")})
    # Clear only clearly garbage prior contact fills
    for row in kept:
        cp = row.get("Contact Person") or ""
        if cp and is_garbage_contact(cp, row.get("Company") or ""):
            row["Contact Person"] = ""
            cname = row.get("Company") or ""
            if cname in ckpt:
                del ckpt[cname]
    empty_contact_before = sum(1 for r in kept if _empty(r.get("Contact Person")))

    if fill_contacts and empty_contact_before:
        settings = Settings.load()
        router = FreeSearchRouter(settings)
        key_oai = (os.getenv("OPENAI_API_KEY") or "").strip()
        if not key_oai:
            print("  WARN: no OPENAI_API_KEY — skip contact fill", flush=True)
        else:
            client = OpenAI(
                api_key=key_oai,
                base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/"),
            )
            model = (os.getenv("OPENAI_MODEL") or "gpt-4o-mini").strip()
            todo_idx = [i for i, r in enumerate(kept) if _empty(r.get("Contact Person"))]
            if limit_contacts is not None:
                todo_idx = todo_idx[: max(0, limit_contacts)]
            print(f"  Contact fill queue: {len(todo_idx)} (of {empty_contact_before} empty)", flush=True)
            sem = asyncio.Semaphore(int(os.getenv("CONTACT_FILL_CONCURRENT") or "3"))
            apollo_client = None
            try:
                from vendor_intel.enrichment.apollo_contacts import ApolloContactClient

                use_apollo = os.getenv("CONTACT_USE_APOLLO", "").lower() in ("1", "true", "yes")
                if use_apollo and (os.getenv("APOLLO_API_KEY") or "").strip():
                    apollo_client = ApolloContactClient(verbose=False)
                    await apollo_client.__aenter__()
            except Exception as exc:
                print(f"  Apollo unavailable: {exc}", flush=True)
                apollo_client = None

            async def one(i: int) -> None:
                nonlocal filled_n
                row = kept[i]
                company = row.get("Company") or ""
                async with sem:
                    print(f"  [contact] {company[:60].encode('ascii', 'replace').decode('ascii')}", flush=True)
                    hit = await resolve_contact(
                        company=company,
                        website=row.get("Website") or "",
                        market_query=m["query"],
                        router=router,
                        client=client,
                        model=model,
                        apollo=apollo_client,
                    )
                    if hit.get("Contact Person"):
                        row["Contact Person"] = hit["Contact Person"]
                        if hit.get("Role") and _empty(row.get("Role")):
                            row["Role"] = hit["Role"]
                        if hit.get("Email") and _empty(row.get("Email")):
                            row["Email"] = hit["Email"]
                        src_tag = hit.get("_via") or "contact_web_fill"
                        src = row.get("Data Sources") or ""
                        tag = f"contact_{src_tag}"
                        if tag not in src:
                            row["Data Sources"] = (src + f"; {tag}").strip("; ")
                        ckpt[company] = {
                            "Contact Person": row["Contact Person"],
                            "Role": row.get("Role") or "",
                            "Email": row.get("Email") or "",
                        }
                        filled_n += 1
                        ckpt_path.write_text(
                            json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8"
                        )

            try:
                await asyncio.gather(*[one(i) for i in todo_idx])
            finally:
                ckpt_path.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
                if apollo_client is not None:
                    await apollo_client.__aexit__(None, None, None)

    empty_contact_after = sum(1 for r in kept if _empty(r.get("Contact Person")))
    empty_founded = sum(1 for r in kept if _empty(r.get("Founded")))
    empty_hq = sum(1 for r in kept if _empty(r.get("Headquarters")))

    # backup + write cleaned
    bak = path.with_suffix(path.suffix + f".bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    if not path.with_name(path.stem + "_CLEANED.xlsx").exists():
        pass
    shutil.copy2(path, bak)
    out = path.with_name(path.stem + "_CLEANED.xlsx")
    audit = {
        "market": m["query"],
        "source_file": str(path),
        "backup": str(bak),
        "before_count": before,
        "after_count": len(kept),
        "removed_count": len(removed),
        "removed_samples": [
            {"company": r.get("Company"), "reason": r.get("_remove_reason")} for r in removed[:40]
        ],
        "empty_contact_before": empty_contact_before,
        "contacts_filled_this_run": filled_n,
        "empty_contact_after": empty_contact_after,
        "empty_founded": empty_founded,
        "empty_hq": empty_hq,
        "business_types": Counter(r.get("Business Type") or "" for r in kept).most_common(15),
        "cleaned_at": datetime.now(timezone.utc).isoformat(),
    }
    banner = f"{m['query']} — Channel partners / distributors only (cleaned)"
    write_company_sheet(
        out,
        sheet_name=m["sheet"],
        headers=[h for h in headers if h],
        rows=kept,
        audit=audit,
        title_banner=banner,
    )
    print(
        f"  before={before} after={len(kept)} removed={len(removed)} "
        f"contacts_filled=+{filled_n} empty_contact={empty_contact_after} "
        f"-> {out.name}",
        flush=True,
    )
    return audit


async def main_async() -> None:
    fill = os.getenv("SKIP_CONTACT_FILL", "").lower() not in ("1", "true", "yes")
    limit = os.getenv("CONTACT_FILL_LIMIT")
    limit_n = int(limit) if limit and limit.isdigit() else None
    # Default: fill all empties (can be slow). Set CONTACT_FILL_LIMIT for smoke.
    summary = []
    for m in MARKETS:
        # Avocado/smartwatch already have contacts — still run OEM filter
        do_fill = fill and m["key"] in ("green", "firewall", "smartphone")
        audit = await process_market(m, fill_contacts=do_fill, limit_contacts=limit_n)
        summary.append(audit)

    out = ROOT / "output" / "five_markets_clean_summary.json"
    out.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\n===== TOTALS =====")
    total_before = sum(a["before_count"] for a in summary)
    total_after = sum(a["after_count"] for a in summary)
    for a in summary:
        print(
            f"{a['market']}: {a['before_count']} -> {a['after_count']} "
            f"(removed {a['removed_count']}, contacts +{a['contacts_filled_this_run']}, "
            f"empty contact left {a['empty_contact_after']})"
        )
    print(f"ALL: {total_before} -> {total_after}")
    print(f"Summary: {out}")


if __name__ == "__main__":
    asyncio.run(main_async())
