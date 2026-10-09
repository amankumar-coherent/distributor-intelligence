#!/usr/bin/env python3
"""Find each company's CEO and CTO (and their emails) via the SalesQL MCP server.

Why MCP and not the REST API: api-public.salesql.com exposes enrichment
only — every person lookup needs a LinkedIn URL, an email, or a name, so it
cannot answer "who is the CEO of X". The MCP server at mcp.salesql.com adds
a `search_people` tool that filters by organization + job title, which is
exactly the missing capability. It authenticates with the same API key over
plain HTTP (no OAuth needed).

Two phases, deliberately separated so spending is a decision and not a
side effect:

  discover (FREE)  enrich_organization -> org uuid
                   search_people       -> CEO/CTO names, titles, LinkedIn,
                                          and MASKED emails ("...@acme.com")
  reveal (CREDITS) enrich_person       -> the full email address

Measured on a live call: both discover steps left the credit balance
unchanged, and search returns emails masked, so you can survey every
company for free and then reveal only where it is worth a credit.

Usage:
    python run_salesql_ceo_cto.py --companies "New company list.xlsx" --limit 20
    python run_salesql_ceo_cto.py --companies "New company list.xlsx"
    python run_salesql_ceo_cto.py --companies "New company list.xlsx" --reveal-emails
"""
import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

import httpx

from run_leadership_aimode_batch import _load_company_list

MCP_URL = "https://mcp.salesql.com/mcp"
ALLOWANCE_URL = "https://api-public.salesql.com/v1/allowance"
OUT_DIR = ROOT / "output" / "salesql"
DEFAULT_CHECKPOINT = OUT_DIR / "salesql_ceo_cto.json"
DEFAULT_OUTPUT = OUT_DIR / "ceo_cto_salesql.xlsx"

ROLES = [
    "MD", "CEO", "COO", "CFO", "COUNTRY HEAD", "CTO",
    "AI HEAD", "MARKETING HEAD", "Founder", "OTHER CXO",
]

# Titles sent to search_people. search is free, so casting a wide net costs
# nothing and catches variant wording ("Chief Technical Officer", "CMO").
SEARCH_TITLES = [
    "CEO", "Chief Executive Officer", "Managing Director", "MD",
    "Founder", "Co-founder", "Founder & CEO", "Co-founder & CEO",
    "CTO", "Chief Technology Officer", "Chief Technical Officer",
    "Co-founder & CTO", "VP Engineering", "Head of Engineering",
    "CIO", "Chief Information Officer",
    "COO", "Chief Operating Officer", "Head of Operations", "VP Operations",
    "CFO", "Chief Financial Officer", "Finance Director", "Head of Finance",
    "CHRO", "Chief Human Resources Officer", "Chief People Officer",
    "CRO", "Chief Revenue Officer", "CPO", "Chief Product Officer",
    "CSO", "Chief Strategy Officer", "Chief Security Officer",
    "CCO", "Chief Commercial Officer", "Chief Customer Officer",
    "CDO", "Chief Data Officer", "Chief Digital Officer",
    "Country Head", "Country Manager", "India Head", "Regional Director",
    "General Manager", "Head of AI", "AI Head", "Chief AI Officer",
    "Head of Data Science", "Head of Machine Learning",
    "CMO", "Chief Marketing Officer", "Marketing Head",
    "Head of Marketing", "VP Marketing", "Marketing Director",
]

# Ordered most-specific first: a title like "Co-founder & CTO" must land in
# CTO rather than Founder, so CTO/AI/marketing patterns are tested before
# the broader founder and CEO ones.
ROLE_PATTERNS = [
    ("AI HEAD", r"\b(chief ai officer|head of (ai|artificial intelligence|data science|machine learning)|ai head|vp,? (ai|data science))\b"),
    ("CTO", r"\b(cto|chief technology officer|chief technical officer|chief information officer|cio|head of engineering|vp,? engineering)\b"),
    ("CFO", r"\b(cfo|chief financial officer|finance director|head of finance|vp,? finance)\b"),
    ("COO", r"\b(coo|chief operating officer|head of operations|vp,? operations)\b"),
    ("MARKETING HEAD", r"\b(cmo|chief marketing officer|marketing head|head of marketing|vp,? marketing|marketing director)\b"),
    ("COUNTRY HEAD", r"\b(country head|country manager|country director|india head|head of india|regional director|general manager)\b"),
    ("MD", r"\b(managing director|\bmd\b)\b"),
    ("CEO", r"\b(ceo|chief executive officer)\b"),
    ("Founder", r"\b(founder|co-?founder)\b"),
    # Any remaining chief-officer title. Listed last so a specific role
    # above claims it first; this only catches CHRO/CRO/CPO/CSO/CCO/CDO and
    # similar, which are worth keeping rather than discarding.
    ("OTHER CXO", r"\b(chro|cro|cpo|cso|cco|cdo|chief (human resources|people|revenue|product|strategy|security|commercial|customer|data|digital|business|growth|compliance|legal|medical|scientific) officer)\b"),
]


def api_key() -> str:
    key = (os.getenv("SALESQL_API_KEY") or "").strip()
    if not key:
        raise SystemExit("Set SALESQL_API_KEY in .env")
    return key


def mcp_call(client: httpx.Client, key: str, tool: str, params: dict) -> dict:
    """One MCP tools/call. Returns the tool's decoded JSON payload."""
    body = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {"name": tool, "arguments": {"params": params}},
    }
    # Retry transient network faults. Measured: a long run died at company
    # 82 of 209 with "connection forcibly closed by the remote host" — the
    # request had already been paid for in wall-clock time and the whole
    # process exited, so a bare call is too fragile for a batch this size.
    last_error: Exception | None = None
    for attempt in range(4):
        try:
            resp = client.post(
                MCP_URL,
                json=body,
                headers={
                    "Authorization": f"Bearer {key}",
                    "Accept": "application/json, text/event-stream",
                    "Content-Type": "application/json",
                },
                timeout=60.0,
            )
            resp.raise_for_status()
            break
        except (httpx.HTTPError, OSError) as exc:
            last_error = exc
            time.sleep(1.5 * (attempt + 1))
    else:
        raise last_error if last_error else RuntimeError("MCP call failed")
    # The server replies as SSE ("event: message\ndata: {...}").
    match = re.search(r"data: (\{.*)", resp.text, re.S)
    if not match:
        return {}
    envelope = json.loads(match.group(1))
    if "error" in envelope:
        return {"_error": envelope["error"]}
    content = (envelope.get("result") or {}).get("content") or []
    if not content:
        return {}
    try:
        return json.loads(content[0].get("text") or "{}")
    except (json.JSONDecodeError, ValueError):
        return {}


def credits(client: httpx.Client, key: str) -> dict:
    try:
        r = client.get(ALLOWANCE_URL, params={"api_key": key}, timeout=20.0)
        return (r.json() or {}).get("credits") or {}
    except Exception:  # noqa: BLE001
        return {}


def classify(title: str) -> list[str]:
    """Every role a job title satisfies.

    Returns a list because one person genuinely holds several of these:
    "Co-founder & CEO" is both Founder and CEO, and recording only the
    first would leave a column blank that the data can actually fill.
    """
    low = (title or "").lower()
    return [role for role, pattern in ROLE_PATTERNS if re.search(pattern, low)]


def pick(people: list[dict]) -> dict:
    """Best person per role, preferring whoever has more emails on file."""
    chosen: dict = {}
    for person in people:
        for role in classify(person.get("title") or person.get("headline") or ""):
            current = chosen.get(role)
            if current is None or len(person.get("emails") or []) > len(current.get("emails") or []):
                chosen[role] = person
    return chosen


def write_excel(data: dict, companies: list[str], output: Path) -> None:
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    rows = []
    for company in companies:
        entry = data.get(company.lower()) or {}
        row = {"Company": company, "Matched Company": entry.get("org_name", "")}
        for role in ROLES:
            person = entry.get(role) or {}
            row[role] = person.get("full_name", "")
            row[f"{role} TITLE"] = person.get("title", "")
            row[f"{role} EMAIL"] = person.get("email", "")
            row[f"{role} LINKEDIN"] = person.get("linkedin_url", "")
        row["Found"] = sum(1 for r in ROLES if (entry.get(r) or {}).get("full_name"))
        rows.append(row)

    df = pd.DataFrame(rows)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CEO and CTO")
        ws = writer.sheets["CEO and CTO"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            ws.column_dimensions[get_column_letter(idx)].width = min(
                max(typical + 2, len(str(column)) + 2, 10), 36
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies", required=True)
    parser.add_argument("--checkpoint", default=str(DEFAULT_CHECKPOINT))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--delay", type=float, default=0.4, help="Seconds between companies")
    parser.add_argument(
        "--reveal-emails",
        action="store_true",
        help="SPENDS CREDITS: call enrich_person to unmask each found email",
    )
    parser.add_argument("--retry-missing", action="store_true", help="Re-query companies with no CEO/CTO yet")
    args = parser.parse_args()

    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
    key = api_key()

    companies = _load_company_list(args.companies)
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    def has_any(name: str) -> bool:
        entry = data.get(name.lower()) or {}
        return any((entry.get(r) or {}).get("full_name") for r in ROLES)

    if args.retry_missing:
        todo = [c for c in companies if not has_any(c)]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    with httpx.Client() as client:
        before = credits(client, key)
        print(f"credits before: {before}")
        print(f"companies     : {len(companies)}   to process: {len(todo)}")
        if args.reveal_emails:
            print("MODE: reveal-emails ON — enrich_person will consume credits")
        if not todo:
            write_excel(data, companies, Path(args.output))
            print(f"nothing to do; Excel refreshed at {args.output}")
            return

        found_roles = found_companies = revealed = 0
        for i, company in enumerate(todo, start=1):
            entry: dict = {"org_name": ""}
            for role in ROLES:
                entry[role] = {}
            try:
                org = mcp_call(client, key, "enrich_organization", {"organization_name": company})
                org_obj = (org or {}).get("organization") or {}
                uuid = org_obj.get("uuid")
                entry["org_name"] = org_obj.get("name", "")
                if uuid:
                    res = mcp_call(
                        client,
                        key,
                        "search_people",
                        {
                            "organization_uuids": [uuid],
                            "job_titles": SEARCH_TITLES,
                            "page_size": 25,
                        },
                    )
                    people = (res or {}).get("results") or []
                    for role, person in pick(people).items():
                        emails = person.get("emails") or []
                        masked = emails[0].get("email", "") if emails else ""
                        entry[role] = {
                            "full_name": person.get("full_name", ""),
                            "title": person.get("title", ""),
                            "linkedin_url": person.get("linkedin_url", ""),
                            "email": masked,
                            "email_count": len(emails),
                        }
                        if args.reveal_emails and person.get("linkedin_url"):
                            det = mcp_call(
                                client,
                                key,
                                "enrich_person",
                                {"linkedin_url": person["linkedin_url"]},
                            )
                            real = ((det or {}).get("person") or {}).get("emails") or []
                            work = next(
                                (e.get("email") for e in real if e.get("type") == "BE" and e.get("email")),
                                next((e.get("email") for e in real if e.get("email")), ""),
                            )
                            if work and "..." not in work:
                                entry[role]["email"] = work
                                revealed += 1
            except Exception as exc:  # noqa: BLE001
                entry["error"] = f"{type(exc).__name__}: {exc}"

            data[company.lower()] = entry
            found_roles += sum(1 for r in ROLES if entry.get(r, {}).get("full_name"))
            if any(entry.get(r, {}).get("full_name") for r in ROLES):
                found_companies += 1

            if i % 10 == 0 or i == len(todo):
                ckpt.parent.mkdir(parents=True, exist_ok=True)
                ckpt.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
                print(
                    f"[{i}/{len(todo)}] companies with data {found_companies}  names {found_roles}"
                    + (f"  revealed {revealed}" if args.reveal_emails else ""),
                    flush=True,
                )
            if args.delay:
                time.sleep(args.delay)

        ckpt.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        after = credits(client, key)

    write_excel(data, companies, Path(args.output))
    total_ceo = sum(1 for c in companies if (data.get(c.lower(), {}).get("CEO") or {}).get("full_name"))
    total_cto = sum(1 for c in companies if (data.get(c.lower(), {}).get("CTO") or {}).get("full_name"))
    print(f"\nCEO found : {total_ceo} of {len(companies)}")
    print(f"CTO found : {total_cto} of {len(companies)}")
    print(f"credits   : before={before}  after={after}")
    print(f"Excel     : {args.output}")


if __name__ == "__main__":
    main()
