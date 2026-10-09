#!/usr/bin/env python3
"""Reveal ONE verified business email per leadership contact and export to Excel.

Selection rule (deliberately strict, because credits are spent per reveal):
  1. type == "BE" (business) AND verified == true
  2. among those, prefer the address on the company's own domain — a
     contact often carries business emails from previous employers, and
     those are the wrong address for this company
  3. if no verified business email exists, the person is skipped rather
     than falling back to a personal/gmail address

Search is free and already reports each email's type/verified/domain, so
the choice is made before any credit is spent. Only the chosen address is
revealed via enrich_person.

Usage:
    python run_salesql_reveal.py --coverage <coverage.json> --dry-run
    python run_salesql_reveal.py --coverage <coverage.json>
"""
import argparse
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

import httpx

from run_salesql_ceo_cto import ROLES, SEARCH_TITLES, credits, mcp_call, pick

OUT_DIR = ROOT / "output" / "salesql"


# Domains that belong to a marketplace, directory or mailbox provider
# rather than to the company itself. SalesQL stores a company's
# "website_domain" from whatever site is on file, and for smaller firms
# that is often an IndiaMART/JustDial listing — so APS Agro came back with
# website_domain "indiamart.com" and the domain-preference below then
# picked an IndiaMART employee as its founder.
NON_COMPANY_DOMAINS = {
    "indiamart.com", "justdial.com", "tradeindia.com", "exportersindia.com",
    "alibaba.com", "made-in-china.com", "linkedin.com", "facebook.com",
    "gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "rediffmail.com",
    "bit.ly", "lnkd.in", "tinyurl.com", "wordpress.com", "wixsite.com",
    "blogspot.com", "sites.google.com", "godaddysites.com",
}


def company_domain(org: dict) -> str:
    domain = (org.get("website_domain") or "").lower().strip()
    if not domain or domain in NON_COMPANY_DOMAINS:
        return ""
    return domain


def _stem(text: str) -> str:
    """Comparable core of a name or domain: lowercase alphanumerics only."""
    import re as _re

    return _re.sub(r"[^a-z0-9]", "", (text or "").lower())


def domain_matches_company(email_domain: str, domain: str, company: str, matched: str) -> bool:
    """True when an email's domain plausibly belongs to this company.

    Accepts an exact/suffix match on the company's own website domain, and
    otherwise falls back to comparing the domain's second-level label with
    the company name ("atirobotics.ai" for "Ati Motors" fails this, which
    is the intended strictness — see the caller, which records such a
    mismatch rather than silently trusting it).
    """
    email_domain = (email_domain or "").lower()
    if not email_domain or email_domain in NON_COMPANY_DOMAINS:
        return False
    if domain and (email_domain == domain or email_domain.endswith("." + domain)):
        return True
    label = _stem(email_domain.split(".")[0])
    if not label:
        return False
    for name in (company, matched):
        stem = _stem(name)
        if not stem:
            continue
        if label in stem or stem in label:
            return True
        if len(label) >= 5 and len(stem) >= 5 and (label[:5] == stem[:5]):
            return True
    return False


def choose_email(
    emails: list[dict], domain: str, company: str = "", matched: str = ""
) -> tuple[dict | None, str]:
    """Pick the single verified business email to reveal.

    Returns (entry, status). Only an address whose domain plausibly belongs
    to this company is returned: a person's record also carries verified
    business emails from PREVIOUS employers, and blindly taking the first
    one attributes a competitor's address to this company. Measured on a
    312-row export, that fallback produced 76 wrong-domain emails.
    """
    business = [
        e for e in (emails or [])
        if (e.get("type") or "").upper() == "BE" and e.get("verified") is True
    ]
    if not business:
        return None, "no verified business email"

    for entry in business:
        email_domain = (entry.get("email") or "").split("@")[-1].lower()
        if domain_matches_company(email_domain, domain, company, matched):
            return entry, "verified business"

    return None, "only other-employer emails"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--coverage", required=True, help="coverage json listing companies with roles")
    parser.add_argument("--output", default=str(OUT_DIR / "leadership_verified_emails.xlsx"))
    parser.add_argument("--checkpoint", default=str(OUT_DIR / "reveal_checkpoint.json"))
    parser.add_argument("--dry-run", action="store_true", help="Show what would be revealed; spend nothing")
    args = parser.parse_args()

    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
    key = (os.getenv("SALESQL_API_KEY") or "").strip()
    if not key:
        raise SystemExit("Set SALESQL_API_KEY in .env")

    coverage = json.loads(Path(args.coverage).read_text(encoding="utf-8"))
    companies = [c for c, v in coverage.items() if v.get("roles")]

    ckpt = Path(args.checkpoint)
    done = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    rows: list[dict] = []
    planned = revealed = skipped = 0

    with httpx.Client() as client:
        before = credits(client, key)
        print(f"credits before : {before}")
        print(f"companies      : {len(companies)}")

        for i, company in enumerate(companies, start=1):
            cached = done.get(company.lower())
            if cached and not args.dry_run:
                rows.extend(cached)
                continue

            # One company failing (network fault, odd response) must not end
            # the whole run: the checkpoint keeps what succeeded, and a
            # resume retries only what is missing.
            try:
                org = mcp_call(client, key, "enrich_organization", {"organization_name": company})
                org_obj = (org or {}).get("organization") or {}
                uuid = org_obj.get("uuid")
                if not uuid:
                    continue
                domain = company_domain(org_obj)
                res = mcp_call(
                    client, key, "search_people",
                    {"organization_uuids": [uuid], "job_titles": SEARCH_TITLES, "page_size": 25},
                )
                chosen = pick(res.get("results") or [])
            except Exception as exc:  # noqa: BLE001
                print(f"  ! {company[:40]}: {type(exc).__name__} — will retry on resume", flush=True)
                continue

            company_rows: list[dict] = []
            for role, person in chosen.items():
                entry, reason = choose_email(
                    person.get("emails") or [],
                    domain,
                    company,
                    org_obj.get("name") or "",
                )
                row = {
                    "Company": company,
                    "Matched Company": org_obj.get("name") or "",
                    "Role": role,
                    "Name": person.get("full_name") or "",
                    "Title": person.get("title") or "",
                    "Email": "",
                    "Email Status": "",
                    "LinkedIn": person.get("linkedin_url") or "",
                }
                if entry is None:
                    row["Email Status"] = reason
                    skipped += 1
                    company_rows.append(row)
                    continue

                planned += 1
                if args.dry_run:
                    row["Email"] = entry.get("email", "")  # still masked
                    row["Email Status"] = "would reveal (verified business)"
                    company_rows.append(row)
                    continue

                try:
                    detail = mcp_call(
                        client, key, "enrich_person",
                        {"linkedin_url": person.get("linkedin_url")} if person.get("linkedin_url")
                        else {"full_name": person.get("full_name"), "organization_name": company},
                    )
                except Exception as exc:  # noqa: BLE001
                    row["Email Status"] = f"reveal failed: {type(exc).__name__}"
                    company_rows.append(row)
                    continue
                real = ((detail or {}).get("person") or {}).get("emails") or []
                masked_domain = (entry.get("email") or "").split("@")[-1].lower()
                match = next(
                    (
                        e.get("email")
                        for e in real
                        if e.get("email")
                        and "..." not in e.get("email")
                        and e.get("email", "").lower().endswith("@" + masked_domain)
                    ),
                    "",
                )
                if match:
                    row["Email"] = match
                    row["Email Status"] = "verified business"
                    revealed += 1
                else:
                    row["Email Status"] = "reveal returned nothing"
                company_rows.append(row)

            rows.extend(company_rows)
            if not args.dry_run:
                done[company.lower()] = company_rows
                ckpt.parent.mkdir(parents=True, exist_ok=True)
                ckpt.write_text(json.dumps(done, indent=1, ensure_ascii=False), encoding="utf-8")
            if i % 5 == 0 or i == len(companies):
                print(f"  [{i}/{len(companies)}] revealed {revealed}  planned {planned}", flush=True)
            time.sleep(0.3)

        after = credits(client, key)

    if args.dry_run:
        print(f"\nDRY RUN — nothing spent")
        print(f"  people with a verified business email : {planned}")
        print(f"  people without one (skipped)          : {skipped}")
        print(f"  estimated reveals                      : {planned}")
        print(f"credits unchanged: {before} -> {after}")
        return

    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    df = pd.DataFrame(rows, columns=[
        "Company", "Matched Company", "Role", "Name", "Title", "Email", "Email Status", "LinkedIn",
    ])
    df = df.sort_values(["Company", "Role"], kind="stable")
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Leadership Emails")
        ws = writer.sheets["Leadership Emails"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            ws.column_dimensions[get_column_letter(idx)].width = min(
                max(typical + 2, len(str(column)) + 2, 10), 40
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24

    with_email = int((df["Email"] != "").sum())
    print(f"\nrows written  : {len(df)}")
    print(f"with email    : {with_email}")
    print(f"revealed      : {revealed}   skipped (no verified BE): {skipped}")
    print(f"credits       : {before} -> {after}")
    print(f"Excel         : {output}")


if __name__ == "__main__":
    main()
