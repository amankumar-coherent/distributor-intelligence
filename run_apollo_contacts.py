#!/usr/bin/env python3
"""Find procurement/sales contacts via Apollo, using only free search.

Used as the fallback where SalesQL has no record. Apollo's free
`mixed_people/api_search` returns title, organisation and a name that is
often complete ("Tommy Ng", "Rachel Li") and otherwise partially masked
("Soon Hu***y"). Both are recorded; the mask is preserved exactly as
returned rather than guessed at, and a Name Complete column says which is
which.

NO CREDITS: only the search endpoint is called. Apollo charges for
`people/match` enrichment, which reveals the full surname and email —
that is deliberately never called here.

Note this key cannot reach people/search or organizations/search (403);
mixed_people/api_search is the one available endpoint.
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

APOLLO_SEARCH = "https://api.apollo.io/api/v1/mixed_people/api_search"

TITLES = [
    "Procurement Manager", "Head of Procurement", "Procurement Director",
    "Chief Procurement Officer", "Purchasing Manager", "Head of Purchasing",
    "Purchasing Director", "Sourcing Manager", "Head of Sourcing",
    "Category Manager", "Buyer", "Senior Buyer", "Supply Chain Manager",
    "Head of Supply Chain", "VP of Sales", "VP Sales", "Head of Sales",
    "Sales Director", "Director of Sales", "National Sales Manager",
    "Regional Sales Manager", "Sales Manager", "Chief Revenue Officer",
    "Chief Commercial Officer", "Business Development Manager",
    "Head of Business Development", "Commercial Director", "Country Manager",
    "General Manager",
]

BUCKETS = [
    ("PROCUREMENT", r"\b(procurement|purchasing|sourcing|buyer|category manager|supply chain)\b"),
    ("VP OF SALES", r"\b(vp of sales|vp sales|vice president of sales)\b"),
    ("HEAD OF SALES", r"\b(head of sales|sales director|director of sales|national sales manager|chief revenue officer)\b"),
    ("SALES", r"\b(sales manager|regional sales|sales lead|account director|sales)\b"),
    ("BUSINESS DEVELOPMENT", r"\b(business development|commercial director|chief commercial officer)\b"),
    ("COUNTRY MANAGER", r"\b(country manager|country head|general manager)\b"),
]

PRIORITY = ["PROCUREMENT", "HEAD OF SALES", "VP OF SALES", "SALES",
            "BUSINESS DEVELOPMENT", "COUNTRY MANAGER"]


def classify(title: str) -> list[str]:
    low = (title or "").lower()
    return [b for b, pat in BUCKETS if re.search(pat, low)]


def clean_domain(website: str) -> str:
    d = (website or "").strip().lower()
    d = re.sub(r"^https?://", "", d).strip("/")
    d = d.split("/")[0]
    return d if "." in d else ""


def person_name(p: dict) -> tuple[str, bool]:
    """(display name, is_complete). Apollo masks some surnames as 'Hu***y'."""
    first = (p.get("first_name") or "").strip()
    last = (p.get("last_name") or "").strip()
    masked = (p.get("last_name_obfuscated") or "").strip()
    if last:
        return f"{first} {last}".strip(), True
    if masked:
        complete = "*" not in masked
        return f"{first} {masked}".strip(), complete
    return first, False


def search(client: httpx.Client, key: str, domain: str, per_page: int = 25) -> list[dict]:
    resp = client.post(
        APOLLO_SEARCH,
        headers={"x-api-key": key, "Content-Type": "application/json", "Cache-Control": "no-cache"},
        json={
            "q_organization_domains_list": [domain],
            "person_titles": TITLES,
            "page": 1,
            "per_page": per_page,
        },
        timeout=40.0,
    )
    if resp.status_code >= 400:
        return []
    data = resp.json() or {}
    return data.get("people") or data.get("contacts") or []


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies-csv", required=True, help="CSV with Company and Website columns")
    parser.add_argument("--skip-json", default=None, help="SalesQL coverage json; companies already solved are skipped")
    parser.add_argument("--checkpoint", default=str(ROOT / "output/apollo/apollo_contacts.json"))
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--delay", type=float, default=1.2)
    args = parser.parse_args()

    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
    key = (os.getenv("APOLLO_API_KEY") or "").strip()
    if not key:
        raise SystemExit("Set APOLLO_API_KEY in .env")

    import pandas as pd

    df = pd.read_csv(args.companies_csv)
    rows = [
        (str(r["Company"]).strip(), clean_domain(str(r.get("Website") or "")))
        for _, r in df.iterrows()
        if str(r["Company"]).strip() and str(r["Company"]).strip().lower() != "nan"
    ]

    solved: set[str] = set()
    if args.skip_json and Path(args.skip_json).exists():
        cov = json.loads(Path(args.skip_json).read_text(encoding="utf-8"))
        solved = {k.lower() for k, v in cov.items() if v.get("roles")}

    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    todo = [
        (name, dom) for name, dom in rows
        if dom and name.lower() not in solved and name.lower() not in data
    ]
    if args.limit:
        todo = todo[: args.limit]

    print(f"companies in file      : {len(rows)}")
    print(f"already solved by SalesQL: {len(solved)}")
    print(f"to query via Apollo    : {len(todo)}")

    found = 0
    with httpx.Client() as client:
        for i, (name, domain) in enumerate(todo, start=1):
            rec: dict = {"domain": domain, "roles": {}}
            try:
                for p in search(client, key, domain):
                    display, complete = person_name(p)
                    if not display:
                        continue
                    title = (p.get("title") or "").strip()
                    for bucket in classify(title):
                        cur = rec["roles"].get(bucket)
                        # Prefer a complete name over a masked one.
                        if cur is None or (complete and not cur.get("complete")):
                            rec["roles"][bucket] = {
                                "name": display,
                                "title": title,
                                "complete": complete,
                                "org": (p.get("organization") or {}).get("name") or "",
                            }
            except Exception as exc:  # noqa: BLE001
                rec["error"] = type(exc).__name__
            data[name.lower()] = rec
            if rec["roles"]:
                found += 1
            if i % 20 == 0 or i == len(todo):
                ckpt.parent.mkdir(parents=True, exist_ok=True)
                ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
                print(f"  [{i}/{len(todo)}] with contacts: {found}", flush=True)
            time.sleep(args.delay)

    ckpt.parent.mkdir(parents=True, exist_ok=True)
    ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    total_roles = sum(len(v.get("roles") or {}) for v in data.values())
    complete = sum(1 for v in data.values() for x in (v.get("roles") or {}).values() if x.get("complete"))
    print(f"\ncompanies with contacts : {found}")
    print(f"contacts found          : {total_roles}   (full name {complete}, partial {total_roles - complete})")
    print(f"checkpoint: {ckpt}")


if __name__ == "__main__":
    main()
