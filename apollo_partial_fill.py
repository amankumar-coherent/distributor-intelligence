#!/usr/bin/env python3
"""Fill remaining blank contacts from Apollo's FREE search — partial names.

Apollo's mixed_people/api_search costs no credits, but it paywalls the
full name: `name`, `last_name`, `linkedin_url` and `email` all come back
null. What it DOES return free is:

    first_name            "Michael"
    last_name_obfuscated  "Sin"
    title                 "Sales (F.S)"

so a usable partial -- "Michael S." -- plus a real job title. The user
asked for partial names explicitly, having no credits.

Role priority matches the AI Mode pass: procurement and sales leadership
first, C-suite only as a fallback. Apollo is queried with the priority
titles first and the fallback titles only if nothing comes back.

Rows already carrying a contact are never touched, so this only ever adds
to the 71 the two AI Mode passes could not name.

Every value written is marked "(Apollo, partial)" in the Role cell so a
partial name is never mistaken for a confirmed full contact.

Usage:
    python apollo_partial_fill.py --file "<path.xlsx>" --dry-run
    python apollo_partial_fill.py --file "<path.xlsx>"
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

APOLLO_URL = "https://api.apollo.io/api/v1/mixed_people/api_search"

# The user's priority. Apollo matches titles loosely, so these cover the
# commercial/buying roles that actually transact with a distributor.
PRIORITY_TITLES = [
    "Procurement Manager", "Purchasing Manager", "Head of Procurement",
    "Procurement Director", "Purchasing Director", "Buyer", "Category Manager",
    "VP of Sales", "Head of Sales", "Sales Director", "Sales Manager",
    "Business Development Manager", "Commercial Director",
]
FALLBACK_TITLES = [
    "CEO", "Chief Executive Officer", "Managing Director", "President",
    "Founder", "COO", "CTO", "CFO", "General Manager", "Director",
]


def api_key() -> str:
    for line in (ROOT / ".env").read_text(encoding="utf-8").splitlines():
        if line.startswith("APOLLO_API_KEY="):
            k = line.split("=", 1)[1].strip()
            if k:
                return k
    raise SystemExit("APOLLO_API_KEY not set in .env")


def domain_of(website: str) -> str:
    w = re.sub(r"^https?://", "", str(website or "").strip()).strip("/")
    w = re.sub(r"^www\.", "", w)
    return w.split("/")[0].lower()


def search(key: str, domain: str, titles: list[str]) -> list[dict]:
    payload = {"q_organization_domains": domain, "page": 1, "per_page": 10}
    if titles:
        payload["person_titles"] = titles
    req = urllib.request.Request(
        APOLLO_URL, data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json",
                 "Cache-Control": "no-cache", "x-api-key": key})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read()).get("people") or []
    except urllib.error.HTTPError as e:
        if e.code == 429:
            time.sleep(10)
        return []
    except Exception:  # noqa: BLE001
        return []


def partial_name(person: dict) -> str:
    """"Michael S." from the free fields; "" if not even that is present."""
    first = str(person.get("first_name") or "").strip()
    if not first:
        return ""
    last = str(person.get("last_name_obfuscated") or "").strip()
    if last:
        # Apollo returns the real surname here for some records and an
        # initial for others; normalise to an initial either way so the
        # cell never looks like a confirmed full name.
        return f"{first} {last[0].upper()}."
    return first


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--file", required=True)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    key = api_key()
    path = Path(args.file)
    wb = load_workbook(path)
    ws = wb.active
    hrow = next(r for r in range(1, 15) if ws.cell(r, 1).value == "#")
    hdr = [ws.cell(hrow, c).value for c in range(1, ws.max_column + 1)]
    cC = hdr.index("Contact Person") + 1
    cR = hdr.index("Role") + 1
    cW = hdr.index("Website") + 1
    cN = hdr.index("Company") + 1

    todo = []
    for r in range(hrow + 1, ws.max_row + 1):
        if not ws.cell(r, cN).value:
            continue
        if str(ws.cell(r, cC).value or "").strip():
            continue                      # already has a contact
        todo.append((r, str(ws.cell(r, cN).value).strip(),
                     domain_of(ws.cell(r, cW).value)))
    if args.limit:
        todo = todo[:args.limit]

    print(f"{path.name}: {len(todo)} companies still blank")
    found = 0
    for i, (row, name, dom) in enumerate(todo, 1):
        if not dom:
            continue
        people = search(key, dom, PRIORITY_TITLES)
        tier = "priority"
        if not people:
            people = search(key, dom, FALLBACK_TITLES)
            tier = "c-suite"
        if not people:
            people = search(key, dom, [])
            tier = "any"

        pick = next((p for p in people if partial_name(p)), None)
        if not pick:
            print(f"  [{i:>3}/{len(todo)}] {name[:34]:36} -")
            time.sleep(1)
            continue

        who = partial_name(pick)
        title = str(pick.get("title") or "").strip() or "Contact"
        role = f"{title} (Apollo, partial)"
        print(f"  [{i:>3}/{len(todo)}] {name[:34]:36}{who[:20]:22}"
              f"{title[:30]:32}{tier}")
        found += 1
        if not args.dry_run:
            ws.cell(row, cC).value = who
            ws.cell(row, cR).value = role
            # Save as we go. A 71-company run takes long enough that an
            # interruption is likely, and an end-only save throws away
            # every result when that happens -- which it already did once.
            if found % 5 == 0:
                wb.save(path)
        time.sleep(1)

    print("")
    print(f"  partial contacts found: {found} of {len(todo)}")
    if args.dry_run:
        print("  dry run - nothing written")
        return
    if found:
        backup = path.with_suffix(f".{datetime.now():%Y%m%d-%H%M%S}.bak.xlsx")
        shutil.copy2(path, backup)
        wb.save(path)
        print(f"  saved (backup: {backup.name})")


if __name__ == "__main__":
    main()
