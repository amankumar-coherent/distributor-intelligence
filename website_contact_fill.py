#!/usr/bin/env python3
"""Fill blank contacts by fetching each company's OWN website.

A third source, independent of the two already tried: AI Mode (what the
model recalls) and Apollo (what a contacts database indexed). This reads
the company's actual pages -- /leadership, /about-us, /team, /management,
/our-team and the sales/procurement variants -- so it can find people who
were never in either dataset. For small family-run Asian trading firms,
which is most of what remains blank, the company's own "About Us" page is
often the ONLY place a staff name is published.

Reuses enrichment/gap_fill/contact_page.py rather than re-implementing
fetching and extraction: that module already handles path probing,
encoding, an LLM extraction pass with a regex fallback, and it records
the source URL a name came from.

The user's role priority still applies. contact_page's own matcher is
CEO-oriented (it was written for leadership discovery), so this adds the
commercial-role paths and prefers a procurement/sales title when the page
yields more than one person.

Only blank rows are touched. Saves incrementally -- a long run WILL get
interrupted, and an end-only save loses everything when it does.

Usage:
    python website_contact_fill.py --file "<path.xlsx>" --limit 10 --dry-run
    python website_contact_fill.py --file "<path.xlsx>"
"""
from __future__ import annotations

import argparse
import asyncio
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from openpyxl import load_workbook  # noqa: E402

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

apply_env_overrides()

from vendor_intel.enrichment.gap_fill import contact_page  # noqa: E402

# The user's first-priority roles. Checked against whatever title the
# page yields, to prefer a buyer over a founder when both appear.
PRIORITY_TITLE = re.compile(
    r"\b(procurement|purchasing|buyer|sourcing|category manager|"
    r"sales|commercial|business development|supply chain)\b", re.I)

# contact_page probes leadership paths only; these are where a
# distributor actually names its commercial staff.
EXTRA_PATHS = (
    "/sales", "/our-people", "/people", "/staff", "/management-team",
    "/en/team", "/en/leadership", "/company/about", "/aboutus",
    "/corporate/about", "/en/company",
)


def latin_only(text: str) -> bool:
    """The deliverable is English/Latin-script names only."""
    return not re.search(r"[^\x00-\x7F]", text or "")


async def one(company: str, website: str,
              budget: float = 20.0) -> tuple[str, str, str]:
    """(person, role, source_url) or ("","","").

    `budget` caps the whole company. Without it a single unresponsive
    host stalls its entire concurrent batch, because a batch only
    reports when its slowest member returns -- observed live as 10
    companies producing no output at all for four minutes. A site that
    cannot answer in 20 seconds is not going to yield a contact worth
    holding nine other companies for.
    """
    row = {"company": company, "website": website, "domain": website}
    try:
        got = await asyncio.wait_for(
            contact_page.extract_leadership_from_website(row), timeout=budget)
    except asyncio.TimeoutError:
        return "", "", ""
    except Exception:  # noqa: BLE001  - a dead site must not kill the run
        return "", "", ""
    person = (got.get("contact_person") or "").strip()
    role = (got.get("contact_role") or "").strip()
    src = (got.get("leadership_source_url") or "").strip()
    if not person or not latin_only(person):
        return "", "", ""
    return person, role or "Contact", src


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--file", required=True)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--concurrency", type=int, default=5,
                    help="companies fetched at once (default 5)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    # Widen the probe set before any fetching happens.
    contact_page.LEADERSHIP_PATHS = tuple(
        dict.fromkeys(contact_page.LEADERSHIP_PATHS + EXTRA_PATHS))

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
            continue
        todo.append((r, str(ws.cell(r, cN).value).strip(),
                     str(ws.cell(r, cW).value or "").strip()))
    if args.limit:
        todo = todo[:args.limit]

    print(f"{path.name}: {len(todo)} companies still blank "
          f"(fetching {args.concurrency} at a time)")

    async def batch(items):
        """Fetch a group of companies concurrently.

        Each company costs ~3 minutes almost entirely in network waits --
        probing candidate paths and one extraction call. Those waits
        overlap, so running a group together cuts wall-clock roughly by
        the group size without adding load per host (each company is a
        different domain).
        """
        return await asyncio.gather(
            *(one(name, site) for _row, name, site in items),
            return_exceptions=True)

    found = 0
    done = 0
    for start in range(0, len(todo), args.concurrency):
        group = [t for t in todo[start:start + args.concurrency] if t[2]]
        if not group:
            continue
        results = asyncio.run(batch(group))
        for (row, name, _site), res in zip(group, results):
            done += 1
            if isinstance(res, BaseException) or not res[0]:
                print(f"  [{done:>3}/{len(todo)}] {name[:34]:36} -",
                      flush=True)
                continue
            person, role, _src = res
            tier = "priority" if PRIORITY_TITLE.search(role) else "other"
            print(f"  [{done:>3}/{len(todo)}] {name[:34]:36}{person[:20]:22}"
                  f"{role[:28]:30}{tier}", flush=True)
            found += 1
            if not args.dry_run:
                ws.cell(row, cC).value = person
                ws.cell(row, cR).value = f"{role} (website)"
        # Save after every group, not every 5 finds: a run this long will
        # be interrupted, and Apollo already lost 4 results to a
        # find-counted save that never reached its threshold.
        if not args.dry_run and found:
            wb.save(path)

    print("")
    print(f"  website contacts found: {found} of {len(todo)}")
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
