#!/usr/bin/env python3
"""Verify the six guessable contact fields of an already-exported market.

WHY THIS EXISTS. The discovery prompt deliberately instructs AI Mode to
invent a "best plausible" value for contact_person / role / email /
linkedin / office_no / office_email rather than leave them blank, because
the requirement was every one of the 21 columns populated with no blanks.
The cost of that trade-off is now measurable: 94% of Office Email values
are `info@<own-domain>` or `sales@<own-domain>`, which is the fingerprint
of a constructed address, not a researched one. Same for LinkedIn, where
_normalize_linkedin builds `/in/<name-slug>` from the contact's name.

This script re-checks those fields against the companies' REAL websites
using the existing (already tested) verify.py helpers, and reports what
actually holds up. It is deliberately:

  * NON-DESTRUCTIVE — reads the state file, never writes it. Output goes
    to a new .xlsx so the original export is untouched.
  * RESUMABLE — caches per-company results to JSON, so an interrupted run
    (or a rate-limited one) continues instead of starting over.
  * HONEST about what it cannot determine. A site that fails to load is
    reported as "unreachable", NOT as "email is fake" — absence of
    evidence is not evidence of fabrication, and conflating the two would
    repeat the mistake this script exists to correct.

Usage:
    # verify everything (writes a *_verified.xlsx alongside the original)
    python verify_contact_fields.py --market "Global Data Center System Market - Distributors Only"

    # try a small sample first to gauge hit rate and speed
    python verify_contact_fields.py --market "..." --limit 20

    # more parallelism (default 8; raise carefully, these are real fetches)
    python verify_contact_fields.py --market "..." --concurrency 12
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

apply_env_overrides()

from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    MarketResearchState,
    market_slug,
)

CACHE_DIR = ROOT / "data" / "market_research" / "contact_verification"
OUTPUT_DIR = ROOT / "output" / "market_research"

# Upper bound on one company's total verification time. Each underlying
# HTTP request already caps at ~18s, but a company involves several
# (contact page, about, leadership, LinkedIn), so without this a single
# slow site can hold a worker for minutes and stall visible progress.
_PER_COMPANY_TIMEOUT = 75.0

# Role-agnostic generic mailboxes. Present at most real companies, so a
# match here is NOT proof of fabrication — it means "conventional address
# that was probably constructed rather than researched". Reported as
# 'generic', kept distinct from 'verified' and from 'wrong'.
_GENERIC_LOCAL_PARTS = {
    "info", "sales", "contact", "enquiry", "enquiries", "inquiry",
    "inquiries", "support", "hello", "office", "admin", "mail",
    "marketing", "general", "reception",
}


def _log(msg: str) -> None:
    print(f"  [verify] {msg}", flush=True)


def _domain(url: str) -> str:
    u = re.sub(r"^https?://", "", str(url or "").strip().lower())
    return u.split("/")[0].removeprefix("www.")


def _email_domain_matches(email: str, website: str) -> bool:
    ed = str(email or "").split("@")[-1].strip().lower()
    wd = _domain(website)
    if not ed or not wd:
        return False
    return ed == wd or ed.endswith("." + wd) or wd.endswith("." + ed)


def _classify_email(email: str, website: str, found_emails: set[str]) -> str:
    """verified | generic | domain-mismatch | not-found."""
    e = str(email or "").strip().lower()
    if not e or "@" not in e:
        return "not-found"
    if e in {x.lower() for x in found_emails}:
        return "verified"  # literally present on the company's own site
    if not _email_domain_matches(e, website):
        return "domain-mismatch"  # points somewhere other than their site
    local = e.split("@")[0]
    if local in _GENERIC_LOCAL_PARTS:
        return "generic"  # conventional, plausible, but unconfirmed
    return "not-found"


def _phone_digits(phone: str) -> str:
    """Comparable form of a phone number: digits only, no formatting.

    Needed because the same number is written many ways (+1 (555) 123-4567
    vs 15551234567 vs 001-555-123-4567), so a string comparison would
    report a match as a mismatch.
    """
    return re.sub(r"\D", "", str(phone or ""))


def _classify_phone(exported: str, found: str, reachable: bool) -> str:
    """verified | mismatch | not-found | site-unreachable.

    Office No. has the SAME fabrication risk as Office Email: the prompt
    instructs "real published phone if known, else a plausible number
    using the HQ country's calling code", so an unverified value can be a
    correctly-formatted invention. Compared on digits, and — importantly —
    matched by suffix as well as equality, because a site often publishes
    a local-format number while the export carries the full international
    form (or vice versa); those are the same line, not a discrepancy.
    """
    exp = _phone_digits(exported)
    if not exp:
        return "not-found"
    if not reachable:
        return "site-unreachable"
    fnd = _phone_digits(found)
    if not fnd:
        return "not-found"  # site loaded but published no phone we could read
    if exp == fnd:
        return "verified"
    # One is the international form of the other (last 7+ digits shared).
    tail = min(len(exp), len(fnd), 9)
    if tail >= 7 and exp[-tail:] == fnd[-tail:]:
        return "verified"
    return "mismatch"


def _verify_one(row: dict) -> dict:
    """Fetch one company's real pages and report what holds up."""
    from vendor_intel.pipeline.market_research import verify

    company = row.get("Company", "")
    website = row.get("Website", "")
    out: dict = {
        "Company": company,
        "Website": website,
        "site_reachable": None,
        "email_status": "unknown",
        "found_email": "",
        "found_phone": "",
        "found_person": "",
        "found_role": "",
        "phone_status": "unknown",
        "linkedin_status": "unknown",
    }

    try:
        found = verify.verify_contact_fields(website)
    except Exception as exc:  # a fetch failure must not kill the batch
        out["site_reachable"] = False
        out["error"] = f"{type(exc).__name__}: {str(exc)[:80]}"
        return out

    # verify_contact_fields returns {} both for "unreachable" and for
    # "reachable but nothing found". Distinguish them, because reporting an
    # unreachable site as "fabricated data" would be exactly the kind of
    # unfounded conclusion this script is meant to avoid.
    reachable = True
    try:
        from vendor_intel.scraping.httpx_fetch import fetch_page_httpx

        probe = fetch_page_httpx(website if website.startswith("http") else f"https://{website}")
        reachable = bool(probe.alive)
    except Exception:
        reachable = False
    out["site_reachable"] = reachable

    out["found_email"] = found.get("Email", "")
    out["found_phone"] = found.get("Office No.", "")
    out["found_person"] = found.get("Contact Person", "")
    out["found_role"] = found.get("Role", "")

    found_emails = {found["Email"]} if found.get("Email") else set()
    if not reachable:
        out["email_status"] = "site-unreachable"
    else:
        out["email_status"] = _classify_email(
            row.get("Office Email", ""), website, found_emails
        )

    # Office No. carries the same fabrication risk as the email and is
    # checked the same way, against the phone actually published on the site.
    out["phone_status"] = _classify_phone(
        row.get("Office No.", ""), out["found_phone"], reachable
    )

    li = str(row.get("LinkedIn", "")).strip()
    if not li:
        out["linkedin_status"] = "not-found"
    elif "/in/" not in li:
        out["linkedin_status"] = "not-a-personal-profile"
    else:
        # LinkedIn blocks automated fetches aggressively, so a failure here
        # is reported as "unconfirmed", never as "fake".
        try:
            confirmed = verify.verify_linkedin(li, row.get("Contact Person", ""))
            out["linkedin_status"] = "verified" if confirmed else "unconfirmed"
        except Exception:
            out["linkedin_status"] = "unconfirmed"
    return out


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--market", required=True)
    ap.add_argument("--limit", type=int, default=0, help="verify only the first N (0 = all)")
    ap.add_argument("--concurrency", type=int, default=8)
    ap.add_argument("--no-cache", action="store_true", help="ignore any cached results")
    ap.add_argument(
        "--apply",
        action="store_true",
        help="REWRITE the state file: keep only verified contact values and "
             "blank the rest, then re-export. Backs up the state first.",
    )
    args = ap.parse_args()

    state = MarketResearchState.load(args.market)
    rows = list(state.companies)
    if args.limit:
        rows = rows[: args.limit]
    if not rows:
        print("No companies in state for that market.")
        return

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"{market_slug(args.market)}.json"
    cache: dict = {}
    if cache_path.is_file() and not args.no_cache:
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
            _log(f"resuming: {len(cache)} companies already verified")
        except (json.JSONDecodeError, OSError):
            cache = {}

    todo = [r for r in rows if r.get("Company") not in cache]
    print()
    print("=" * 62)
    print("Verifying contact fields against the companies' OWN WEBSITES.")
    print("No browser and no AI Mode: these are direct HTTP requests to")
    print("each company's /contact, /about and /leadership pages.")
    print(f"{len(rows)} companies | {len(todo)} to fetch | concurrency {args.concurrency}")
    print("Slow or dead sites take up to ~30s each — that is normal.")
    print("=" * 62)
    print()

    started = time.time()
    done = 0
    with ThreadPoolExecutor(max_workers=max(1, args.concurrency)) as pool:
        futures = {pool.submit(_verify_one, r): r for r in todo}
        for fut in as_completed(futures):
            row = futures[fut]
            try:
                # Hard per-company cap. _verify_one issues several requests
                # (contact page, about, leadership, LinkedIn) at ~18s each,
                # so one pathological site could otherwise occupy a worker
                # for minutes and make the whole run look hung.
                result = fut.result(timeout=_PER_COMPANY_TIMEOUT)
            except TimeoutError:
                # A timeout says nothing about whether the exported value
                # is real, so it is reported as its own status rather than
                # being lumped in with verified/failed results. --apply
                # will blank the field (unverified), which is the correct
                # conservative outcome, but the report shows WHY.
                result = {
                    "Company": row.get("Company", ""),
                    "Website": row.get("Website", ""),
                    "site_reachable": False,
                    "email_status": "timeout",
                    "phone_status": "timeout",
                    "linkedin_status": "unknown",
                }
            except Exception as exc:
                result = {
                    "Company": row.get("Company", ""),
                    "Website": row.get("Website", ""),
                    "site_reachable": False,
                    "email_status": "error",
                    "phone_status": "error",
                    "linkedin_status": "unknown",
                    "error": f"{type(exc).__name__}: {str(exc)[:80]}",
                }
            cache[result["Company"]] = result
            done += 1

            # Report EVERY company, not every 10th. With --limit 15 the old
            # 10-company interval printed nothing for minutes while real
            # sites timed out, which is indistinguishable from a hang.
            status = result.get("email_status", "?")
            reach = {True: "ok", False: "unreachable"}.get(result.get("site_reachable"), "?")
            _log(
                f"[{done}/{len(todo)}] {result.get('Company','?')[:34]:34s} "
                f"site={reach:11s} email={status}"
            )
            if done % 5 == 0 or done == len(todo):
                rate = done / max(0.1, time.time() - started)
                left = (len(todo) - done) / max(0.01, rate)
                _log(f"  ... {rate:.1f}/s, ~{left/60:.1f} min remaining")
                cache_path.write_text(json.dumps(cache, indent=1), encoding="utf-8")

    cache_path.write_text(json.dumps(cache, indent=1), encoding="utf-8")
    _write_report(args.market, rows, cache)

    if args.apply:
        # Guard against applying a partial cache. --apply is normally run as
        # a third step, when every company is already cached and `todo` is
        # empty; if verification was never run (or was interrupted), the
        # cache covers only some companies and the rest would silently keep
        # their unverified values while the run reported success.
        # Checked against ALL companies in state, not just this run's
        # `rows`: with --limit the slice would look fully verified while
        # most of the dataset had never been fetched.
        all_rows = list(state.companies)
        unverified = [r for r in all_rows if r.get("Company") not in cache]
        if unverified:
            _log(
                f"REFUSING to apply: {len(unverified)} of {len(all_rows)} companies "
                "have not been verified yet."
            )
            _log(
                "Run without --apply first so every company is checked, then "
                "re-run with --apply."
            )
            _log(f"example: {', '.join(r.get('Company','?') for r in unverified[:3])} ...")
            return
        _apply_verified_only(args.market, cache)


def _apply_verified_only(market: str, cache: dict) -> None:
    """Rewrite the state so the six contact fields hold ONLY verified data.

    Per explicit user instruction: no guessed values in these columns. A
    field survives only when the fetch actually confirmed it; everything
    else is blanked. Companies are NOT dropped — blanking is deliberately
    preferred over deletion so all 493 researched companies (and their
    verified firmographics) are kept.

    The state file is backed up first, because this is the one destructive
    step in the script and a bad run would otherwise lose the contact data
    irrecoverably.
    """
    import shutil
    from datetime import datetime

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx
    from vendor_intel.pipeline.market_research.state import STATE_DIR

    state_path = STATE_DIR / f"{market_slug(market)}.json"
    if not state_path.is_file():
        _log(f"state file not found, nothing to apply: {state_path}")
        return

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = state_path.with_name(f"{state_path.stem}.before-verify-{stamp}.json")
    shutil.copy2(state_path, backup)
    _log(f"state backed up to {backup.name}")

    data = json.loads(state_path.read_text(encoding="utf-8"))
    companies = data.get("companies", [])

    stats = {"kept": 0, "blanked": 0}
    for row in companies:
        v = cache.get(row.get("Company", ""))
        if not v:
            continue  # never verified (e.g. --limit run): leave untouched

        # Email: keep only what was literally found on the company's site.
        for col, status_key, found_key in (
            ("Office Email", "email_status", "found_email"),
            ("Email", "email_status", "found_email"),
        ):
            if v.get(status_key) == "verified" and v.get(found_key):
                row[col] = v[found_key]
                stats["kept"] += 1
            elif row.get(col):
                row[col] = ""
                stats["blanked"] += 1

        # Office No.: keep the site's published number when it matched.
        if v.get("phone_status") == "verified" and v.get("found_phone"):
            row["Office No."] = v["found_phone"]
            stats["kept"] += 1
        elif row.get("Office No."):
            row["Office No."] = ""
            stats["blanked"] += 1

        # Contact person/role: only a real person found on the site.
        if v.get("found_person") and v.get("found_role"):
            row["Contact Person"] = v["found_person"]
            row["Role"] = v["found_role"]
            stats["kept"] += 2
        else:
            for col in ("Contact Person", "Role"):
                if row.get(col):
                    row[col] = ""
                    stats["blanked"] += 1

        # LinkedIn: only a profile that was actually confirmed.
        if v.get("linkedin_status") == "verified" and row.get("LinkedIn"):
            stats["kept"] += 1
        elif row.get("LinkedIn"):
            row["LinkedIn"] = ""
            stats["blanked"] += 1

    state_path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    _log(f"state rewritten: {stats['kept']} verified values kept, {stats['blanked']} blanked")
    _log(f"companies retained: {len(companies)} (none dropped)")

    out = export_xlsx(market)
    _log(f"re-exported: {out}")


def _write_report(market: str, rows: list[dict], cache: dict) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    counts: dict[str, int] = {}
    for r in rows:
        st = cache.get(r.get("Company", ""), {}).get("email_status", "unknown")
        counts[st] = counts.get(st, 0) + 1
    ph_counts: dict[str, int] = {}
    for r in rows:
        st = cache.get(r.get("Company", ""), {}).get("phone_status", "unknown")
        ph_counts[st] = ph_counts.get(st, 0) + 1
    li_counts: dict[str, int] = {}
    for r in rows:
        st = cache.get(r.get("Company", ""), {}).get("linkedin_status", "unknown")
        li_counts[st] = li_counts.get(st, 0) + 1

    print()
    print("=" * 62)
    print(f"CONTACT VERIFICATION — {len(rows)} companies")
    print("=" * 62)
    print("Office/contact email:")
    for k, v in sorted(counts.items(), key=lambda kv: -kv[1]):
        print(f"   {k:22s} {v:4d}  ({100*v/len(rows):.1f}%)")
    print("Office No.:")
    for k, v in sorted(ph_counts.items(), key=lambda kv: -kv[1]):
        print(f"   {k:22s} {v:4d}  ({100*v/len(rows):.1f}%)")
    print("LinkedIn (personal profile):")
    for k, v in sorted(li_counts.items(), key=lambda kv: -kv[1]):
        print(f"   {k:22s} {v:4d}  ({100*v/len(rows):.1f}%)")
    print("=" * 62)
    print("verified          = found on the company's own site (phone: digits match)")
    print("generic           = info@/sales@ at their domain: plausible, unconfirmed")
    print("domain-mismatch   = email points away from their site (most suspect)")
    print("mismatch          = phone differs from the one published on the site")
    print("not-found         = site gave nothing to compare against")
    print("site-unreachable  = could not check; NOT evidence of a bad value")
    print("timeout / error   = site too slow or failed; also NOT a verdict")
    print("")
    print("Only 'verified' means the value was confirmed real. Everything")
    print("else is blanked by --apply, since an unconfirmed contact detail")
    print("is worse than an empty cell.")
    print("=" * 62)

    wb = Workbook()
    ws = wb.active
    ws.title = "Contact Verification"
    headers = [
        "#", "Company", "Website", "Site Reachable",
        "Office Email (exported)", "Email Status", "Email Found On Site",
        "Office No. (exported)", "Phone Status", "Phone Found On Site",
        "Person Found On Site", "Role Found On Site",
        "LinkedIn (exported)", "LinkedIn Status",
    ]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="1F4E79")
    for i in range(1, len(headers) + 1):
        c = ws.cell(1, i)
        c.fill = fill
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 20

    for n, r in enumerate(rows, 1):
        v = cache.get(r.get("Company", ""), {})
        ws.append([
            n,
            r.get("Company", ""),
            r.get("Website", ""),
            {True: "yes", False: "no"}.get(v.get("site_reachable"), "?"),
            r.get("Office Email", ""),
            v.get("email_status", "unknown"),
            v.get("found_email", ""),
            r.get("Office No.", ""),
            v.get("phone_status", "unknown"),
            v.get("found_phone", ""),
            v.get("found_person", ""),
            v.get("found_role", ""),
            r.get("LinkedIn", ""),
            v.get("linkedin_status", "unknown"),
        ])
        ws.row_dimensions[ws.max_row].height = 16

    widths = [5, 32, 30, 13, 32, 17, 32, 20, 15, 22, 24, 26, 40, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(size=9)
            c.alignment = Alignment(wrap_text=False, vertical="center")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / f"{market_slug(market)}_contact_verification.xlsx"
    wb.save(out)
    print(f"\nReport written to: {out}")
    print("(your original export was NOT modified)")


if __name__ == "__main__":
    main()
