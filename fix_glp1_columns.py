#!/usr/bin/env python3
"""Repair the geography and contact columns of an existing market workbook.

WHY. A measurement pass over GLP1_Receptor_Agonist_Distributors.xlsx (270
rows) found:

  * Continent / Geography named NO country in 270/270 rows ("Asia",
    "Europe - United Kingdom", "Asia / Global").
  * Operational Presence used counts or hand-waving in 59 rows
    ("India + international export", "USA, Canada + international").
  * Contact Person was a team or company in 128 rows (47%) --
    "ALDAWLIA Management Team", "A.K. Pharma Management".
  * LinkedIn held only 2 real personal profiles; 156 were SEARCH URLs
    (".../search/results/companies/?keywords=...") and 95 were company
    pages. A search URL is proof of construction, not research.
  * Email equalled Office Email in 268/270 rows (99%) -- one field copied
    into the other.

This asks Google AI Mode to redo those columns with an explicit
instruction that an unknown value must come back empty, then writes "NPD"
(No Published Data) wherever nothing could be established.

NON-DESTRUCTIVE: the original workbook is copied to a timestamped backup
before anything is written, and every change is also recorded in an
"Audit" sheet so each edit can be traced back.

HONEST LIMITATION: AI Mode is a generative model, not a registry. It is
better than the original pass because the prompt now REWARDS "unknown"
instead of forbidding it, but a value it returns is still unverified
unless separately checked against the company's own website. Treat "NPD"
as trustworthy (nothing was claimed) and a filled value as a lead.

Usage:
    # see what would change, write nothing
    python fix_glp1_columns.py --file "<path.xlsx>" --dry-run --limit 18

    # geography only, or contacts only
    python fix_glp1_columns.py --file "<path.xlsx>" --only geo
    python fix_glp1_columns.py --file "<path.xlsx>" --only contacts

    # full run (backs up first)
    python fix_glp1_columns.py --file "<path.xlsx>"
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

# Windows consoles are cp1252 and print() RAISES on an unmappable
# character -- a Turkish company name killed a whole run this way. This
# data is worldwide, so non-ASCII names are normal input.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

apply_env_overrides()

from vendor_intel.pipeline.market_research import ai_mode_client  # noqa: E402

NPD = "NPD"

GEO_COLS = ("Continent / Geography", "Operational Presence")
CONTACT_COLS = (
    "Contact Person", "Role", "Email", "LinkedIn", "Office No.", "Office Email",
)

# Batch sizes differ because the answers differ in size: a geography row
# can be a 300-character country list, so fewer per query.
GEO_BATCH = 8
CONTACT_BATCH = 15


def _log(msg: str) -> None:
    text = f"  [fix] {msg}"
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "ascii"
        print(text.encode(enc, errors="replace").decode(enc), flush=True)


# --------------------------------------------------------------------------
# Prompts
# --------------------------------------------------------------------------

_GEO_SYSTEM = (
    "You are a precise company-geography researcher. You name individual "
    "countries, never regions or counts. You only state facts you actually "
    "know; for anything you do not know you return an empty string."
)

_GEO_PROMPT = """List the countries for each company below. Fill BOTH fields.

continent_geography = every country where the company has a physical site (headquarters, office, warehouse, distribution centre, plant, lab).
Write as: Region(Country, Country); Region(Country)
Like this: North America(United States, Canada); Europe(Germany, France, United Kingdom, Spain); Asia(India, Singapore)

operational_presence = every country where the company does business: sells, supplies, distributes, exports to, or services. This is WIDER than continent_geography and must include all of those countries plus any others. Write as a plain list: Country; Country; Country

BE COMPLETE. A large distributor operates in dozens of countries — list them all, not just the head office country. Think about its subsidiaries, regional branches and export markets. Ten or twenty countries is a normal answer for a multinational; one country is normal only for a genuinely single-country business.

Name every country individually. Never answer with only a region name, never with a number of countries, and never with words like global or international.

If operational_presence is unknown, repeat the continent_geography countries there instead of leaving it blank.

Companies:
{companies}

Reply with only a JSON array, one object per company, same order:
[{{"company": "", "continent_geography": "", "operational_presence": ""}}]"""


_CONTACT_SYSTEM = (
    "You are a precise B2B contact researcher. You report ONLY real, "
    "published contact details that you actually know. You never invent or "
    "construct a person's name, an email address, a phone number or a "
    "LinkedIn URL. An empty string is the correct answer when you do not "
    "know a value."
)

_CONTACT_PROMPT = """For each company below, give ONLY its REAL, PUBLISHED contact details.

DO NOT HALLUCINATE. DO NOT RANDOM-FILL. Return "" for anything you do not \
actually know. These specific failures are being corrected, so do not \
reproduce them:
- contact_person MUST be a real NAMED INDIVIDUAL. "ALDAWLIA Management \
Team", "A.K. Pharma Management", "Sales Team", or the company's own name \
are all WRONG — return "" instead.
- linkedin MUST be that person's personal profile \
("linkedin.com/in/<slug>"). A company page ("/company/...") is WRONG. A \
SEARCH URL ("/search/results/companies/?keywords=...") is WRONG and is \
never a real profile — return "" instead.
- NEVER build an email from the domain. "info@<their-domain>" invented \
from the website is FORBIDDEN unless that exact address is genuinely \
published.
- NEVER invent a phone number, and never generate one from a country \
calling code.
- office_email must be independently known. Do NOT simply repeat `email`.

An empty field is CORRECT and preferred. A wrong contact detail is much \
worse than an empty one.

Companies:
{companies}

Reply with ONLY a JSON array, no markdown fences and no commentary, one \
object per company IN THE SAME ORDER:
[{{"company": "", "contact_person": "", "role": "", "email": "", \
"linkedin": "", "office_no": "", "office_email": ""}}]"""


# --------------------------------------------------------------------------
# Validation — the prompt is a request, the code is the guarantee
# --------------------------------------------------------------------------

_VAGUE_GEO = re.compile(
    r"\b(global|worldwide|international|multinational|nationwide|"
    r"\d+\s*\+?\s*countries?)\b", re.I)

_BARE_REGION = re.compile(
    r"^\s*(north america|south america|latin america|europe|asia|africa|"
    r"middle east|oceania|apac|emea|eu|na)"
    r"(\s*[/&+,-]+\s*(north america|south america|latin america|europe|asia|"
    r"africa|middle east|oceania|global|worldwide))*\s*$", re.I)

_TEAMISH = re.compile(
    r"\b(management|team|dept|department|division|board|committee|"
    r"customer\s*(care|service|support)|admin|sales)\b|"
    r"\b(pvt|ltd|inc|llc|gmbh|corp|corporation|co|group|pharma|"
    r"pharmaceuticals?|healthcare|labs?|laboratories)\b\.?\s*$", re.I)

_LINKEDIN_IN = re.compile(r"linkedin\.com/in/[A-Za-z0-9\-_%]+", re.I)
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")
_GENERIC_LOCAL = frozenset({
    "info", "sales", "contact", "enquiry", "enquiries", "inquiry", "support",
    "hello", "office", "admin", "mail", "general", "reception", "marketing",
    "customer.service", "customerservice", "rx", "investor.relations",
})

_UNKNOWN = re.compile(
    r"^\s*(|-|n/?a|na|none|null|nil|unknown|npd|not\s+\w+|no\s+\w+.*|tbd)\s*\.?\s*$",
    re.I)


def _clean(value: object) -> str:
    text = " ".join(str(value or "").split())
    return "" if _UNKNOWN.match(text) else text


def countries_in(text: str) -> list[str]:
    """Country names inside a "REGION(A, B); REGION(C)" or "A; B" value.

    Used to enforce the relationship between the two columns: Operational
    Presence is defined as BROADER than physical presence, so it must
    contain at least the physical countries.
    """
    t = _clean(text)
    if not t:
        return []
    inner = re.findall(r"\(([^)]*)\)", t)
    if inner:
        blob = ", ".join(inner)
    else:
        # No brackets: the whole value is the list. Strip a leading
        # "Region:" label if present, then split on separators. A single
        # bare country name ("India") must survive this path.
        blob = re.sub(r"^[^(:]*?:", "", t)
    parts = re.split(r"[;,]", blob)
    out, seen = [], set()
    for raw in parts:
        c = re.sub(r"\s+", " ", raw).strip(" .;")
        # The old file used "Europe - United Kingdom" / "Asia / Global".
        # Without splitting on those separators the whole string is taken
        # as ONE country name, so a bare region sneaks through validation
        # as if a country had been named. Keep the part after the dash or
        # slash, which is where the actual country sits.
        for sep in (" - ", " -- ", " / ", " – ", " — "):
            if sep in c:
                c = c.split(sep)[-1].strip()
        # Drop region words and non-place words outright, so they cannot
        # be counted or re-emitted as if they were countries.
        if not c or _BARE_REGION.match(c) or _VAGUE_GEO.search(c):
            continue
        k = c.casefold()
        if k not in seen:
            seen.add(k)
            out.append(c)
    return out


def widen_operational(physical: str, operational: str) -> str:
    """Guarantee Operational Presence >= Continent / Geography.

    The model repeatedly returned an EMPTY operational_presence while
    filling continent_geography, which the validator then turned into
    "NPD" — losing information we demonstrably had. By definition a
    company operates everywhere it has a physical site, so the physical
    countries are a valid floor. Any extra countries the model supplied
    are merged in, never dropped.
    """
    phys = countries_in(physical)
    oper = countries_in(operational)
    if not phys and not oper:
        return ""
    merged = list(oper)
    have = {c.casefold() for c in merged}
    for c in phys:
        if c.casefold() not in have:
            have.add(c.casefold())
            merged.append(c)
    return "; ".join(merged)


def valid_geo(text: str) -> bool:
    """A geography value is acceptable only if it names countries."""
    t = _clean(text)
    if not t:
        return False
    if _VAGUE_GEO.search(t) or _BARE_REGION.match(t):
        return False
    # A single country is a LEGITIMATE answer -- many regional
    # distributors operate in exactly one. Requiring a separator here
    # rejected "India" and overwrote it with NPD, which discarded 159
    # correct values in the first real run. What actually matters is that
    # at least one COUNTRY is named, which countries_in() determines.
    return bool(countries_in(t))


def is_person_name(text: str) -> bool:
    """True when the value looks like an actual individual.

    Rejects "X Management Team" and bare company names, which made up 47%
    of the original Contact Person column.
    """
    t = _clean(text)
    if not t or _TEAMISH.search(t):
        return False
    words = [w for w in re.split(r"\s+", t) if w]
    if len(words) < 2:          # a lone first name is not a usable contact
        return False
    return all(re.match(r"^[A-Za-z][A-Za-z.'À-ɏ-]*$", w) for w in words)


def clean_linkedin(text: str) -> str:
    """Only a real personal /in/ profile survives; never repair one."""
    t = _clean(text)
    if not t or "/search/" in t.lower() or "/company/" in t.lower():
        return ""
    m = _LINKEDIN_IN.search(t)
    if not m:
        return ""
    url = m.group(0)
    return url if url.lower().startswith("http") else f"https://www.{url}"


def clean_email(text: str, website: str, *, other: str = "") -> str:
    """Reject a generic mailbox at the company's own domain, and a value
    merely copied from the sibling email field."""
    t = _clean(text).lower()
    if not t or not _EMAIL_RE.match(t):
        return ""
    if other and t == _clean(other).lower():
        return ""       # Email == Office Email was 99% of the original file
    local, _, domain = t.partition("@")
    site = re.sub(r"^https?://", "", str(website or "").lower()).split("/")[0]
    site = site.removeprefix("www.")
    if site and (domain == site or domain.endswith("." + site)):
        if local in _GENERIC_LOCAL:
            return ""   # the exact info@<domain> construction being removed
    return t


def clean_phone(text: str) -> str:
    t = _clean(text)
    return t if len(re.sub(r"\D", "", t)) >= 7 else ""


# --------------------------------------------------------------------------
# AI Mode plumbing
# --------------------------------------------------------------------------

def _extract_array(raw: str, first_key: str) -> list[dict]:
    """Pull the JSON array out of AI Mode's rendered page text.

    AI Mode echoes the prompt, which ENDS with an empty JSON template, so
    a naive find("[") locks onto the echo and yields nothing. Take the LAST
    array that opens with a real key, and slice FROM it (slicing past the
    marker drops the first record).
    """
    text = (raw or "").replace("Use code with caution.", "").strip()
    starts = [m.start() for m in re.finditer(r'\[\s*\{\s*"' + first_key, text)]
    if starts:
        text = text[starts[-1]:]
    a, b = text.find("["), text.rfind("]")
    if a < 0 or b <= a:
        return []
    candidate = text[a:b + 1]
    for attempt in (candidate, re.sub(r"\s*\n\s*", " ", candidate)):
        try:
            data = json.loads(attempt)
        except (json.JSONDecodeError, ValueError):
            continue
        if isinstance(data, list):
            return [d for d in data if isinstance(d, dict)]
    return []


# Google refuses BURSTS of requests ("It looks like there's no response
# available for this search. Try asking something else.") and then
# recovers on its own — verified by re-running the identical prompt for the
# identical company and getting a 52-country answer. Giving up after one
# refusal therefore threw away whole batches for no reason: 4 consecutive
# batches came back empty while the same query worked when retried by hand.
_BATCH_RETRIES = 4
_BATCH_BACKOFF_SECONDS = 30.0


def _ask(system: str, prompt: str, first_key: str) -> list[dict]:
    """Ask one batch, retrying through transient refusals.

    Backs off between attempts because an immediate retry usually returns
    the same refusal, while waiting half a minute usually clears it.
    """
    for attempt in range(1, _BATCH_RETRIES + 1):
        try:
            raw = ai_mode_client.chat(system, prompt)
        except Exception as exc:                   # CAPTCHA / refusal / timeout
            raw = ""
            reason = type(exc).__name__
        else:
            out = _extract_array(raw, first_key)
            if out:
                return out
            reason = "unparseable or refused"

        if attempt < _BATCH_RETRIES:
            wait = _BATCH_BACKOFF_SECONDS * attempt
            _log(
                f"batch attempt {attempt}/{_BATCH_RETRIES} got nothing "
                f"({reason}) — waiting {wait:.0f}s and retrying"
            )
            time.sleep(wait)
        else:
            _log(f"batch failed after {_BATCH_RETRIES} attempts ({reason}) — skipping")
    return []


def _match_by_name(rows: list[dict], answers: list[dict]) -> dict[str, dict]:
    """Match answers to companies BY NAME, never by position alone.

    The model reorders and drops entries; position-matching would write
    one company's data onto another.
    """
    by_name = {str(a.get("company", "")).strip().casefold(): a for a in answers}
    out: dict[str, dict] = {}
    for n, row in enumerate(rows):
        name = row["Company"]
        ans = by_name.get(name.strip().casefold())
        if ans is None and len(answers) == len(rows):
            ans = answers[n]                       # safe only when counts match
        if ans is not None:
            out[name] = ans
    return out


# --------------------------------------------------------------------------
# Passes
# --------------------------------------------------------------------------

def fix_geography(rows: list[dict]) -> dict[str, dict]:
    results: dict[str, dict] = {}
    batches = [rows[i:i + GEO_BATCH] for i in range(0, len(rows), GEO_BATCH)]
    for n, batch in enumerate(batches, 1):
        listing = "\n".join(
            f"{k}. {r['Company']} — website {r.get('Website','')} — HQ {r.get('Headquarters','')}"
            for k, r in enumerate(batch, 1)
        )
        answers = _ask(_GEO_SYSTEM, _GEO_PROMPT.format(companies=listing), "company")
        matched = _match_by_name(batch, answers)
        kept = 0
        for name, ans in matched.items():
            cg = _clean(ans.get("continent_geography"))
            op = _clean(ans.get("operational_presence"))
            good_cg = cg if valid_geo(cg) else ""
            # Operational Presence is BROADER by definition, so backfill it
            # from the physical countries rather than writing NPD over data
            # we already have.
            widened = widen_operational(good_cg, op if valid_geo(op) else "")
            results[name] = {
                "Continent / Geography": good_cg,
                "Operational Presence": widened if valid_geo(widened) else "",
            }
            kept += sum(1 for v in results[name].values() if v)
        _log(f"geo batch {n}/{len(batches)}: {len(matched)} answered, {kept} usable values")
    return results


def fix_contacts(rows: list[dict]) -> dict[str, dict]:
    results: dict[str, dict] = {}
    batches = [rows[i:i + CONTACT_BATCH] for i in range(0, len(rows), CONTACT_BATCH)]
    for n, batch in enumerate(batches, 1):
        listing = "\n".join(
            f"{k}. {r['Company']} — website {r.get('Website','')} — HQ {r.get('Headquarters','')}"
            for k, r in enumerate(batch, 1)
        )
        answers = _ask(_CONTACT_SYSTEM, _CONTACT_PROMPT.format(companies=listing), "company")
        matched = _match_by_name(batch, answers)
        kept = 0
        for name, ans in matched.items():
            site = next((r.get("Website", "") for r in batch if r["Company"] == name), "")
            person = _clean(ans.get("contact_person"))
            role = _clean(ans.get("role"))
            if not is_person_name(person):
                person, role = "", ""              # a team name is not a contact
            if not role:
                person = ""                        # a name with no title is unusable
            email = clean_email(ans.get("email", ""), site)
            office_email = clean_email(ans.get("office_email", ""), site, other=email)
            linkedin = clean_linkedin(ans.get("linkedin", ""))
            if not person:
                linkedin = ""                      # cannot attribute a profile
            fixed = {
                "Contact Person": person,
                "Role": role,
                "Email": email,
                "LinkedIn": linkedin,
                "Office No.": clean_phone(ans.get("office_no", "")),
                "Office Email": office_email,
            }
            results[name] = fixed
            kept += sum(1 for v in fixed.values() if v)
        _log(f"contact batch {n}/{len(batches)}: {len(matched)} answered, {kept} real values")
    return results


# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", required=True)
    ap.add_argument("--sheet", default="")
    ap.add_argument("--only", choices=("geo", "contacts", "both"), default="both")
    ap.add_argument("--limit", type=int, default=0, help="first N rows only (0 = all)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--placeholder", default=NPD,
                    help='text for unverifiable cells (default "NPD"; use "" for blank)')
    args = ap.parse_args()

    import openpyxl

    path = Path(args.file)
    wb = openpyxl.load_workbook(path)
    ws = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]
    hdr = [c.value for c in ws[1]]
    idx = {h: n for n, h in enumerate(hdr)}

    rows: list[dict] = []
    for excel_row in range(2, ws.max_row + 1):
        rec = {h: str(ws.cell(excel_row, n + 1).value or "").strip()
               for h, n in idx.items() if h}
        if not rec.get("Company"):
            continue
        rec["_row"] = excel_row
        rows.append(rec)
    if args.limit:
        rows = rows[: args.limit]

    targets: list[str] = []
    if args.only in ("geo", "both"):
        targets += [c for c in GEO_COLS if c in idx]
    if args.only in ("contacts", "both"):
        targets += [c for c in CONTACT_COLS if c in idx]

    print()
    print("=" * 70)
    print(f"Re-researching {len(rows)} companies via Google AI Mode")
    print(f"columns      : {', '.join(targets)}")
    print(f"unverifiable : {args.placeholder!r}")
    print("Unknown values are NOT guessed — they become the placeholder.")
    if args.dry_run:
        print("DRY RUN — nothing will be written.")
    print("=" * 70)
    print()

    geo = fix_geography(rows) if args.only in ("geo", "both") else {}
    con = fix_contacts(rows) if args.only in ("contacts", "both") else {}

    changes: list[tuple[str, str, str, str]] = []
    stats = {"filled": 0, "placeholder": 0, "unanswered": 0}
    for rec in rows:
        name = rec["Company"]
        merged = {**geo.get(name, {}), **con.get(name, {})}
        if not merged:
            stats["unanswered"] += 1
            continue
        for col, new in merged.items():
            old = rec.get(col, "")
            final = new or args.placeholder
            if final == old:
                continue
            changes.append((name, col, old, final))
            if new:
                stats["filled"] += 1
            else:
                stats["placeholder"] += 1
            if not args.dry_run:
                ws.cell(rec["_row"], idx[col] + 1).value = final

    print()
    print("=" * 70)
    print(f"real values written      : {stats['filled']}")
    print(f"set to {args.placeholder!r:<18}: {stats['placeholder']}")
    print(f"companies not answered   : {stats['unanswered']} (left unchanged)")
    print("=" * 70)
    for name, col, old, new in changes[:15]:
        print(f"  {name[:24]:24s} {col:22s} {old[:26]!r} -> {new[:40]!r}")
    if len(changes) > 15:
        print(f"  ... and {len(changes)-15} more")

    if args.dry_run:
        print("\nDRY RUN — the workbook was NOT modified.")
        return

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = path.with_name(f"{path.stem}.before-fix-{stamp}.xlsx")
    shutil.copy2(path, backup)
    _log(f"backup written: {backup.name}")

    audit = wb.create_sheet("Audit")
    audit.append(["Company", "Column", "Before", "After"])
    for row in changes:
        audit.append(list(row))
    for col, width in zip("ABCD", (34, 24, 60, 60)):
        audit.column_dimensions[col].width = width

    wb.save(path)
    _log(f"saved: {path}")
    print("\nEvery edit is listed on the 'Audit' sheet.")


if __name__ == "__main__":
    main()
