#!/usr/bin/env python3
"""Fill Office No. and Office Email only, allowing MULTIPLE values per company.

Coverage before this existed: Office No. 383/1891 (20%), Office Email
268/1891 (14%), and not one row carried more than a single value.

Two reasons the existing refill could not do better:

  1. Its prompt asks for "a real published telephone number" -- singular.
     A distributor with offices in eight countries publishes eight
     numbers, and all of them are genuine; asking for one throws the rest
     away. Companies here average physical presence in several countries
     (that is what the Continent / Geography column records), so
     single-value was the wrong shape from the start.

  2. It rejects a generic mailbox at the company's own domain
     (info@acme.com) because the ORIGINAL discovery prompt used to
     fabricate exactly that pattern from the domain. That rule is right
     for a PERSONAL contact email, and wrong for an OFFICE one: a
     published info@ / sales@ address is usually the only address a
     company lists, and refusing it is why Office Email sits at 14%.

     The distinction kept here: a generic address is accepted when the
     model states it is actually published on the site, and still
     rejected when the model cannot point to it -- so the anti-fabrication
     intent survives without discarding real data.

Everything else about the anti-fabrication rules is unchanged: never
invent a number, never derive one from a country calling code, never
guess an address, and an empty answer is always acceptable.

Existing values are never overwritten. New values are appended only when
they are genuinely new, so running this twice is safe.

Usage:
    python fill_office_contacts.py --market "Global Polyurethane Market - Distributors Only" --dry-run
    python fill_office_contacts.py --all --export
    python fill_office_contacts.py --all --export --only-empty   # skip rows that already have both
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

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from dotenv import load_dotenv  # noqa: E402

load_dotenv(ROOT / ".env")

from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)
from vendor_intel.scraping import google_ai_mode as gam  # noqa: E402

CHUNK_SIZE = 6          # small: each answer carries several numbers/emails
_ASK_RETRIES = 3
_MAX_VALUES = 6         # per field, per company

PHONE = "Office No."
EMAIL = "Office Email"


# --- prompt ---------------------------------------------------------------


def _prompt(rows: list[dict]) -> str:
    listing = "\n".join(
        f"{i}. {r.get('Company', '')} — {r.get('Website', '')} "
        f"(offices: {str(r.get('Continent / Geography') or '')[:120]})"
        for i, r in enumerate(rows, 1)
    )
    return (
        "For each company below, list the OFFICE telephone numbers and "
        "OFFICE email addresses that are actually PUBLISHED on its own "
        "website (contact page, footer, regional office pages) or in its "
        "public filings.\n\n"
        "These companies have offices in several countries, so MORE THAN "
        "ONE number and address is normal and expected. List every one you "
        "genuinely know, separated by '; '.\n\n"
        f"{listing}\n\n"
        "Reply with ONLY a JSON array, no markdown fences, no commentary:\n"
        '[{"company": "", "office_no": "", "office_email": "", '
        '"published": true}]\n'
        "Rules:\n"
        "- office_no: real published numbers in international format "
        '(e.g. "+49 211 5401 0; +1 630 587 9600"). NEVER invent a number '
        "and never build one from a country calling code.\n"
        "- office_email: real published addresses "
        '(e.g. "info@acme.com; sales.uk@acme.com"). A generic mailbox '
        "like info@ or sales@ IS acceptable when the company genuinely "
        "publishes it — that is what most companies list.\n"
        "- NEVER construct an address from the domain. If you do not know "
        'what the company publishes, return "" for that field.\n'
        '- published: true only if you are recalling values you actually '
        "know are on the company's own site; false if you are unsure.\n"
        '- Use "" for anything you do not know. An empty answer is always '
        "acceptable; a fabricated one is not.\n"
        "- One object per company, same order as the list."
    )


# --- parsing --------------------------------------------------------------


def _extract(text: str) -> list[dict] | None:
    """Pull the answer array out of AI Mode's rendered page text.

    AI Mode echoes the prompt back before answering, and this prompt
    contains a JSON template. Scanning from the start therefore locks onto
    the echoed template and yields its empty placeholders -- the failure
    that made an earlier backfill silently refill nothing. Cutting at the
    LAST template marker keeps the generated payload.
    """
    if not text:
        return None
    cleaned = re.sub(r"^\s*```(?:json)?|```\s*$", "", text.strip(),
                     flags=re.I | re.M)

    starts = [m.start() for m in re.finditer(r'\[\s*\{\s*"company"', cleaned)]
    if starts:
        cleaned = cleaned[starts[-1]:]

    start = cleaned.find("[")
    while start != -1:
        depth = 0
        for i in range(start, len(cleaned)):
            if cleaned[i] == "[":
                depth += 1
            elif cleaned[i] == "]":
                depth -= 1
                if depth == 0:
                    try:
                        obj = json.loads(cleaned[start : i + 1])
                    except json.JSONDecodeError:
                        break
                    if isinstance(obj, list) and obj:
                        return [o for o in obj if isinstance(o, dict)]
                    break
        start = cleaned.find("[", start + 1)
    return None


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name or "").lower())


_UNKNOWN = re.compile(
    r"^\s*(|-|n/?a|na|none|null|unknown|not\s+(available|found|known|published)"
    r"|no\s+(data|info|information)|tbd|npd)\s*$", re.I)


def _split(value: object) -> list[str]:
    """Split a multi-value answer into individual entries."""
    text = " ".join(str(value or "").split())
    if _UNKNOWN.match(text):
        return []
    parts = re.split(r"\s*[;|]\s*|\s{2,}|\s*,\s*(?=\+?\d|[A-Za-z0-9._%-]+@)", text)
    return [p.strip(" .,;") for p in parts if p.strip(" .,;")]


def clean_phones(value: object) -> list[str]:
    """Valid published phone numbers, deduplicated.

    A real number needs enough digits to be dialable. The 7-digit floor
    is the same one contact_lookup uses, and it rejects the truncated
    fragments a split occasionally produces.
    """
    out, seen = [], set()
    for part in _split(value):
        digits = re.sub(r"\D", "", part)
        if len(digits) < 7 or len(digits) > 18:
            continue
        if digits in seen:
            continue
        seen.add(digits)
        out.append(part)
    return out[:_MAX_VALUES]


_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")


def clean_emails(value: object) -> list[str]:
    """Syntactically valid email addresses, deduplicated.

    Generic mailboxes (info@, sales@) are KEPT here, unlike in
    contact_lookup. That rule exists because the discovery prompt used to
    fabricate info@<domain> from the website, but for an OFFICE address a
    published info@ is usually the only address there is -- rejecting it
    is why this column sat at 14%. The `published` flag from the model is
    what guards against invention instead.
    """
    out, seen = [], set()
    for part in _split(value):
        addr = part.strip("<>()[]").lower()
        if not _EMAIL_RE.match(addr) or addr in seen:
            continue
        seen.add(addr)
        out.append(addr)
    return out[:_MAX_VALUES]


def merge(existing: object, found: list[str], *, is_email: bool) -> str:
    """Existing values first, then genuinely new ones appended.

    Never overwrites: a value already in the sheet was verified earlier
    and this pass has no grounds to replace it.
    """
    current = clean_emails(existing) if is_email else clean_phones(existing)
    seen = {(re.sub(r"\D", "", c) if not is_email else c.lower()) for c in current}
    merged = list(current)
    for value in found:
        key = value.lower() if is_email else re.sub(r"\D", "", value)
        if key in seen:
            continue
        seen.add(key)
        merged.append(value)
    return "; ".join(merged[:_MAX_VALUES])


# --- the run --------------------------------------------------------------


class RateLimited(RuntimeError):
    """Google has no AI Mode answer budget left for this account."""


def _all_blank(answers: list[dict]) -> bool:
    """True when every field of every answer is empty.

    A real chunk of six companies never comes back completely empty. What
    does look like this is the prompt's own JSON template echoed back with
    its placeholders stripped -- which is exactly what the page contains
    when Google has stopped answering. Treating that as "the model does
    not know these six" is how ~118 chunks silently did nothing.
    """
    for a in answers:
        if str(a.get("company") or "").strip():
            return False
        if str(a.get("office_no") or "").strip():
            return False
        if str(a.get("office_email") or "").strip():
            return False
    return True


def _ask(chunk: list[dict]) -> dict[str, dict]:
    prompt = _prompt(chunk)
    for attempt in range(_ASK_RETRIES):
        text = gam.ask(prompt, retries=2)

        # Quota exhausted: stop the whole run. Retrying cannot help, and
        # continuing would march through every remaining chunk producing
        # nothing while looking like a successful pass.
        if text and gam.is_ai_mode_rate_limited(text):
            raise RateLimited(
                "Google: 'You have reached the request limit for AI "
                "responses'. Progress is saved — re-run this command later "
                "and it will resume from here."
            )

        if text and not gam.is_ai_mode_refusal(text):
            answers = _extract(text)
            if answers and not _all_blank(answers):
                return {_norm(a.get("company")): a for a in answers}
        print(f"      retry {attempt + 1}/{_ASK_RETRIES}")
        gam.shutdown()
        time.sleep(3)
    return {}


def process(market: str, *, dry_run: bool, only_empty: bool, limit: int) -> dict:
    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        print(f"  no state file: {path.name}")
        return {}

    data = json.loads(path.read_text(encoding="utf-8"))
    companies = data.get("companies") or []

    if only_empty:
        targets = [r for r in companies
                   if not str(r.get(PHONE) or "").strip()
                   or not str(r.get(EMAIL) or "").strip()]
    else:
        targets = list(companies)
    if limit:
        targets = targets[:limit]

    before_p = sum(1 for r in companies if str(r.get(PHONE) or "").strip())
    before_e = sum(1 for r in companies if str(r.get(EMAIL) or "").strip())
    print(f"{market}")
    print(f"  {len(companies)} companies | phone {before_p}, email {before_e} "
          f"| asking about {len(targets)}")

    if dry_run:
        print("  DRY RUN — nothing will be written")
        return {}

    added_p = added_e = 0
    total_chunks = (len(targets) + CHUNK_SIZE - 1) // CHUNK_SIZE
    for n, i in enumerate(range(0, len(targets), CHUNK_SIZE), 1):
        chunk = targets[i : i + CHUNK_SIZE]
        print(f"    chunk {n}/{total_chunks} ({len(chunk)})...", flush=True)
        try:
            answers = _ask(chunk)
        except RateLimited as exc:
            print(f"  STOPPED: {exc}")
            raise
        if not answers:
            print("      no usable answer — leaving these rows unchanged")
            continue

        for row in chunk:
            a = answers.get(_norm(row.get("Company"))) or {}
            if not a or str(a.get("company") or "").startswith("<"):
                continue
            # The model's own confidence flag. False means it is guessing,
            # and a guessed phone number is exactly what this pipeline
            # spent a lot of effort removing.
            if a.get("published") is False:
                continue

            phones = clean_phones(a.get("office_no"))
            emails = clean_emails(a.get("office_email"))
            if phones:
                new = merge(row.get(PHONE), phones, is_email=False)
                if new != str(row.get(PHONE) or ""):
                    row[PHONE] = new
                    added_p += 1
            if emails:
                new = merge(row.get(EMAIL), emails, is_email=True)
                if new != str(row.get(EMAIL) or ""):
                    row[EMAIL] = new
                    added_e += 1

        # Checkpoint every chunk: this is a long run and an interruption
        # must not discard the answers already obtained.
        data["companies"] = companies
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                        encoding="utf-8")

    after_p = sum(1 for r in companies if str(r.get(PHONE) or "").strip())
    after_e = sum(1 for r in companies if str(r.get(EMAIL) or "").strip())
    multi = sum(1 for r in companies
                if ";" in str(r.get(PHONE) or "") or ";" in str(r.get(EMAIL) or ""))
    print(f"  phone {before_p} -> {after_p} | email {before_e} -> {after_e} "
          f"| {multi} rows now hold multiple values")
    return {"phone": after_p, "email": after_e}


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--market", action="append", default=[])
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--only-empty", action="store_true",
                    help="skip rows that already have BOTH fields")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--export", action="store_true")
    args = ap.parse_args()

    markets = list(args.market)
    if args.all:
        from run_all_markets import MARKETS  # noqa: PLC0415

        markets = list(MARKETS)
    if not markets:
        ap.error("pass --market or --all")

    stopped = False
    try:
        for market in markets:
            # Back up once per market, before any write.
            path = STATE_DIR / f"{market_slug(market)}.json"
            if path.is_file() and not args.dry_run:
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                shutil.copy2(path, path.with_suffix(
                    f".before-office-contacts-{stamp}.json"))
            try:
                process(market, dry_run=args.dry_run,
                        only_empty=args.only_empty, limit=args.limit)
            except RateLimited:
                stopped = True
                break
            print()
    finally:
        try:
            gam.shutdown()
        except Exception:
            pass

    if stopped:
        print("Run stopped early on Google's AI Mode quota. Everything "
              "obtained so far is saved; re-run the same command later to "
              "continue.")

    if args.export and not args.dry_run:
        from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

        for market in markets:
            print(f"re-exported: {export_xlsx(market).name}")


if __name__ == "__main__":
    main()
