"""Re-ask Google AI Mode for the six contact fields, real data only.

WHY A SECOND PASS. The discovery prompt originally required all 21 columns
to be non-blank, which forced the model to construct contact details when
it did not know them: 94% of exported Office Email values came out as
info@/sales@ at the company's own domain, LinkedIn URLs were built from
the contact's name, and phone numbers were generated from the HQ country's
calling code. The discovery prompt has since been changed to demand ""
instead of a guess, but the already-collected rows still carry the
constructed values.

This module re-queries AI Mode with ONE focused question per batch of
companies and an explicit permission to answer "unknown" — the thing the
original prompt forbade. Fields it declines are then blanked.

AN HONEST LIMITATION, stated plainly because it affects how much the
result is worth: AI Mode is the same source that produced the guesses, so
this is not independent verification. It can repeat a plausible-looking
invention. Two things still make it useful:

  * the prompt now REWARDS "unknown" instead of penalising it, so a
    field the model never actually knew tends to come back empty; and
  * a short, single-purpose question about one company grounds better in
    AI Mode's web search than a 20-company request for 19 fields each.

For genuinely independent confirmation, verify_contact_fields.py fetches
the companies' own websites over HTTP and is the stronger check. Use this
module to shrink the fabricated set, that one to confirm what remains.
"""
from __future__ import annotations

import json
import re
import sys
from typing import Any

from vendor_intel.pipeline.market_research import ai_mode_client

# 15 companies per query, per explicit user instruction. Also about the
# largest batch that still leaves room for a real answer: the reply must
# carry 6 fields x 15 companies and AI Mode truncates long answers, while
# the request itself has to stay under the ~7.8 KB search-URL ceiling.
BATCH_SIZE = 15

_SYSTEM = (
    "You are a precise B2B contact researcher. You report ONLY real, "
    "published contact details that you actually know. You never invent, "
    "construct, pattern-generate or guess a person's name, an email "
    "address, a phone number or a LinkedIn URL. An empty string is the "
    "correct and expected answer whenever you do not genuinely know a "
    "value."
)

_PROMPT = """For each company below, give ONLY its REAL, PUBLISHED contact details.

CRITICAL — DO NOT HALLUCINATE. DO NOT RANDOM-FILL:
- Return "" (empty string) for anything you do not actually know.
- NEVER build an email from the domain. "info@<their-domain>" invented \
from the website is FORBIDDEN unless that exact address is genuinely \
published by the company.
- NEVER build a LinkedIn URL from a person's name. Only a profile URL you \
actually know.
- NEVER invent a phone number, and never generate one from a country \
calling code.
- NEVER invent a person's name or job title.
- An empty field is CORRECT and preferred. A wrong contact detail is much \
worse than an empty one. Do not try to fill every field.

WHICH PERSON TO NAME (in this priority order):
1. FIRST choice - Procurement Manager, Purchasing Manager, Head of
Procurement, VP of Sales, Head of Sales, Sales Director, or the
equivalent commercial/buying lead.
2. ONLY IF no such person is known - any C-suite officer (CEO,
Managing Director, CTO, COO, CFO) or the company's Representative.
Still return "" if you know NEITHER. The priority is about WHICH
known person to pick, never a reason to invent one.

For each company return these six fields:
- contact_person: the real named person chosen by the priority above, or ""
- role: that person's real job title, or ""
- email: a real published company/sales email address, or ""
- linkedin: that person's real personal LinkedIn profile URL \
(linkedin.com/in/...), or ""
- office_no: a real published telephone number, or ""
- office_email: a real published office/contact email, or ""

Companies:
{companies}

Reply with ONLY a JSON array, no markdown fences and no commentary, one \
object per company IN THE SAME ORDER:
[{{"company": "", "contact_person": "", "role": "", "email": "", \
"linkedin": "", "office_no": "", "office_email": ""}}]"""

_CONTACT_FIELDS = (
    "contact_person", "role", "email", "linkedin", "office_no", "office_email",
)

# Values that mean "I don't know" however the model phrases it. Kept here
# rather than reusing schema.is_placeholder because the model also answers
# in prose ("not publicly available", "no information found").
_UNKNOWN_RE = re.compile(
    r"^\s*(|-|n/?a|na|none|null|nil|unknown|not\s+(known|available|found|"
    r"published|disclosed|provided|listed|specified)|no\s+(data|info|"
    r"information|email|phone|contact).*|tbd|undisclosed|not\s+applicable)"
    r"\s*\.?\s*$",
    re.I,
)


def _log(msg: str) -> None:
    """Print a progress line, never crashing the run over an encoding issue.

    A Windows cp1252 console cannot encode most non-Western characters, and
    print() RAISES on an unmappable one — which killed a run mid-batch on a
    Turkish company name containing "S" with cedilla (U+015E). Since this
    pipeline deliberately searches worldwide, non-ASCII names are normal
    input, so logging must degrade rather than abort.
    """
    text = f"  [contact-lookup] {msg}"
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        enc = (getattr(sys.stdout, "encoding", None) or "ascii")
        print(text.encode(enc, errors="replace").decode(enc), flush=True)


def is_unknown(value: object) -> bool:
    """True when the model effectively said "I don't know"."""
    return bool(_UNKNOWN_RE.match(str(value or "")))


def _clean(value: object) -> str:
    """Normalise one answer value, mapping any "unknown" phrasing to ""."""
    text = " ".join(str(value or "").split())
    return "" if is_unknown(text) else text


# An email that is just a generic mailbox at the company's own domain is
# exactly what the original prompt was told to construct, so it cannot be
# distinguished from an invention. Rejected here unless the model returns
# something more specific.
_GENERIC_LOCAL = frozenset({
    "info", "sales", "contact", "enquiry", "enquiries", "inquiry",
    "inquiries", "support", "hello", "office", "admin", "mail", "general",
    "reception", "marketing",
})

_LINKEDIN_IN_RE = re.compile(r"linkedin\.com/in/[A-Za-z0-9\-_%]+", re.I)


def _looks_constructed_email(email: str, website: str) -> bool:
    """True when `email` is a generic mailbox at the company's own domain.

    This is the exact shape the discovery prompt used to fabricate
    ("website acme.com -> info@acme.com"), so it is treated as unverified
    rather than accepted. A real, specific address (j.smith@acme.com) or a
    generic one at a DIFFERENT domain is not caught by this.
    """
    if not email or "@" not in email:
        return False
    local, _, domain = email.lower().partition("@")
    site = re.sub(r"^https?://", "", str(website or "").lower()).split("/")[0]
    site = site.removeprefix("www.")
    if not site or not domain:
        return False
    same_domain = domain == site or domain.endswith("." + site)
    return same_domain and local in _GENERIC_LOCAL


def _sanitize(answer: dict[str, Any], website: str, *, strict: bool) -> dict[str, str]:
    """Keep only values that survive the anti-fabrication checks."""
    out = {f: _clean(answer.get(f)) for f in _CONTACT_FIELDS}

    # LinkedIn must be a real personal profile URL, never a company page
    # and never a name-derived slug.
    match = _LINKEDIN_IN_RE.search(out["linkedin"])
    if match:
        url = match.group(0)
        out["linkedin"] = url if url.lower().startswith("http") else f"https://www.{url}"
    else:
        out["linkedin"] = ""

    # A role without a person (or vice versa) is not a usable contact.
    if not out["contact_person"] or not out["role"]:
        out["contact_person"] = ""
        out["role"] = ""
    # A LinkedIn profile with no named person cannot be attributed.
    if not out["contact_person"]:
        out["linkedin"] = ""

    if strict:
        for field in ("email", "office_email"):
            if _looks_constructed_email(out[field], website):
                out[field] = ""

    for field in ("email", "office_email"):
        if out[field] and not re.match(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$", out[field]):
            out[field] = ""  # not a syntactically valid address

    # A phone number needs enough digits to be a real number.
    if len(re.sub(r"\D", "", out["office_no"])) < 7:
        out["office_no"] = ""

    return out


def _format_companies(rows: list[dict[str, str]]) -> str:
    lines = []
    for n, row in enumerate(rows, 1):
        name = row.get("Company", "")
        site = row.get("Website", "")
        hq = row.get("Headquarters", "")
        lines.append(f"{n}. {name} — website {site} — HQ {hq}")
    return "\n".join(lines)


def _extract_json_array(raw: str) -> list[dict[str, Any]]:
    """Pull the JSON array out of AI Mode's rendered answer text."""
    text = (raw or "").replace("Use code with caution.", "").strip()

    # AI Mode echoes the question, and this prompt ENDS with an empty JSON
    # template — so a naive find("[") locks onto the echoed template and
    # yields zero results. Take the LAST array that starts with a real
    # "company" key. (Same failure mode already fixed in discovery.py.)
    starts = [m.start() for m in re.finditer(r'\[\s*\{\s*"company"', text)]
    if starts:
        text = text[starts[-1]:]

    start, end = text.find("["), text.rfind("]")
    if start < 0 or end <= start:
        return []
    candidate = text[start : end + 1]
    for attempt in (candidate, re.sub(r"\s*\n\s*", " ", candidate)):
        try:
            data = json.loads(attempt)
        except (json.JSONDecodeError, ValueError):
            continue
        if isinstance(data, list):
            return [d for d in data if isinstance(d, dict)]
    return []


def lookup_contacts(
    rows: list[dict[str, str]], *, strict: bool = True
) -> dict[str, dict[str, str]]:
    """Ask AI Mode for the real contact details of one batch of companies.

    Returns {company_name: {field: value}} carrying only values that
    survived sanitisation. A company the model said nothing usable about
    maps to all-empty strings, which the caller treats as "blank it".

    `strict` (default) additionally rejects generic info@/sales@ addresses
    at the company's own domain, because that is precisely the pattern the
    original prompt was instructed to construct.
    """
    if not rows:
        return {}

    prompt = _PROMPT.format(companies=_format_companies(rows))
    try:
        raw = ai_mode_client.chat(_SYSTEM, prompt)
    except Exception as exc:  # CAPTCHA, timeout, refusal
        _log(f"AI Mode unavailable for this batch ({type(exc).__name__}) — leaving unchanged")
        return {}

    answers = _extract_json_array(raw)
    if not answers:
        _log("no parseable answer for this batch — leaving unchanged")
        return {}

    # Match answers back to companies by name where possible, falling back
    # to position. Never trust position alone: the model can drop or
    # reorder entries, and a misaligned answer would write one company's
    # contact details onto another.
    by_name = {str(a.get("company", "")).strip().casefold(): a for a in answers}
    out: dict[str, dict[str, str]] = {}
    for n, row in enumerate(rows):
        name = row.get("Company", "")
        answer = by_name.get(name.strip().casefold())
        if answer is None and len(answers) == len(rows):
            answer = answers[n]  # safe only when counts line up exactly
        if answer is None:
            continue
        out[name] = _sanitize(answer, row.get("Website", ""), strict=strict)
    return out
