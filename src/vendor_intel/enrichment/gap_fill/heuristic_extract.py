"""Regex/heuristic field extraction from search/scrape text — no LLM / OpenAI SDK."""
from __future__ import annotations

import re
from typing import Any


_FOUNDED_RE = re.compile(
    r"(?:founded|established|incorporated|started)\s+"
    r"(?:on\s+|in\s+)?"
    r"(?:[A-Z][a-z]+\.?\s+\d{1,2},?\s+)?"  # optional "Month Day, " (e.g. "April 4, ")
    r"(?:\d{1,2}\s+[A-Z][a-z]+\.?,?\s+)?"  # optional "Day Month, " (e.g. "4 April, ")
    r"((?:19|20)\d{2})\b",
    re.I,
)
_HQ_RE = re.compile(
    r"(?:headquartered\s+in|headquarters\s+(?:in|at|:)|based\s+in|hq\s*(?:in|:))\s*"
    r"([A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+){0,2}"
    r"(?:,\s*[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+){0,2}){0,2})",
)
_EMP_RE = re.compile(
    r"(?:(?:about|approximately|approx\.?|over|more than|~)\s*)?"
    r"([\d,]{2,7})\+?\s*(?:employees|staff|people|workers)\b",
    re.I,
)
_EMP_RANGE_RE = re.compile(
    r"(?:employees?|staff)\s*(?:of|:)?\s*([\d,]{1,6}\s*[-–—]\s*[\d,]{1,6})",
    re.I,
)
_REV_RE = re.compile(
    r"(?:revenue|turnover|sales)\s*(?:of|about|approximately|approx\.?|:)?\s*"
    r"(\$?\s*[\d]+(?:[.,]\d+)?\s*(?:[-–—]\s*\$?\s*[\d]+(?:[.,]\d+)?\s*)?"
    r"(?:billion|million|bn|mn|m|b)\b)",
    re.I,
)
_LINKEDIN_RE = re.compile(
    r"(https?://(?:www\.)?linkedin\.com/(?:company|in)/[A-Za-z0-9\-_%]+)",
    re.I,
)
_SALES_TITLE = (
    r"(?:Head of Sales|VP Sales|Vice President of Sales|Sales Director|"
    r"Director of Sales|Commercial Director|Chief Commercial Officer|"
    r"Chief Revenue Officer|Chief Sales Officer|Head of Procurement|"
    r"Procurement Director|Purchasing Director|Head of Purchasing|"
    r"National Sales|Channel Sales)"
)
_SALES_COMMA_RE = re.compile(
    rf"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){{1,3}})\s*,\s*{_SALES_TITLE}",
    re.I,
)
_SALES_REV_RE = re.compile(
    rf"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){{1,3}})\s*"
    rf"(?:is\s+)?(?:the\s+)?{_SALES_TITLE}",
    re.I,
)
_SALES_TITLE_RE = re.compile(_SALES_TITLE, re.I)

# --- Website / email / phone -------------------------------------------
# Google's Overview answers these three plainly ("The official website of X
# is minke.de", "Email: info@minke.de", "Phone: +49 203 28101-0") but the
# extractor had no patterns for them, so every such answer was discarded and
# the Website/Email/Office No. columns stayed near zero.

# Directories and social sites that appear as CITATIONS beside the answer —
# never the company's own site.
_NON_COMPANY_HOST = re.compile(
    r"(?:^|\.)(?:google|dnb|bloomberg|linkedin|facebook|twitter|x|instagram|"
    r"youtube|wikipedia|wikidata|crunchbase|kompass|europages|yumpu|zoominfo|"
    r"owler|glassdoor|indeed|amazon|ebay|maps|goo)\.",
    re.I,
)

# Anchored on the phrase Google actually uses, so a citation link elsewhere
# in the blob cannot win.
_WEBSITE_RE = re.compile(
    r"(?:official\s+website|website|homepage|web\s?site)\s*"
    # The company name may sit between the phrase and the domain ("website
    # of Fritz Minke GmbH & Co. KG is minke.de"). It can contain periods
    # ("Co. KG", "S.p.A."), so this span must allow them while still
    # stopping at a line break.
    r"(?:(?:of|for)\s+[^:\n]{0,80}?)?"
    r"(?:\s+is|:|\s+at)\s+"
    r"(?:https?://)?(?:www\.)?"
    r"([a-z0-9][a-z0-9\-]{1,62}(?:\.[a-z0-9\-]{2,24}){1,3})",
    re.I,
)

_EMAIL_RE = re.compile(
    r"\b([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,24})\b"
)
# Placeholder/tracking addresses that carry no contact value.
_JUNK_EMAIL = re.compile(
    r"(?:^|@)(?:example|test|sentry|wixpress|domain|email|yourcompany)\.|"
    r"^(?:no-?reply|do-?not-?reply|postmaster|abuse|webmaster)@|"
    r"\.(?:png|jpg|jpeg|gif|webp|svg)$",
    re.I,
)
# A generic inbox is the right value for Office Email; a named one is a
# personal contact. Both are useful, so record which is which.
_GENERIC_LOCAL = re.compile(
    r"^(?:info|office|contact|kontakt|mail|sales|vertrieb|hello|enquiry|"
    r"enquiries|inquiry|support|service|admin|general)@",
    re.I,
)

_PHONE_RE = re.compile(
    r"(?:phone|tel|telephone|call|fon|contact\s+number)\s*(?:number)?\s*[:.]?\s*"
    r"(\+?\d[\d\s().\-/]{6,24}\d)",
    re.I,
)


def _clean_phone(raw: str) -> str:
    """Keep a phone only when it has enough digits to be dialable."""
    text = re.sub(r"\s+", " ", raw or "").strip(" .,;:-/")
    digits = re.sub(r"\D", "", text)
    # 8 digits is the shortest realistic national number; 15 is the E.164 max
    # (allow a little slack for extensions written inline).
    if not (8 <= len(digits) <= 18):
        return ""
    return text


def heuristic_extract_fields(text: str, *, company: str = "") -> dict[str, Any]:
    """Pull firmographic fields from raw markdown/snippets without an LLM."""
    if not (text or "").strip():
        return {}
    blob = text[:12000]
    out: dict[str, Any] = {}

    m = _FOUNDED_RE.search(blob)
    if m:
        out["founded_year"] = m.group(1)

    m = _HQ_RE.search(blob)
    if m:
        hq = re.sub(r"\s+", " ", m.group(1)).strip(" .,;:")
        # Drop trailing junk words that sometimes leak past the city/country
        stop = {"it", "has", "with", "and", "the", "a", "an", "is", "was", "of"}
        parts = [p for p in re.split(r",\s*|\s+", hq) if p]
        cleaned: list[str] = []
        for p in parts:
            tok = p.strip(" .,;:")
            if not tok or tok.lower() in stop:
                break
            cleaned.append(tok)
        hq = ", ".join(cleaned) if cleaned else hq.strip(" .,;:")
        if len(hq) >= 3:
            out["headquarters"] = hq

    m = _EMP_RANGE_RE.search(blob)
    if m:
        out["employee_count"] = re.sub(r"\s+", "", m.group(1).replace("—", "-").replace("–", "-"))
    else:
        m = _EMP_RE.search(blob)
        if m:
            out["employee_count"] = m.group(1).replace(",", "")

    m = _REV_RE.search(blob)
    if m:
        out["revenue"] = re.sub(r"\s+", " ", m.group(1)).strip()

    m = _LINKEDIN_RE.search(blob)
    if m:
        out["linkedin_url"] = m.group(1).rstrip("/")

    m = _WEBSITE_RE.search(blob)
    if m:
        host = m.group(1).lower().strip(" .,;:/")
        if "." in host and not _NON_COMPANY_HOST.search(host):
            out["website"] = host

    # Prefer an email on the company's own domain; a citation page can leak a
    # third-party address into the same blob.
    site_host = str(out.get("website") or "").lower()
    emails: list[str] = []
    for raw in _EMAIL_RE.findall(blob):
        addr = raw.strip().lower()
        if _JUNK_EMAIL.search(addr):
            continue
        host = addr.split("@", 1)[1]
        if _NON_COMPANY_HOST.search(host):
            continue
        emails.append(addr)
    if emails:
        if site_host:
            same = [e for e in emails if e.split("@", 1)[1].endswith(site_host)]
            emails = same or emails
        generic = [e for e in emails if _GENERIC_LOCAL.search(e)]
        named = [e for e in emails if not _GENERIC_LOCAL.search(e)]
        if generic:
            out["office_email"] = generic[0]
        if named:
            out["email"] = named[0]
        elif generic:
            out["email"] = generic[0]

    m = _PHONE_RE.search(blob)
    if m:
        phone = _clean_phone(m.group(1))
        if phone:
            out["phone"] = phone

    m = _SALES_COMMA_RE.search(blob) or _SALES_REV_RE.search(blob)
    if m:
        person = m.group(1).strip().rstrip(".")
        if company and person.lower() == company.lower():
            pass
        elif len(person.split()) >= 2:
            out["contact_person"] = person
            role_m = _SALES_TITLE_RE.search(blob[max(0, m.start() - 10) : m.end() + 80])
            if role_m:
                out["contact_role"] = role_m.group(0)

    return out
