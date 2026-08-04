"""Regex/heuristic field extraction from search/scrape text — no LLM / OpenAI SDK."""
from __future__ import annotations

import re
from typing import Any


_FOUNDED_RE = re.compile(
    r"(?:founded|established|incorporated|started)\s+(?:in\s+)?((?:19|20)\d{2})\b",
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
_CEO_RE = re.compile(
    r"(?:CEO|Chief Executive Officer|Founder|Co-Founder|Managing Director|"
    r"President)\s*[:\-]?\s*"
    r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})",
)
_CEO_REV_RE = re.compile(
    r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})\s*"
    r"(?:is\s+)?(?:the\s+)?(?:CEO|Chief Executive Officer|Founder|Managing Director)",
)
# ZoomInfo / RocketReach style: "Philip Freeman, Chief Executive Officer & Founder"
_CEO_COMMA_RE = re.compile(
    r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})\s*,\s*"
    r"(?:Chief Executive Officer|CEO|Founder|Co-Founder|Managing Director|President)"
    r"(?:\s*[&/]\s*(?:Founder|CEO|Co-Founder))?",
    re.I,
)
# "Phil Freeman (CEO and Founder)"
_CEO_PAREN_RE = re.compile(
    r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})\s*\(\s*"
    r"(?:CEO|Chief Executive Officer|Founder)(?:\s+and\s+(?:Founder|CEO))?\s*\)",
    re.I,
)


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

    m = (
        _CEO_COMMA_RE.search(blob)
        or _CEO_PAREN_RE.search(blob)
        or _CEO_RE.search(blob)
        or _CEO_REV_RE.search(blob)
    )
    if m:
        person = m.group(1).strip().rstrip(".")
        if company and person.lower() == company.lower():
            pass
        elif len(person.split()) >= 2:
            out["contact_person"] = person
            role_m = re.search(
                r"(CEO|Chief Executive Officer|Founder|Co-Founder|Managing Director|President)",
                blob[max(0, m.start() - 10) : m.end() + 60],
                re.I,
            )
            if role_m:
                out["contact_role"] = role_m.group(1)

    return out
