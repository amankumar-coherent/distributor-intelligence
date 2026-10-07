"""Website reachability + identity verification, with ChatGPT-scraper-based
domain correction for dead/wrong websites.

Problem this solves: a discovered "Website" value can be (a) genuinely dead
(DNS failure, timeout, parked domain, 404), or (b) reachable but wrong — the
classic same-company-name trap, e.g. "Acme Inc" the packaging distributor in
Ohio vs. "Acme Inc" the unrelated software company in California that
happens to rank first in a generic search. Neither case is caught by a
plain "does this URL respond" check alone.

Flow, reusing existing primitives (no duplicate HTTP/scraper code):
  1. Reachability  -> vendor_intel.scraping.httpx_fetch.fetch_page_httpx
                      (fast, synchronous, no browser needed)
  2. Identity check -> does the fetched page's title/text actually mention
                       the company name? A reachable-but-wrong domain fails
                       here even though it's "alive".
  3. If (1) or (2) fails -> vendor_intel.scraping.chatgpt_scraper.ask_chatgpt
                       is asked for the official website, given company name
                       + headquarters (the disambiguating context spec asks
                       for, since many companies share a name).
  4. The ChatGPT-proposed replacement is itself re-verified through steps
     1-2 before being trusted — never accept an AI-proposed URL blind, same
     rule already enforced for Kompass page discovery.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from urllib.parse import urlparse

NOT_DISCLOSED = "Not publicly disclosed"


def _log(msg: str) -> None:
    print(f"  [website-verify] {msg}", flush=True)


def _detail(msg: str) -> None:
    print(f"    -> {msg}", flush=True)


def website_verification_enabled() -> bool:
    return (os.getenv("WEBSITE_VERIFICATION_ENABLED", "true") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


@dataclass
class WebsiteVerificationResult:
    website: str
    is_reachable: bool
    is_identity_match: bool
    was_corrected: bool
    source: str  # "unchanged" | "chatgpt_scraper" | "none_found"
    reason: str = ""


def _normalize_url(url: str) -> str:
    u = (url or "").strip()
    if not u or u.lower() in {"not publicly disclosed", "n/a", "na", "unknown", "none"}:
        return ""
    if "://" not in u:
        u = f"https://{u}"
    return u


def _domain_tokens(url: str) -> set[str]:
    try:
        host = urlparse(url).netloc.lower()
    except Exception:
        return set()
    host = host[4:] if host.startswith("www.") else host
    # Strip the TLD/suffix chain, split remaining on non-alnum.
    base = host.split(".")[0] if host else ""
    return {t for t in re.split(r"[^a-z0-9]+", base) if len(t) > 2}


def _name_tokens(company_name: str) -> set[str]:
    """Significant tokens from a company name, dropping generic corporate
    suffixes that would otherwise cause false identity matches (e.g. two
    unrelated companies both being "... Inc" or "... Group")."""
    stop = {
        "inc", "incorporated", "ltd", "limited", "llc", "llp", "co", "corp",
        "corporation", "company", "group", "holdings", "gmbh", "sa", "srl",
        "bv", "plc", "pvt", "private", "the", "and", "of",
    }
    tokens = re.split(r"[^A-Za-z0-9]+", company_name.lower())
    return {t for t in tokens if t and t not in stop and len(t) > 2}


def _looks_like_identity_match(company_name: str, url: str, page_title: str, page_text: str) -> bool:
    """Heuristic, source-agnostic identity check: does the page actually
    look like it belongs to this company, not just some other business at a
    similarly-named domain? Checks the company name's significant tokens
    against the domain itself, the page title, and (as a weaker signal) the
    visible text — any one strong match is accepted; this is intentionally
    lenient (false negatives are cheaper than losing a real match) but still
    catches the common case of a domain with a completely different name."""
    name_tokens = _name_tokens(company_name)
    if not name_tokens:
        return True  # nothing to check against — don't block on this alone

    domain_tokens = _domain_tokens(url)
    if domain_tokens and (name_tokens & domain_tokens):
        return True
    # Domains often glue words together with no separator (e.g.
    # "acmeelectronics.com" for "Acme Electronics") — _domain_tokens can't
    # split that back into ["acme", "electronics"], so also check whether a
    # name token appears as a substring of the (unsplit) domain base.
    if domain_tokens and any(
        any(name_tok in dom_tok for dom_tok in domain_tokens) for name_tok in name_tokens
    ):
        return True

    title_l = (page_title or "").lower()
    if title_l and any(t in title_l for t in name_tokens):
        return True

    text_l = (page_text or "").lower()[:2000]
    # Require at least half the significant name tokens to appear in the
    # opening page text — a single generic word match isn't enough evidence.
    if name_tokens:
        hits = sum(1 for t in name_tokens if t in text_l)
        if hits >= max(1, len(name_tokens) // 2 + 1):
            return True

    return False


_SSL_ERROR_MARKERS = ("ssl", "certificate", "cert_verify", "hostname mismatch", "tlsv1")


def _looks_like_ssl_error(error: str) -> bool:
    e = (error or "").lower()
    return any(marker in e for marker in _SSL_ERROR_MARKERS)


def _fetch_ignoring_ssl(url: str, *, timeout: float = 15.0):
    """Fallback for a real, working site with a broken/mismatched TLS cert —
    a common misconfiguration (wrong cert on a shared host, expired
    intermediate, etc.) that does NOT mean the company/domain is dead. Only
    used after the normal verified-TLS fetch fails specifically with an
    SSL-shaped error — never used as the first attempt, so this doesn't
    weaken verification for the common "actually dead domain" case."""
    import httpx as _httpx

    from vendor_intel.scraping.httpx_fetch import _DEFAULT_HEADERS, _extract_text

    try:
        with _httpx.Client(
            follow_redirects=True, timeout=timeout, headers=_DEFAULT_HEADERS, verify=False
        ) as client:
            resp = client.get(url)
            if resp.status_code >= 400:
                return None
            html = resp.text or ""
            text = _extract_text(html)
            from vendor_intel.scraping.page_result import PageFetchResult

            return PageFetchResult(
                url=url, final_url=str(resp.url), alive=len(text.strip()) >= 80,
                text=text, html=html[:8000], source="httpx-no-verify",
            )
    except Exception:
        return None


def check_website_identity(company_name: str, website: str) -> tuple[bool, bool, str]:
    """Returns (is_reachable, is_identity_match, reason). Never raises —
    any fetch failure is treated as not-reachable, not an error to propagate
    (one bad website must never stop the pipeline).

    A fetch that fails specifically due to an SSL/certificate error (not a
    DNS/connection failure) is retried once over an unverified-TLS
    connection before being treated as dead — a mismatched or expired cert
    is a real but common misconfiguration on an otherwise live, legitimate
    site, and must not by itself trigger a "find a different domain" search."""
    url = _normalize_url(website)
    if not url:
        return False, False, "no website to check"

    try:
        from vendor_intel.scraping.httpx_fetch import fetch_page_httpx

        result = fetch_page_httpx(url)
    except Exception as exc:  # noqa: BLE001
        return False, False, f"fetch error: {type(exc).__name__}: {exc}"

    if not result.alive:
        if _looks_like_ssl_error(result.error):
            retry = _fetch_ignoring_ssl(url)
            if retry is not None and retry.alive:
                _detail(f"SSL cert issue on {url} but site is live — accepted anyway")
                result = retry
            else:
                return False, False, result.error or "unreachable (ssl)"
        else:
            return False, False, result.error or "unreachable"

    match = _looks_like_identity_match(company_name, result.final_url or url, result.title, result.text)
    if not match:
        return True, False, f"reachable but does not appear to belong to {company_name!r}"
    return True, True, ""


_WEBSITE_SEARCH_PROMPT = """Company: {company}
Headquarters/location: {hq}

Find the OFFICIAL website/domain for this SPECIFIC company (not a different
company that happens to share the same or a similar name — use the
headquarters/location above to disambiguate, since many companies share a
name). Reply with a single JSON object, no markdown, no commentary:
{{"website": "<the official https:// URL, or empty string if you cannot
confirm one>", "confidence": <0.0-1.0>}}
Rules:
- Only return a URL you are reasonably confident is this exact company's
  real, official site — never a plausible-looking guess.
- If you cannot confirm which company this is or find no official site,
  return an empty string and low confidence rather than guessing."""


def _find_website_via_chatgpt_scraper(company_name: str, headquarters: str) -> str:
    """Ask the free ChatGPT web-UI scraper for the company's official
    website, using name + HQ to disambiguate from same-named companies.
    Returns "" on any failure — caller leaves the website as-is/empty rather
    than propagate an error."""
    from vendor_intel.scraping.google_ai_mode import ask_generative, generative_backend_enabled

    if not generative_backend_enabled():
        _detail("web LLM: skipped (no generative web backend enabled)")
        return ""

    hq = headquarters.strip() if headquarters and not _is_missing_hq(headquarters) else "unknown"
    prompt = _WEBSITE_SEARCH_PROMPT.format(company=company_name, hq=hq)
    try:
        reply = ask_generative(prompt)
    except Exception as exc:  # noqa: BLE001
        _detail(f"web LLM error: {type(exc).__name__}: {exc}")
        return ""

    import json

    text = (reply or "").strip()
    if text.startswith("```"):
        text = text.strip("`").removeprefix("json").strip()
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        return ""
    try:
        data = json.loads(text[start : end + 1])
    except (json.JSONDecodeError, ValueError):
        return ""
    return str((data or {}).get("website") or "").strip()


def _is_missing_hq(value: str) -> bool:
    v = (value or "").strip().lower()
    return not v or v in {"n/a", "na", "unknown", "none", "not publicly disclosed"}


def verify_and_fix_website(
    company_name: str,
    headquarters: str,
    current_website: str,
    *,
    max_attempts: int = 2,
) -> WebsiteVerificationResult:
    """Full check-and-repair flow for one company's website. Never raises —
    any failure at any step just means the website stays as its last known
    (possibly still empty/unverified) value, per the pipeline's "one failed
    source must never stop the whole run" rule."""
    if not website_verification_enabled():
        return WebsiteVerificationResult(
            website=current_website, is_reachable=False, is_identity_match=False,
            was_corrected=False, source="unchanged", reason="verification disabled",
        )

    current = _normalize_url(current_website)
    if current:
        reachable, identity_ok, reason = check_website_identity(company_name, current)
        if reachable and identity_ok:
            _detail(f"website OK: {current}")
            return WebsiteVerificationResult(
                website=current, is_reachable=True, is_identity_match=True,
                was_corrected=False, source="unchanged",
            )
        _detail(f"website FAILED ({reason}): {current} — searching for correct domain")
    else:
        _detail("no website on file — searching for official domain")

    tried: set[str] = {current} if current else set()
    for attempt in range(max_attempts):
        candidate = _find_website_via_chatgpt_scraper(company_name, headquarters)
        candidate = _normalize_url(candidate)
        if not candidate or candidate in tried:
            break
        tried.add(candidate)

        reachable, identity_ok, reason = check_website_identity(company_name, candidate)
        if reachable and identity_ok:
            _detail(f"corrected website (attempt {attempt + 1}): {candidate}")
            return WebsiteVerificationResult(
                website=candidate, is_reachable=True, is_identity_match=True,
                was_corrected=True, source="chatgpt_scraper",
            )
        _detail(f"candidate rejected ({reason}): {candidate}")

    # Nothing verifiable found — keep the old value only if it was at least
    # reachable (better than nothing); otherwise report empty rather than
    # keep a confirmed-dead domain in the export.
    if current:
        reachable, _, _ = check_website_identity(company_name, current)
        if reachable:
            return WebsiteVerificationResult(
                website=current, is_reachable=True, is_identity_match=False,
                was_corrected=False, source="unchanged",
                reason="no better match found; keeping reachable-but-unconfirmed original",
            )
    return WebsiteVerificationResult(
        website="", is_reachable=False, is_identity_match=False,
        was_corrected=bool(current), source="none_found",
        reason="no verifiable website found",
    )
