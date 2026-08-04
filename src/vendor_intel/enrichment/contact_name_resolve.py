"""Resolve full contact person names without Apollo enrich (no email).

Waterfall for partial Apollo names (Pe***o) or missing contact_person:
  1) LinkedIn company search / employees (MCP)
  2) Company leadership / about page
  3) Google / web search (LinkedIn profile snippets)
  4) RocketReach (API if ROCKETREACH_API_KEY set, else site:rocketreach.co search)
"""
from __future__ import annotations

import json
import os
import re
from typing import Any

from vendor_intel.enrichment.gap_fill.gaps import (
    domain_from_row,
    has_contact_person,
    merge_apollo_contact,
    norm,
)
from vendor_intel.enrichment.gap_fill.sources import _search_snippets

_PARTIAL = re.compile(r"\*")
_NAME_IN_SNIPPET = re.compile(
    r"(?!(?:LinkedIn|Profile|Sign)\b)"
    r"([A-Z][a-z]+(?:\s+[A-Z][a-z'.-]+){0,3})\s*[-–—|,]\s*"
    r"(?:CISO|CEO|Chief|Founder|President|Managing Director|VP|Head of Security|Director)",
    re.I,
)
_LINKEDIN_IN = re.compile(r"linkedin\.com/in/([a-zA-Z0-9_-]+)", re.I)
_LINKEDIN_PREFIX = re.compile(
    r"^(?:linkedin|view(?:\s+\w+){0,4}\s+on\s+linkedin|sign\s+in(?:\s+to\s+linkedin)?)\s*",
    re.I,
)
_GLUED_LINKEDIN = re.compile(r"^LinkedIn(?=[A-Z][a-z])", re.I)
_JUNK_NAME_TOKENS = frozenset(
    {
        "linkedin",
        "profile",
        "sign",
        "login",
        "connect",
        "follow",
        "view",
        "official",
        "website",
        # Sentence / snippet fragments mistaken for names ("He is the co")
        "he",
        "she",
        "they",
        "him",
        "her",
        "his",
        "their",
        "is",
        "are",
        "was",
        "were",
        "the",
        "a",
        "an",
        "of",
        "and",
        "or",
        "at",
        "as",
        "to",
        "for",
        "with",
        "who",
        "whom",
        "that",
        "this",
        "co",
        "new",
    }
)
# Job titles alone are not person names (e.g. google_search -> 'CFO').
_ROLE_ONLY_NAMES = frozenset(
    {
        "ceo",
        "cfo",
        "cto",
        "coo",
        "ciso",
        "cio",
        "cmo",
        "founder",
        "co-founder",
        "cofounder",
        "president",
        "director",
        "manager",
        "vp",
        "svp",
        "evp",
        "chairman",
        "chairwoman",
        "owner",
        "partner",
    }
)
# Company names that are also industry keywords — require employer-at-company proof.
_GENERIC_COMPANY_NAMES = frozenset(
    {
        "smartwatch",
        "smartwatches",
        "health",
        "fitness",
        "technology",
        "technologies",
        "solutions",
        "digital",
        "global",
        "systems",
        "services",
        "group",
        "international",
    }
)
_EMPLOYER_AT = re.compile(
    r"(?:founder|co-founder|ceo|cto|ciso|director|president|vp|head)\s+(?:&\s*)?"
    r"(?:co-founder|ceo|cto)?\s*(?:at|@)\s+([A-Za-z][\w\s&'.-]{2,60})",
    re.I,
)


def sanitize_contact_person(name: str) -> str:
    """Strip search-snippet junk glued to person names (e.g. LinkedInJay Chaudhry)."""
    s = norm(name)
    if not s:
        return ""
    s = _GLUED_LINKEDIN.sub("", s)
    s = _LINKEDIN_PREFIX.sub("", s).strip()
    s = re.sub(r"\s*\|\s*LinkedIn.*$", "", s, flags=re.I).strip()
    s = re.sub(r"\s+on\s+LinkedIn\s*$", "", s, flags=re.I).strip()
    return re.sub(r"\s+", " ", s).strip()


def is_plausible_person_name(name: str) -> bool:
    s = sanitize_contact_person(name)
    if len(s) < 3 or is_partial_name(s):
        return False
    if "linkedin" in s.lower():
        return False
    # Reject sentence fragments cut from bios ("He is the co-founder…").
    if re.search(r"\b(?:he|she|they)\s+is\b", s, re.I):
        return False
    if re.search(r"\bis\s+the\s+(?:co[- ]?)?(?:founder|ceo|president|owner)\b", s, re.I):
        return False
    parts = s.split()
    if not 1 <= len(parts) <= 5:
        return False
    low_parts = [re.sub(r"[^a-z\-]", "", p.lower()) for p in parts]
    if len(parts) == 1 and low_parts[0] in _ROLE_ONLY_NAMES:
        return False
    # Real names are proper nouns — reject if any token is a sentence stopword.
    if any(p in _JUNK_NAME_TOKENS for p in low_parts):
        return False
    for part in parts:
        if not re.match(r"^[A-Za-z]", part):
            return False
    return True


def _company_context_tokens(company: str, domain: str) -> set[str]:
    tokens: set[str] = set()
    for w in re.findall(r"[a-z0-9]{4,}", (company or "").lower()):
        if w not in ("inc", "ltd", "corp", "company", "holdings", "group", "global"):
            tokens.add(w)
    dom = (domain or "").lower().split("/")[0].split(".")[0]
    if dom.startswith("www."):
        dom = dom[4:]
    if len(dom) >= 4:
        tokens.add(dom.replace("-", ""))
    return tokens


def _company_in_employment_context(blob: str, token: str) -> bool:
    """True when token appears as current employer, not only in a school/workshop list."""
    return bool(
        re.search(
            rf"(?:at|@)\s*{re.escape(token)}\b"
            rf"|{re.escape(token)}\s*[|\-–—]"
            rf"|{re.escape(token)}\s+(?:inc|ltd|corp|group|technologies)\b",
            blob,
            re.I,
        )
    )


def _partial_obfuscated_prefix(name: str) -> str:
    """Apollo partial like Ru***o -> 'ru'."""
    m = re.search(r"(\w{2,})\*+", name or "")
    if m:
        return m.group(1).lower()
    parts = (name or "").split()
    return parts[-1][:2].lower() if parts else ""


def partial_upgrade_compatible(partial: str, full: str) -> bool:
    """Reject upgrades where the full last name does not match the obfuscated Apollo hint."""
    if not is_partial_name(partial):
        return True
    prefix = _partial_obfuscated_prefix(partial)
    if len(prefix) < 2:
        return True
    full_parts = sanitize_contact_person(full).split()
    if not full_parts:
        return False
    last = full_parts[-1].lower()
    # Match Ru***o -> Russo; reject Ru***o -> De Vita (last token 'vita').
    return last.startswith(prefix[:2]) or prefix.startswith(last[:2])


def _is_generic_company_name(company: str) -> bool:
    w = re.sub(r"[^a-z0-9]+", " ", (company or "").lower()).strip()
    if not w:
        return False
    if w in _GENERIC_COMPANY_NAMES:
        return True
    parts = w.split()
    return len(parts) == 1 and parts[0] in _GENERIC_COMPANY_NAMES


def _domain_context_in_blob(domain: str, blob: str) -> bool:
    dom = (domain or "").lower().split("/")[0].removeprefix("www.")
    if not dom:
        return False
    base = dom.split(".")[0]
    return dom in blob or base in blob


def _employment_search_tokens(company: str, domain: str, tokens: set[str]) -> set[str]:
    """Extra employer names that are the same org (e.g. Noise ↔ Gonoise)."""
    extra: set[str] = set()
    blob_company = (company or "").lower()
    dom = (domain or "").lower().split("/")[0].removeprefix("www.").split(".")[0]
    if "gonoise" in blob_company or dom == "gonoise" or blob_company == "noise":
        extra.update({"noise", "gonoise"})
    return tokens | extra


def _employer_matches_company(employer: str, company: str, tokens: set[str]) -> bool:
    employer_low = employer.strip().lower()
    company_low = (company or "").lower()
    if not employer_low:
        return False
    if company_low in employer_low or employer_low in company_low:
        return True
    if any(t in employer_low or employer_low in t for t in tokens):
        return True
    if "noise" in employer_low and ("gonoise" in company_low or "gonoise" in tokens):
        return True
    return False
def _employer_conflicts_with_company(blob: str, company: str, tokens: set[str]) -> bool:
    """True when snippet shows a current role at a different organization."""
    company_low = (company or "").lower()
    for m in _EMPLOYER_AT.finditer(blob):
        employer = m.group(1).strip().lower()
        if not employer or len(employer) < 3:
            continue
        if _employer_matches_company(employer, company, tokens):
            continue
        return True
    return False


def snippets_support_contact(
    person: str,
    company: str,
    domain: str,
    snippets: str,
) -> bool:
    """Require search evidence to mention this company/domain, not an unrelated hit."""
    blob = (snippets or "").lower()
    if not blob.strip():
        return False
    tokens = _company_context_tokens(company, domain)
    if not tokens:
        return True
    if not any(t in blob for t in tokens):
        return False

    if _employer_conflicts_with_company(blob, company, tokens):
        return False

    search_tokens = _employment_search_tokens(company, domain, tokens)
    employment_tokens = [t for t in search_tokens if _company_in_employment_context(blob, t)]
    generic = _is_generic_company_name(company)

    if not employment_tokens:
        if re.search(
            r"workshop|activities and societies|studied at|master\b.*\bat\b",
            blob,
            re.I,
        ):
            return False
        if generic:
            return False

    person_low = sanitize_contact_person(person).lower()
    if employment_tokens and person_low and person_low in blob:
        if generic and not _domain_context_in_blob(domain, blob):
            return False
        return True
    if employment_tokens and _LINKEDIN_IN.search(snippets):
        if generic and not _domain_context_in_blob(domain, blob):
            return False
        return True
    return False


def _accept_contact_candidate(
    parsed: dict[str, str],
    *,
    company: str,
    domain: str,
    snippets: str = "",
    require_company_context: bool = False,
    existing_person: str = "",
) -> dict[str, str] | None:
    person = sanitize_contact_person(parsed.get("contact_person", ""))
    if not is_plausible_person_name(person):
        return None
    if existing_person and is_partial_name(existing_person):
        if not partial_upgrade_compatible(existing_person, person):
            return None
    if require_company_context and not snippets_support_contact(person, company, domain, snippets):
        return None
    out = dict(parsed)
    out["contact_person"] = person
    return out


def is_partial_name(name: str) -> bool:
    return bool(_PARTIAL.search(name or ""))


def has_linkedin_name_artifact(name: str) -> bool:
    """True when search-snippet junk was stored as the contact name."""
    s = norm(name)
    if not s:
        return False
    if _GLUED_LINKEDIN.search(s):
        return True
    if _LINKEDIN_PREFIX.match(s):
        return True
    return bool(re.search(r"\blinkedin\b", s, re.I))


def contact_person_needs_upgrade(person: str) -> bool:
    """Stored contact_person should be replaced (partial, junk, or implausible)."""
    if not person:
        return False
    if is_partial_name(person):
        return True
    if has_linkedin_name_artifact(person):
        return True
    if not is_plausible_person_name(person):
        return True
    return False


def needs_name_resolve(row: dict[str, Any]) -> bool:
    crawl = row.get("_crawl") or {}
    # Waterfall already ran for this row — skip on resume (clear name_resolve_attempted to retry).
    if crawl.get("name_resolve_attempted"):
        return False
    person = norm(crawl.get("contact_person"))
    if not person:
        return True
    return contact_person_needs_upgrade(person)


def _fill_empty_only_for_row(row: dict[str, Any]) -> bool:
    """Allow overwrite when the stored name is partial or junk."""
    person = norm((row.get("_crawl") or {}).get("contact_person"))
    return not contact_person_needs_upgrade(person)


def _merge_resolved_contact(
    row: dict[str, Any],
    extracted: dict[str, str],
    *,
    source: str,
) -> None:
    merge_apollo_contact(
        row,
        extracted,
        source=source,
        fill_empty_only=_fill_empty_only_for_row(row),
        fields=frozenset({"contact_person", "contact_role", "linkedin"}),
    )


async def _llm_name_from_snippets(
    company: str,
    domain: str,
    snippets: str,
    *,
    hint_name: str = "",
) -> dict[str, str]:
    if not snippets.strip():
        return {}
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return {}
    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()
    hint = f"\nPartial name hint from prior search: {hint_name}" if hint_name else ""
    prompt = (
        f"Company: {company}\nDomain: {domain}{hint}\n\n"
        "From the search snippets below, extract ONLY if explicitly stated:\n"
        "contact_person (full name — no asterisks/masking), contact_role (job title).\n"
        "Prefer CISO, CEO, Founder, or Head of Security when multiple people appear.\n"
        "The person MUST work at the Company above — ignore unrelated executives from other firms.\n"
        "Never include 'LinkedIn', 'Profile', or site names inside contact_person.\n"
        "Return JSON: contact_person, contact_role. Empty string if unknown. Do not guess.\n\n"
        f"{snippets[:6000]}"
    )
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        text = (resp.choices[0].message.content or "").strip()
        data = json.loads(text) if text else {}
        if not isinstance(data, dict):
            return {}
        out: dict[str, str] = {}
        person = sanitize_contact_person(norm(data.get("contact_person")))
        role = norm(data.get("contact_role"))
        if person and is_plausible_person_name(person):
            out["contact_person"] = person
        if role:
            out["contact_role"] = role
        return out
    except Exception:
        return {}


def _heuristic_name_from_snippets(snippets: str, *, first_hint: str = "") -> dict[str, str]:
    for line in snippets.splitlines():
        m = _NAME_IN_SNIPPET.search(line)
        if not m:
            continue
        name = sanitize_contact_person(m.group(1).strip())
        if first_hint and first_hint.lower() not in name.lower():
            continue
        if is_plausible_person_name(name):
            return {"contact_person": name}
    return {}


async def resolve_name_from_leadership(row: dict[str, Any]) -> dict[str, str] | None:
    from vendor_intel.enrichment.gap_fill.contact_page import extract_leadership_from_website

    extracted = await extract_leadership_from_website(row)
    person = sanitize_contact_person(norm(extracted.get("contact_person")))
    if person and is_plausible_person_name(person):
        cleaned = dict(extracted)
        cleaned["contact_person"] = person
        return cleaned
    return None


async def resolve_name_from_linkedin_mcp(row: dict[str, Any]) -> dict[str, str] | None:
    """LinkedIn MCP: company employees + people search (before DDGS site:linkedin)."""
    try:
        from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_available, resolve_contact_via_mcp
    except ImportError:
        return None
    if not linkedin_mcp_available():
        return None
    company = norm(row.get("company") or row.get("brand"))
    if not company:
        return None
    crawl = row.get("_crawl") or {}
    co_url = norm(crawl.get("linkedin") or row.get("linkedin"))
    return await resolve_contact_via_mcp(company, linkedin_company_url=co_url)


async def resolve_name_from_linkedin_search(
    row: dict[str, Any],
    router: Any,
) -> dict[str, str] | None:
    company = norm(row.get("company") or row.get("brand"))
    domain = domain_from_row(row)
    if not company:
        return None

    crawl = row.get("_crawl") or {}
    existing = norm(crawl.get("contact_person"))
    first = existing.split()[0] if existing else ""

    queries: list[str] = []
    if first and is_partial_name(existing):
        queries.append(f'site:linkedin.com/in "{first}" "{company}"')
    queries.append(
        f'site:linkedin.com/in "{company}" (CISO OR CEO OR Founder OR "Chief Security Officer")'
    )
    if domain:
        queries.append(f'site:linkedin.com/in CISO OR CEO site:{domain}')

    for q in queries:
        snippets = await _search_snippets(router, q, max_hits=8)
        if not snippets:
            continue
        parsed = _heuristic_name_from_snippets(snippets, first_hint=first)
        if not parsed:
            parsed = await _llm_name_from_snippets(
                company, domain, snippets, hint_name=existing
            )
        if parsed.get("contact_person"):
            accepted = _accept_contact_candidate(
                parsed,
                company=company,
                domain=domain,
                snippets=snippets,
                require_company_context=True,
                existing_person=existing,
            )
            if not accepted:
                continue
            li_match = _LINKEDIN_IN.search(snippets)
            if li_match:
                accepted["linkedin"] = f"https://www.linkedin.com/in/{li_match.group(1)}"
            return accepted
    return None


async def resolve_name_from_google_search(
    row: dict[str, Any],
    router: Any,
) -> dict[str, str] | None:
    """Google/web search for full name matching a partial Apollo contact + company."""
    company = norm(row.get("company") or row.get("brand"))
    domain = domain_from_row(row)
    if not company:
        return None

    crawl = row.get("_crawl") or {}
    existing = norm(crawl.get("contact_person"))
    role = norm(crawl.get("contact_role"))
    first = existing.split()[0] if existing else ""
    prefix = _partial_obfuscated_prefix(existing) if is_partial_name(existing) else ""

    queries: list[str] = []
    if first and is_partial_name(existing):
        # Prefer queries that unlock Apollo obfuscation
        if role:
            queries.append(f'"{first}" "{company}" {role}')
        queries.append(f'"{first}" "{company}" (CEO OR Founder OR President OR Director)')
        if prefix:
            queries.append(f'"{first}" {prefix} "{company}" LinkedIn')
    queries.append(f'"{company}" CEO OR Founder OR President OR "Managing Director"')
    if domain:
        queries.append(f'"{company}" site:{domain} (CEO OR Founder OR leadership OR team)')

    for q in queries[:5]:
        try:
            snippets = await _search_snippets(router, q, max_hits=8)
        except Exception:
            continue
        if not snippets:
            continue
        parsed = _heuristic_name_from_snippets(snippets, first_hint=first)
        if not parsed:
            parsed = await _llm_name_from_snippets(
                company, domain, snippets, hint_name=existing
            )
        if not parsed.get("contact_person"):
            continue
        accepted = _accept_contact_candidate(
            parsed,
            company=company,
            domain=domain,
            snippets=snippets,
            require_company_context=True,
            existing_person=existing,
        )
        if accepted:
            return accepted
    return None


async def resolve_name_from_rocketreach(
    row: dict[str, Any],
    router: Any,
) -> dict[str, str] | None:
    """RocketReach: paid API when key set, else public search snippets."""
    company = norm(row.get("company") or row.get("brand"))
    domain = domain_from_row(row)
    if not company:
        return None
    crawl = row.get("_crawl") or {}
    existing = norm(crawl.get("contact_person"))
    first = existing.split()[0] if existing and is_partial_name(existing) else ""

    api_key = (os.getenv("ROCKETREACH_API_KEY") or os.getenv("ROCKET_REACH_API_KEY") or "").strip()
    if api_key:
        try:
            import httpx

            query: dict[str, Any] = {"current_employer": [company]}
            if domain and "." in domain:
                query["company_domain"] = [domain.split("/")[0].removeprefix("www.")]
            if first:
                query["name"] = [first]
            # Prefer leadership titles when we don't have a first-name hint
            if not first:
                query["current_title"] = [
                    "CEO",
                    "Chief Executive Officer",
                    "Founder",
                    "President",
                    "Managing Director",
                ]
            async with httpx.AsyncClient(timeout=30.0) as client:
                resp = await client.post(
                    "https://api.rocketreach.co/api/v2/person/search",
                    headers={"Api-Key": api_key, "Content-Type": "application/json"},
                    json={"query": query, "page_size": 5},
                )
            if resp.status_code in (200, 201):
                data = resp.json() if resp.content else {}
                profiles = data.get("profiles") or data.get("people") or []
                for p in profiles if isinstance(profiles, list) else []:
                    if not isinstance(p, dict):
                        continue
                    name = sanitize_contact_person(
                        str(
                            p.get("name")
                            or p.get("full_name")
                            or " ".join(
                                str(x) for x in (p.get("first_name"), p.get("last_name")) if x
                            )
                        ).strip()
                    )
                    title = str(
                        p.get("current_title") or p.get("title") or p.get("job_title") or ""
                    ).strip()
                    if not name or not is_plausible_person_name(name):
                        continue
                    if existing and is_partial_name(existing):
                        if not partial_upgrade_compatible(existing, name):
                            continue
                    out = {"contact_person": name}
                    if title:
                        out["contact_role"] = title
                    li = str(p.get("linkedin_url") or p.get("linkedin") or "").strip()
                    if li:
                        out["linkedin"] = li
                    return out
            else:
                print(
                    f"    [rocketreach] search HTTP {resp.status_code}: {(resp.text or '')[:160]}",
                    flush=True,
                )
        except Exception as exc:
            print(f"    [rocketreach] API error: {exc}", flush=True)

    # Free fallback: search indexed RocketReach / people-directory pages
    queries = []
    if first:
        queries.append(f'site:rocketreach.co "{first}" "{company}"')
    queries.append(f'site:rocketreach.co "{company}" (CEO OR Founder OR President)')
    for q in queries:
        try:
            snippets = await _search_snippets(router, q, max_hits=6)
        except Exception:
            continue
        if not snippets:
            continue
        parsed = _heuristic_name_from_snippets(snippets, first_hint=first)
        if not parsed:
            parsed = await _llm_name_from_snippets(
                company, domain, snippets, hint_name=existing
            )
        if parsed.get("contact_person"):
            accepted = _accept_contact_candidate(
                parsed,
                company=company,
                domain=domain,
                snippets=snippets,
                require_company_context=True,
                existing_person=existing,
            )
            if accepted:
                return accepted
    return None


async def resolve_contact_name_for_row(
    row: dict[str, Any],
    router: Any,
    *,
    skip_leadership: bool = False,
    skip_linkedin_mcp: bool = False,
    skip_google: bool = False,
    skip_rocketreach: bool = False,
) -> list[str]:
    """Waterfall: LinkedIn MCP → leadership page → Google → RocketReach."""
    if not needs_name_resolve(row):
        return []

    updates: list[str] = []
    company = norm(row.get("company") or row.get("brand"))[:50]

    # 1) LinkedIn company search / leadership via MCP
    if not skip_linkedin_mcp:
        print(f"    [waterfall] {company}: LinkedIn MCP…", flush=True)
        linkedin_mcp_hit = await resolve_name_from_linkedin_mcp(row)
        if linkedin_mcp_hit:
            accepted = _accept_contact_candidate(
                linkedin_mcp_hit,
                company=norm(row.get("company") or row.get("brand")),
                domain=domain_from_row(row),
                require_company_context=False,
            )
            if accepted:
                _merge_resolved_contact(row, accepted, source="linkedin_mcp")
                updates.append("linkedin_mcp")

    if not needs_name_resolve(row):
        return updates

    # 2) Company leadership / about page
    if not skip_leadership:
        print(f"    [waterfall] {company}: leadership page…", flush=True)
        leadership = await resolve_name_from_leadership(row)
        if leadership:
            _merge_resolved_contact(row, leadership, source="leadership_page")
            updates.append("leadership_page")

    if not needs_name_resolve(row):
        return updates

    # 3) Google / web search (incl. LinkedIn profile snippets)
    if not skip_google:
        print(f"    [waterfall] {company}: Google/web search…", flush=True)
        google_hit = await resolve_name_from_google_search(row, router)
        if google_hit:
            _merge_resolved_contact(row, google_hit, source="google_search")
            updates.append("google_search")
        elif needs_name_resolve(row):
            # Keep legacy LinkedIn-site search as part of Google step
            linkedin_hit = await resolve_name_from_linkedin_search(row, router)
            if linkedin_hit:
                _merge_resolved_contact(row, linkedin_hit, source="linkedin_search")
                updates.append("linkedin_search")

    if not needs_name_resolve(row):
        return updates

    # 4) RocketReach
    if not skip_rocketreach:
        print(f"    [waterfall] {company}: RocketReach…", flush=True)
        rr_hit = await resolve_name_from_rocketreach(row, router)
        if rr_hit:
            _merge_resolved_contact(row, rr_hit, source="rocketreach")
            updates.append("rocketreach")

    return updates


async def resolve_contact_names_for_market(
    rows: list[dict[str, Any]],
    *,
    concurrent: int = 3,
    skip_leadership: bool = False,
    skip_linkedin_mcp: bool = False,
    skip_google: bool = False,
    skip_rocketreach: bool = False,
    persist_fn: Any = None,
) -> dict[str, Any]:
    import asyncio

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings
    from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_available

    router = FreeSearchRouter(Settings.load())
    todo = [r for r in rows if needs_name_resolve(r)]
    skipped_done = sum(1 for r in rows if r.get("is_relevant", True) and not needs_name_resolve(r))
    if not todo:
        return {"queued": 0, "resolved": 0, "still_partial": 0, "skipped_already_done": skipped_done}

    # MCP is serial; keep one worker when using LinkedIn MCP
    if not skip_linkedin_mcp and linkedin_mcp_available():
        concurrent = min(concurrent, 1)

    sem = asyncio.Semaphore(max(1, concurrent))
    resolved = 0
    still_partial = 0

    print(
        f"  [name_resolve] queue={len(todo)} (skip already done={skipped_done}) waterfall="
        f"LinkedIn→leadership→Google→RocketReach workers={max(1, concurrent)}",
        flush=True,
    )

    def _checkpoint() -> None:
        if persist_fn is None:
            return
        try:
            persist_fn()
        except Exception as exc:
            print(f"  [name_resolve] checkpoint save failed: {exc}", flush=True)

    async def _one(row: dict[str, Any], idx: int) -> None:
        nonlocal resolved, still_partial
        async with sem:
            name = str(row.get("company") or "?")[:50]
            before = norm((row.get("_crawl") or {}).get("contact_person"))
            print(f"  [name_resolve] starting [{idx}/{len(todo)}] {name} ({before!r})", flush=True)
            try:
                applied = await resolve_contact_name_for_row(
                    row,
                    router,
                    skip_leadership=skip_leadership,
                    skip_linkedin_mcp=skip_linkedin_mcp,
                    skip_google=skip_google,
                    skip_rocketreach=skip_rocketreach,
                )
                if applied:
                    resolved += 1
                    crawl = dict(row.get("_crawl") or {})
                    crawl["name_resolve_attempted"] = True
                    row["_crawl"] = crawl
                    person = norm((row.get("_crawl") or {}).get("contact_person"))
                    print(
                        f"  [name_resolve] [{idx}/{len(todo)}] {name} +{applied} -> {person!r}",
                        flush=True,
                    )
                elif needs_name_resolve(row):
                    still_partial += 1
                    crawl = dict(row.get("_crawl") or {})
                    crawl["name_resolve_attempted"] = True
                    row["_crawl"] = crawl
                    print(
                        f"  [name_resolve] [{idx}/{len(todo)}] {name} still partial",
                        flush=True,
                    )
            except Exception as exc:
                print(f"  [name_resolve] [{idx}/{len(todo)}] {name} ERROR: {exc}", flush=True)
                if needs_name_resolve(row):
                    still_partial += 1
            _checkpoint()

    await asyncio.gather(*(_one(r, i) for i, r in enumerate(todo, 1)))

    return {
        "queued": len(todo),
        "resolved": resolved,
        "still_partial": still_partial,
        "skipped_already_done": skipped_done,
    }
