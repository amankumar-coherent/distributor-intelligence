"""Apollo.io contact lookup — search + optional enrich for contact person name."""
from __future__ import annotations

import asyncio
import os
import re
from typing import Any
from urllib.parse import urlparse

import httpx

APOLLO_BASE = "https://api.apollo.io/api/v1"

DEFAULT_REQUEST_DELAY_SEC = float(os.getenv("APOLLO_REQUEST_DELAY_SEC", "3.5"))
DEFAULT_MAX_RETRIES = int(os.getenv("APOLLO_MAX_RETRIES", "4"))
MAX_RETRY_WAIT_SEC = float(os.getenv("APOLLO_MAX_RETRY_WAIT_SEC", "900"))

# Prefer sales / channel / procurement leadership for Contact Person + Role (not CEO/CISO).
# Higher score = preferred. Scores below SALES_TITLE_MIN_SCORE are ignored.
TITLE_KEYWORDS: list[tuple[str, int]] = [
    (r"head of sales|sales head", 100),
    (r"\bvp\b.*sales|vice president.*sales|svp.*sales", 95),
    (r"chief (?:revenue|commercial) officer|\bcro\b", 92),
    (r"sales director|director.*sales|director of sales", 90),
    (r"head of (?:business development|bd|channel|partner)", 88),
    (r"head of procurement|procurement head|chief procurement|\bcpo\b", 88),
    (r"vp.*(?:business development|bd|channel|partner)|vice president.*(?:channel|partner)", 86),
    (r"\bvp\b.*procurement|vice president.*procurement|svp.*procurement", 86),
    (r"commercial director|commercial head", 84),
    (r"director.*(?:business development|bd|channel|partner)", 82),
    (r"director.*procurement|procurement director", 82),
    (r"national sales|regional sales|sales manager", 78),
    (r"procurement manager|purchasing manager|head of purchasing|purchasing director", 78),
    (r"business development|channel (?:sales|manager|partner)|partner manager", 74),
    (r"account (?:executive|director|manager).*sales|enterprise sales", 70),
]

# Minimum score to accept a person (blocks CEO/Founder fallbacks).
SALES_TITLE_MIN_SCORE = 70

SEARCH_TITLES = [
    "Head of Sales",
    "VP Sales",
    "Vice President of Sales",
    "Sales Director",
    "Director of Sales",
    "Chief Revenue Officer",
    "CRO",
    "Commercial Director",
    "Head of Business Development",
    "VP Business Development",
    "Director of Business Development",
    "Head of Channel",
    "Channel Sales",
    "Partner Manager",
    "National Sales Manager",
    "Regional Sales Manager",
    "Head of Procurement",
    "VP Procurement",
    "Vice President of Procurement",
    "Procurement Manager",
    "Procurement Director",
    "Director of Procurement",
    "Chief Procurement Officer",
    "Purchasing Manager",
    "Head of Purchasing",
]

# Avoid c_suite / founder so Apollo does not prefer CEO over sales leads.
SEARCH_SENIORITIES = ["vp", "head", "director", "manager"]

# Wrong domains in some exports — map to corporate site for Apollo search.
DOMAIN_OVERRIDES: dict[str, str] = {
    "office.com": "microsoft.com",
    "store.kaspersky-labs.com": "kaspersky.com",
    "kaspersky-labs.com": "kaspersky.com",
}


def _api_key() -> str:
    key = (os.getenv("APOLLO_API_KEY") or "").strip()
    if not key:
        raise RuntimeError("APOLLO_API_KEY is not set in agents/vendor-intel/.env")
    return key


def company_domain(row: dict[str, Any]) -> str:
    for raw in (row.get("domain"), row.get("website")):
        s = str(raw or "").strip().lower()
        if not s:
            continue
        if "@" in s:
            s = s.split("@", 1)[1]
        if "://" in s:
            s = urlparse(s if "://" in s else f"https://{s}").netloc or s
        s = s.removeprefix("www.").split("/")[0].strip()
        if s and "." in s:
            return DOMAIN_OVERRIDES.get(s, s)
    crawl = row.get("_crawl") if isinstance(row.get("_crawl"), dict) else {}
    for raw in (crawl.get("website"),):
        s = str(raw or "").strip().lower()
        if s and "." in s:
            if "://" in s:
                s = urlparse(s).netloc or s
            return s.removeprefix("www.").split("/")[0]
    return ""


DEPT_OK = re.compile(
    r"sales|procurement|purchasing|"
    r"\bcro\b|chief revenue|revenue officer|"
    r"commercial|"
    r"business development|\bbd\b|"
    r"channel sales|partner sales",
    re.I,
)
DEPT_BLOCK = re.compile(
    r"\b(ceo|cfo|cto|coo|ciso|founder|co-founder|chairman|owner|hr|"
    r"human resources|marketing|operations|project director)\b",
    re.I,
)


def _score_title(title: str) -> int:
    t = (title or "").lower()
    # Reject noisy multi-role dumps from Apollo
    if t.count("|") >= 2 or len(t) > 140:
        return 0
    if not DEPT_OK.search(t):
        return 0
    if DEPT_BLOCK.search(t) and not re.search(
        r"sales|procurement|purchasing|commercial|revenue", t, re.I
    ):
        return 0
    best = 0
    for pattern, pts in TITLE_KEYWORDS:
        if re.search(pattern, t, re.I):
            best = max(best, pts)
    return best


def _person_display_name(person: dict[str, Any]) -> str:
    name = str(person.get("name") or "").strip()
    if name:
        return name
    first = str(person.get("first_name") or "").strip()
    last = str(person.get("last_name") or "").strip()
    if first and last and "*" not in last:
        return f"{first} {last}".strip()
    ob = str(person.get("last_name_obfuscated") or "").strip()
    if first and ob:
        return f"{first} {ob}".strip()
    return first or ""


def _name_usable(name: str, *, allow_obfuscated: bool) -> bool:
    """First name is enough; obfuscated last name (Mo***s) allowed when flagged."""
    s = (name or "").strip()
    if not s:
        return False
    first = s.split()[0]
    letters = re.sub(r"[^A-Za-z]", "", first)
    if len(letters) < 2:
        return False
    if "*" in s:
        return allow_obfuscated
    return True


def _first_token(name: str) -> str:
    return (name or "").strip().lower().split()[0] if name else ""


def contact_names_match(existing: str, candidate: str) -> bool:
    """True when Apollo candidate is likely the same person already on the row."""
    a = (existing or "").strip().lower()
    b = (candidate or "").strip().lower()
    if not a or not b:
        return False
    if a == b:
        return True
    if a in b or b in a:
        return True
    af, bf = _first_token(a), _first_token(b)
    if af and bf and af == bf and len(af) >= 3:
        return True
    if af and bf and af == bf and (len(a.split()) == 1 or len(b.split()) == 1):
        return True
    return False


class ApolloContactClient:
    def __init__(
        self,
        *,
        api_key: str | None = None,
        timeout: float = 45.0,
        request_delay: float | None = None,
        max_retries: int | None = None,
        verbose: bool = True,
    ) -> None:
        self.api_key = (api_key or _api_key()).strip()
        self.timeout = timeout
        self.request_delay = DEFAULT_REQUEST_DELAY_SEC if request_delay is None else request_delay
        self.max_retries = DEFAULT_MAX_RETRIES if max_retries is None else max(1, max_retries)
        self.verbose = verbose
        self._client: httpx.AsyncClient | None = None
        self._request_lock = asyncio.Lock()
        self._last_request_at = 0.0
        self.credits_exhausted = False

    def _log(self, message: str) -> None:
        if self.verbose:
            print(f"    {message}", flush=True)

    async def _throttled_post(self, url: str, **kwargs: Any) -> httpx.Response:
        assert self._client is not None
        resp: httpx.Response | None = None
        endpoint = url.rsplit("/", 1)[-1]
        for attempt in range(self.max_retries):
            async with self._request_lock:
                loop = asyncio.get_running_loop()
                wait = self.request_delay - (loop.time() - self._last_request_at)
                if wait > 0:
                    self._log(f"spacing {wait:.1f}s before {endpoint}...")
                    await asyncio.sleep(wait)
                self._log(f"POST {endpoint} (attempt {attempt + 1}/{self.max_retries})...")
                resp = await self._client.post(url, **kwargs)
                self._last_request_at = loop.time()

            if resp.status_code != 429:
                if resp.status_code >= 400 and self.verbose:
                    self._log(f"{endpoint} -> HTTP {resp.status_code}")
                return resp

            retry_after = resp.headers.get("Retry-After", "").strip()
            if retry_after.isdigit():
                raw_backoff = float(retry_after)
                # Honor Apollo's cool-down up to MAX_RETRY_WAIT_SEC (default 15m)
                backoff = min(MAX_RETRY_WAIT_SEC, max(30.0, raw_backoff))
                if raw_backoff > MAX_RETRY_WAIT_SEC:
                    self._log(
                        f"rate limited (429) — Apollo Retry-After={raw_backoff:.0f}s, "
                        f"waiting {backoff:.0f}s (cap)..."
                    )
                else:
                    self._log(f"rate limited (429) — waiting {backoff:.0f}s then retry...")
            else:
                backoff = min(MAX_RETRY_WAIT_SEC, (2**attempt) * 8.0 + 5.0)
                self._log(f"rate limited (429) — waiting {backoff:.0f}s then retry...")
            if attempt < self.max_retries - 1:
                await asyncio.sleep(backoff)
                # After a long cool-down, one clean retry is enough
                if backoff >= 120:
                    continue
            else:
                break

        assert resp is not None
        return resp

    async def __aenter__(self) -> ApolloContactClient:
        self._client = httpx.AsyncClient(
            timeout=self.timeout,
            headers={
                "x-api-key": self.api_key,
                "Content-Type": "application/json",
                "Cache-Control": "no-cache",
            },
        )
        return self

    async def __aexit__(self, *args: object) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None

    async def search_people(self, domain: str, *, company_name: str = "") -> list[dict[str, Any]]:
        assert self._client is not None
        params: list[tuple[str, str | int]] = [
            ("page", 1),
            ("per_page", 10),
        ]
        for d in [domain]:
            params.append(("q_organization_domains_list[]", d))
        for t in SEARCH_TITLES:
            params.append(("person_titles[]", t))
        for s in SEARCH_SENIORITIES:
            params.append(("person_seniorities[]", s))

        resp = await self._throttled_post(f"{APOLLO_BASE}/mixed_people/api_search", params=params)
        if resp.status_code == 429:
            raise RuntimeError(
                f"Apollo rate limit (429) on people search after {self.max_retries} retries — "
                f"wait a few minutes or increase APOLLO_REQUEST_DELAY_SEC (now {self.request_delay}s)"
            )
        resp.raise_for_status()
        people = list(resp.json().get("people") or [])
        self._log(f"domain search ({domain or company_name}): {len(people)} people")
        if not people and company_name:
            self._log(f"no results for domain — trying company name: {company_name}")
            params2: list[tuple[str, str | int]] = [
                ("page", 1),
                ("per_page", 10),
                ("q_organization_name", company_name),
            ]
            for t in SEARCH_TITLES:
                params2.append(("person_titles[]", t))
            for s in SEARCH_SENIORITIES:
                params2.append(("person_seniorities[]", s))
            resp2 = await self._throttled_post(f"{APOLLO_BASE}/mixed_people/api_search", params=params2)
            if resp2.status_code == 429:
                raise RuntimeError(
                    f"Apollo rate limit (429) on people search after {self.max_retries} retries"
                )
            if resp2.status_code == 200:
                people = list(resp2.json().get("people") or [])
                self._log(f"name search: {len(people)} people")
        return people

    async def enrich_person(self, person_id: str, *, domain: str) -> dict[str, Any] | None:
        assert self._client is not None
        if self.credits_exhausted:
            return None
        resp = await self._throttled_post(
            f"{APOLLO_BASE}/people/match",
            params={"id": person_id, "domain": domain},
        )
        if resp.status_code == 429:
            return None
        if resp.status_code == 422 and "insufficient credits" in resp.text.lower():
            self.credits_exhausted = True
            return None
        if resp.status_code != 200:
            return None
        data = resp.json()
        person = data.get("person")
        return person if isinstance(person, dict) else None

    def _build_contact(self, person: dict[str, Any], *, fallback_title: str = "") -> dict[str, str]:
        name = _person_display_name(person)
        title = str(person.get("title") or fallback_title or "").strip()
        linkedin = str(person.get("linkedin_url") or "").strip()
        email = str(person.get("email") or "").strip()
        phone = ""
        phones = person.get("phone_numbers")
        if isinstance(phones, list) and phones:
            p0 = phones[0]
            if isinstance(p0, dict):
                phone = str(p0.get("sanitized_number") or p0.get("raw_number") or "").strip()
            else:
                phone = str(p0).strip()
        if not phone:
            phone = str(person.get("phone") or person.get("sanitized_phone") or "").strip()

        out: dict[str, str] = {
            "contact_person": name,
            "contact_role": title,
            "apollo_person_id": str(person.get("id") or ""),
            "apollo_source_url": "https://app.apollo.io/",
        }
        if linkedin:
            out["linkedin"] = linkedin
        if email and "@" in email:
            out["email"] = email
        if phone and len(re.sub(r"\D", "", phone)) >= 8:
            out["phone"] = phone
        if "*" in name:
            out["apollo_name_obfuscated"] = "true"
        if self.credits_exhausted:
            out["apollo_credits_exhausted"] = "true"
        return out

    async def find_contact(
        self,
        row: dict[str, Any],
        *,
        enrich_match: bool = True,
    ) -> dict[str, str] | None:
        domain = company_domain(row)
        company = str(row.get("company") or row.get("brand") or "").strip()
        if not domain and not company:
            return None

        self._log(f"lookup domain={domain or '?'} company={company}")

        try:
            people = await self.search_people(domain, company_name=company)
        except Exception as exc:
            return {"error": str(exc)}

        if not people:
            return None

        ranked: list[tuple[int, dict[str, Any]]] = []
        for p in people:
            if not isinstance(p, dict):
                continue
            title = str(p.get("title") or "")
            ranked.append((_score_title(title), p))
        ranked.sort(key=lambda x: x[0], reverse=True)
        # Sales / channel titles only — do not fall back to CEO/Founder/random execs.
        candidates = [
            (score, p) for score, p in ranked if score >= SALES_TITLE_MIN_SCORE
        ]
        if not candidates:
            self._log(
                f"no Head-of-Sales / sales-leadership title "
                f"(min score {SALES_TITLE_MIN_SCORE}) — skip contact"
            )
            return None

        # 1) Prefer any usable search name (full OR first + partial last). 0 credits.
        usable = [
            (score, p)
            for score, p in candidates
            if _name_usable(_person_display_name(p), allow_obfuscated=True)
        ]
        if usable:
            usable.sort(key=lambda x: x[0], reverse=True)
            return self._build_contact(
                usable[0][1],
                fallback_title=str(usable[0][1].get("title") or ""),
            )

        # 2) Try enrich on top candidates while credits remain.
        if enrich_match and domain and not self.credits_exhausted:
            for score, candidate in candidates[:5]:
                person_id = str(candidate.get("id") or "").strip()
                if not person_id:
                    continue
                enriched = await self.enrich_person(person_id, domain=domain)
                if not enriched:
                    continue
                name = _person_display_name(enriched)
                if _name_usable(name, allow_obfuscated=False):
                    return self._build_contact(
                        enriched,
                        fallback_title=str(candidate.get("title") or ""),
                    )

        # 3) Credits exhausted or enrich unavailable — keep obfuscated search names.
        allow_obfuscated = not enrich_match or self.credits_exhausted
        for score, candidate in candidates:
            name = _person_display_name(candidate)
            if _name_usable(name, allow_obfuscated=allow_obfuscated):
                return self._build_contact(
                    candidate,
                    fallback_title=str(candidate.get("title") or ""),
                )

        return None

    async def find_role_for_existing_contact(self, row: dict[str, Any]) -> dict[str, str] | None:
        """Return role/title for the contact_person already on the row (same Apollo person)."""
        crawl = row.get("_crawl") if isinstance(row.get("_crawl"), dict) else {}
        existing = str(crawl.get("contact_person") or "").strip()
        if not existing or len(existing) < 3:
            return None

        domain = company_domain(row)
        company = str(row.get("company") or row.get("brand") or "").strip()
        if not domain and not company:
            return None

        self._log(f"role lookup for existing contact={existing!r} domain={domain or '?'}")

        try:
            people = await self.search_people(domain, company_name=company)
        except Exception as exc:
            return {"error": str(exc)}

        matched: list[tuple[int, dict[str, Any]]] = []
        for p in people:
            if not isinstance(p, dict):
                continue
            cand = _person_display_name(p)
            if not contact_names_match(existing, cand):
                continue
            title = str(p.get("title") or "")
            matched.append((_score_title(title), p))

        if not matched:
            return None

        matched.sort(key=lambda x: x[0], reverse=True)
        _, person = matched[0]
        title = str(person.get("title") or "").strip()
        if not title:
            return None

        return {
            "contact_person": existing,
            "contact_role": title,
            "apollo_person_id": str(person.get("id") or crawl.get("apollo_person_id") or ""),
            "apollo_source_url": "https://app.apollo.io/",
        }
