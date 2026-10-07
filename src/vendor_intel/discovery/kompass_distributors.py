"""Global Kompass distributor discovery pipeline.

INPUT MARKET -> AI finds Kompass product/classification page -> extract
product id -> find global distributor candidates -> open each company
profile -> verify distributor status is MARKET-SPECIFIC (not just "the
company is a distributor of something") -> extract company fields ->
extract Kompass company id -> dedupe -> score -> filter by
MIN_DISTRIBUTOR_CONFIDENCE -> return SLIM_COLUMNS-shaped rows.

Hard rules enforced throughout (do not weaken these without re-confirming
with the user first) — a company is only accepted when ALL of these hold:
  1. Kompass explicitly types it as (or including) "Distributor".
  2. There is concrete evidence tying that distributor role to the
     REQUESTED market/product — not merely that the market keyword appears
     somewhere on the page (description, unrelated product line, generic
     software/service offering, etc).
  3. The company carries NO manufacturer/OEM/producer role at all — not on
     the matched market classification, and not anywhere else on the
     profile. Even a "Manufacturer Distributor" combo, which the pipeline
     originally accepted, is rejected outright per explicit user
     instruction: no manufacturer/OEM involvement, pure distributors only.
AI is used only for page discovery and assisting verification when
deterministic rules are inconclusive — it never invents company data
(founded year, employees, phone, email, website, Kompass IDs). Every
important field is sourced from an actually-fetched page.
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Protocol
from urllib.parse import urljoin, urlparse

from vendor_intel.clients import kompass as kompass_client
from vendor_intel.discovery.chatgpt_company_screen import screening_enabled
from vendor_intel.discovery.country_sweep import world_countries
from vendor_intel.discovery.kompass_ai_search import discover_product_page
from vendor_intel.export.landscape_slim import SLIM_COLUMNS


class _CheckpointLike(Protocol):
    """Structural type for vendor_intel.pipeline.openai_expand_checkpoint
    .ExpandCheckpoint — duck-typed here so this module doesn't need to
    import from pipeline/ (kept as a plain dependency-free discovery module).
    Any object with this shape works: bump()/data()/save() are all
    ExpandCheckpoint already has."""

    def bump(self, step: str, substep: str, *, note: str = "", **data: Any) -> None: ...
    def data(self, key: str, default: Any = None) -> Any: ...
    def save(self, note: str = "") -> None: ...

NOT_DISCLOSED = "Not publicly disclosed"

# --- Kompass company-type strings that count as "distributor" (spec §5) ---
_DISTRIBUTOR_TYPE_RE = re.compile(r"\bdistributor\b", re.I)
# Manufacturer/OEM/producer roles — a company carrying ANY of these is
# rejected outright, even if it also carries a Distributor role for the same
# or a different classification. Pure distributors only, no manufacturer/OEM
# involvement (user explicitly excluded combo roles like "Manufacturer
# Distributor" — stricter than plain "must have a Distributor tag").
_MANUFACTURER_TYPE_RE = re.compile(r"\b(manufacturer|oem|producer)\b", re.I)

# Company-type badge text on Kompass profiles, e.g. "Manufacturer Distributor
# Service" — Kompass concatenates several role words together. Fallback only
# — the per-activity classification tree (_ACTIVITY_ITEM_RE) is authoritative
# when present.
_COMPANY_TYPE_BLOCK_RE = re.compile(
    r"(Manufacturer|Distributor|Importer|Exporter|Service|Retailer|Wholesaler|"
    r"Agent|Trading Company|Trader)(?:\s*[/,•|]?\s*"
    r"(Manufacturer|Distributor|Importer|Exporter|Service|Retailer|Wholesaler|"
    r"Agent|Trading Company|Trader))*",
    re.I,
)

# Kompass's real per-classification structure: each activity a company is
# tagged with carries its OWN role marker(s) ("circle distributor"/"circle
# supplier"/"circle service") and its own classification id/name — e.g.
#   <a id="8125039" class="activityLink" href=".../telecommunications-
#     equipment-wholesale/8125039/"><i class="circle distributor"></i>
#     <span class="text">Telecommunications equipment (wholesale)</span></a>
# This is the authoritative, market-SPECIFIC signal the spec requires: a
# company can be a "distributor" for one activity and merely a "supplier"
# or "service provider" for another. Confirmed against real Kompass HTML for
# https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/.
_ACTIVITY_ITEM_RE = re.compile(
    r'<a id="(\d+)" class="activityLink" href="([^"]+)">(.*?)</a>', re.S
)
_ACTIVITY_CIRCLE_RE = re.compile(r'circle (\w+)')
_ACTIVITY_TEXT_RE = re.compile(r'<span class="text">([^<]+)</span>')

_FOUNDED_RE = re.compile(
    r"(?:Year established|Founded|Established)\s*[:\-]?\s*(\d{4})", re.I
)
_EMPLOYEES_EXACT_RE = re.compile(r"\b(\d{1,7})\s*Employees?\b", re.I)
_EMPLOYEES_RANGE_RE = re.compile(r"\b(\d{1,6}\s*-\s*\d{1,6})\s*Employees?\b", re.I)
_EMPLOYEES_NO_EMPLOYEE_RE = re.compile(r"\bNo employees\b", re.I)

_WEBSITE_RE = re.compile(r'https?://[^\s"\'<>)]+', re.I)
_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_PHONE_RE = re.compile(r"(?:\+\d{1,3}[\s.\-]?)?(?:\(\d{1,4}\)[\s.\-]?)?\d{2,4}[\s.\-]\d{2,4}[\s.\-]?\d{0,4}")

_COMPANY_LINK_RE = re.compile(
    r'href="((?:https?://[^"]*kompass\.com)?/(?:z/[a-z]{2}/)?c/[^"/]+/[a-z]{2,4}\d{4,}/?)"', re.I
)
_NEXT_PAGE_RE = re.compile(r'href="([^"]+)"[^>]*>\s*(?:Next|»|&raquo;)', re.I)
# Kompass paginates by PATH (".../23200/page-2/"), rendered as numbered
# "pageBtn" links — it has no "Next"/"»" link at all, so _NEXT_PAGE_RE alone
# never advances past page 1 (observed live: 77 companies found instead of
# thousands). Match the numbered page links and pick the next one by number.
_PAGE_LINK_RE = re.compile(r'href="([^"]*?/page-(\d+)/?)"', re.I)
_CURRENT_PAGE_RE = re.compile(r"/page-(\d+)/?$")
_JSONLD_BLOCK_RE = re.compile(
    r'<script type="application/ld\+json">\s*(\{.*?\})\s*</script>', re.S
)


def _log(msg: str) -> None:
    print(f"  [kompass] {msg}", flush=True)


def _sub(msg: str) -> None:
    print(f"    -> {msg}", flush=True)


def _detail(msg: str) -> None:
    print(f"      . {msg}", flush=True)


def max_companies_per_market() -> int:
    """0 (or unset/blank) means no upper limit — Kompass discovery keeps
    paginating until it genuinely runs out of pages/candidates."""
    raw = (os.getenv("MAX_COMPANIES_PER_MARKET") or "0").strip()
    try:
        n = int(raw)
    except ValueError:
        return 0
    return max(0, n)


def target_distributors_per_market() -> int:
    """Stop screening once this many VERIFIED DISTRIBUTORS are accepted.

    Distinct from MAX_COMPANIES_PER_MARKET, which caps how many candidates
    are DISCOVERED. Only ~21% of candidates survive screening, so capping
    discovery at 500 yields ~105 distributors, not 500. This caps the thing
    the user actually wants a fixed number of.

    0 (or unset) means no cap — screen every candidate.
    """
    raw = (os.getenv("TARGET_DISTRIBUTORS_PER_MARKET") or "0").strip()
    try:
        return max(0, int(raw))
    except ValueError:
        return 0


def min_companies_per_market() -> int:
    """Floor for TOTAL verified distributors for a market. If Kompass alone
    doesn't reach this, the ChatGPT/DeepSeek expand pipeline is used to top
    up (see run_cursor_expand.py --backend kompass)."""
    try:
        return max(0, int(os.getenv("MIN_COMPANIES_PER_MARKET", "500")))
    except ValueError:
        return 500


def min_distributor_confidence() -> int:
    try:
        return max(0, min(100, int(os.getenv("MIN_DISTRIBUTOR_CONFIDENCE", "65"))))
    except ValueError:
        return 65


def global_search_enabled() -> bool:
    return (os.getenv("GLOBAL_SEARCH", "true") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


@dataclass
class KompassCompanyCandidate:
    """A company found on a Kompass distributor-results listing, before its
    profile page has been opened/verified."""

    name: str
    profile_url: str
    country_hint: str = ""
    # Facts the listing's JSON-LD already gives us. Carried through so the
    # enrichment stage doesn't re-fetch the company's Kompass profile just
    # to learn them — those per-profile fetches are what Kompass
    # rate-limits hardest (HTTP 405 / login redirects).
    city_hint: str = ""
    phone_hint: str = ""
    employees_hint: str = ""


@dataclass
class VerifiedDistributor:
    company_name: str
    market: str
    kompass_company_id: str
    kompass_profile_url: str
    kompass_product_id: str
    kompass_company_type: str
    is_actual_market_distributor: bool
    verification_confidence: int
    rejection_reason: str = ""

    website: str = ""
    headquarters: str = ""
    country: str = ""
    city: str = ""
    founded_year: str = ""
    employees: str = ""
    description: str = ""
    brands: str = ""
    products: str = ""
    contact_person: str = ""
    email: str = ""
    phone: str = ""
    address: str = ""
    verification_urls: list[str] = field(default_factory=list)
    # Any additional SLIM_COLUMNS values researched for this company that
    # don't map onto a dedicated field above (e.g. Role, LinkedIn,
    # Specialty Focus). Applied verbatim by to_slim_row(), so a new export
    # column can be filled without widening this dataclass each time.
    extra_columns: dict[str, str] = field(default_factory=dict)


@dataclass
class MarketDistributorResult:
    market: str
    market_status: str  # "ok" | "kompass_market_not_found" | "kompass_access_error"
    kompass_product_id: str = ""
    kompass_product_url: str = ""
    distributors: list[VerifiedDistributor] = field(default_factory=list)
    rejected_count: int = 0


# --------------------------------------------------------------------------
# Step 1-2: market -> Kompass product page -> product id
# --------------------------------------------------------------------------


def find_market_product_page(market: str) -> tuple[str, str] | None:
    """Returns (kompass_product_url, kompass_product_id) or None. The id is
    always re-extracted from the verified fetched URL — never taken on the
    AI's word (see kompass_ai_search.discover_product_page)."""
    # An operator-supplied id short-circuits AI discovery. Automatic
    # discovery legitimately fails for markets named the way analyst reports
    # name them ("Semiconductor Market", "Data Center Systems Market") —
    # Kompass indexes by product classification, so there is no page with
    # that title even though a perfectly good category exists
    # ("Semiconductors NES", "Data centres"). Rather than return nothing,
    # let the id be passed in via KOMPASS_PRODUCT_ID.
    #
    # The id is still VERIFIED by fetching it: a wrong id is rejected here
    # exactly as an AI-invented one would be, so this is a shortcut past
    # discovery, never past verification.
    forced = (os.getenv("KOMPASS_PRODUCT_ID") or "").strip()
    if forced:
        url = f"https://www.kompass.com/z/ww/a/x/{forced}/"
        page = kompass_client.fetch_kompass_page(url)
        product_id = kompass_client.extract_product_id(page.final_url or "") or (
            kompass_client.extract_product_id(url)
        )
        if page.alive and product_id:
            _log(f"using KOMPASS_PRODUCT_ID={product_id} -> {page.final_url}")
            return page.final_url or url, product_id
        _log(
            f"KOMPASS_PRODUCT_ID={forced} did not verify "
            f"(alive={page.alive}) — falling back to AI discovery"
        )

    candidate = discover_product_page(market)
    if candidate is None:
        return None
    return candidate.kompass_product_url, candidate.kompass_product_id


# --------------------------------------------------------------------------
# Step 3-4: global distributor candidate discovery from the product page
# --------------------------------------------------------------------------


_ACTIVITY_PATH_RE = re.compile(r"/a/([^/]+)/(\d+)/?$")


def kompass_country_domains() -> list[str]:
    """Country subdomains to sweep for distributor-filtered listings.

    Override with KOMPASS_COUNTRY_DOMAINS (comma-separated, e.g. "de,fr,in").
    The default list mirrors Kompass's own country sites; each returns its
    OWN companies with no cross-domain overlap (verified live: lb->LB,
    de->DEEX, fr->FR, in->IN, us->USC, ...), so sweeping them is what
    produces genuinely global distributor coverage."""
    raw = (os.getenv("KOMPASS_COUNTRY_DOMAINS") or "").strip()
    if raw:
        return [d.strip().lower() for d in raw.split(",") if d.strip()]
    return [
        "us", "gb", "de", "fr", "it", "es", "nl", "be", "ch", "at", "se", "no", "dk", "fi",
        "pl", "cz", "hu", "ro", "gr", "pt", "ie", "tr", "ru", "ua", "in", "cn", "jp", "kr",
        "sg", "my", "th", "id", "ph", "vn", "au", "nz", "ca", "mx", "br", "ar", "cl", "co",
        "za", "eg", "ma", "tn", "dz", "ng", "ke", "ae", "sa", "qa", "kw", "lb", "il", "pk",
        "bd", "lk", "tw", "hk",
    ]


def _build_distributor_listing_urls(product_url: str) -> list[str]:
    """Build candidate listing URLs from a verified product page, most
    valuable first.

    1. The global listing (/z/ww/a/...) — one URL covering every country at
       once. This is the primary source.
    2. Kompass's per-country distributor listings
       (/x/distributor/a/<slug>/<id>/ on each country domain). Kompass DOES
       expose a real server-side distributor filter, but it does NOT combine
       with global scope: supplying /z/ww/ and /x/distributor/ together
       silently drops the filter and redirects back to the unfiltered global
       listing. Each country domain also returns a disjoint company set
       (verified: de->DEEX, fr->FR, in->IN, us->USC, ...), so sweeping them
       after the global pass picks up companies the global listing's own
       pagination didn't reach.

    Note the distributor filter is a coverage aid, not a guarantee —
    companies from those listings still verify as Manufacturer for a given
    activity, so verify_company() remains the authoritative check either way.

    Every URL here is independently verified with a real fetch before use
    (spec §4/§25) — none of these shapes is assumed to exist."""
    candidates: list[str] = [product_url]

    match = _ACTIVITY_PATH_RE.search(urlparse(product_url).path)
    if match:
        slug, activity_id = match.group(1), match.group(2)
        for domain in kompass_country_domains():
            candidates.append(
                f"https://{domain}.kompass.com/x/distributor/a/{slug}/{activity_id}/"
            )
    return candidates


def _listing_jsonld_companies(html: str) -> dict[str, dict[str, Any]]:
    """Map profile-URL -> company facts taken from the listing's JSON-LD
    ItemList. Kompass embeds name, url, city/country, phone, employee range
    and a description for every company right in the listing page, so these
    fields are available WITHOUT opening each company's profile (which is
    rate-limited and frequently 405s/redirects to login).

    Returns {} if the block is missing or unparseable — callers fall back to
    link-only extraction."""
    out: dict[str, dict[str, Any]] = {}
    for block in _JSONLD_BLOCK_RE.findall(html):
        try:
            data = json.loads(block)
        except (json.JSONDecodeError, ValueError):
            continue
        items = (data.get("mainEntity") or {}).get("itemListElement") or []
        for entry in items:
            org = (entry or {}).get("item") or {}
            url = str(org.get("url") or "").strip()
            if not url:
                continue
            out[url.rstrip("/").lower()] = org
    return out


def _extract_company_links(html: str, base_url: str) -> list[KompassCompanyCandidate]:
    """Collect company candidates from a listing page.

    Names (and other facts) come from the page's JSON-LD ItemList where
    available, so downstream steps don't have to open each profile just to
    learn the company's name.

    NOTE: there is deliberately no per-card "is this a distributor?"
    pre-filter here. Kompass's listing HTML carries the companyType role
    tags only once, as a hidden template (wrapped in French dev comments,
    "si la société est un producteur") — real cards get their tags
    injected client-side by JS, so the server HTML we fetch has no
    per-company role data to filter on."""
    facts = _listing_jsonld_companies(html)
    # Index the JSON-LD by COMPANY ID as well as by full URL. The same
    # company appears in two URL shapes on the same page — JSON-LD uses
    # "/c/<slug>/<id>" while the HTML links use "/z/ww/c/<slug>/<id>" — so a
    # URL-keyed lookup silently missed every company on those listings and
    # the name came back "". Nameless candidates are skipped before
    # screening, which is why some markets showed 2,200 screened and 0
    # accepted. The id is stable across both shapes.
    facts_by_id: dict[str, dict] = {}
    for key, org in facts.items():
        cid = kompass_client.extract_company_id(key)
        if cid:
            facts_by_id[cid.lower()] = org

    out: list[KompassCompanyCandidate] = []
    seen: set[str] = set()
    for m in _COMPANY_LINK_RE.finditer(html):
        url = urljoin(base_url, m.group(1))
        key = url.rstrip("/").lower()
        if key in seen:
            continue
        seen.add(key)
        org = facts.get(key) or {}
        if not org:
            cid = kompass_client.extract_company_id(url)
            if cid:
                org = facts_by_id.get(cid.lower()) or {}
        address = org.get("address") or {}
        out.append(
            KompassCompanyCandidate(
                name=str(org.get("name") or "").strip(),
                profile_url=url,
                country_hint=str(address.get("addressCountry") or "").strip(),
                city_hint=str(address.get("addressLocality") or "").strip(),
                phone_hint=str(org.get("telephone") or "").strip(),
                employees_hint=_employee_range(org.get("numberOfEmployees")),
            )
        )
    return out


def _employee_range(value: Any) -> str:
    """Format JSON-LD's {minValue, maxValue} employee count as "20-49"."""
    if not isinstance(value, dict):
        return str(value or "").strip()
    lo = value.get("minValue")
    hi = value.get("maxValue")
    if lo is not None and hi is not None:
        return f"{lo}-{hi}"
    return str(lo or hi or "").strip()


def _find_next_page(html: str, base_url: str) -> str | None:
    """Return the URL of the page AFTER base_url, or None if there is none.

    Handles both a conventional "Next"/"»" link and Kompass's actual
    markup, which has no next-link at all — just numbered "pageBtn" links
    to path-based pages (".../23200/page-2/"). For the numbered form we
    work out the current page number from base_url and pick the smallest
    linked page number greater than it, so pagination advances one page at
    a time even though the listing only links a sparse subset of pages
    (e.g. 2, 3, 5, 10, 15, 20, 40)."""
    m = _NEXT_PAGE_RE.search(html)
    if m:
        return urljoin(base_url, m.group(1))

    current_match = _CURRENT_PAGE_RE.search(urlparse(base_url).path)
    current = int(current_match.group(1)) if current_match else 1

    best_url: str | None = None
    best_num: int | None = None
    for href, num_text in _PAGE_LINK_RE.findall(html):
        try:
            num = int(num_text)
        except ValueError:
            continue
        if num <= current:
            continue
        if best_num is None or num < best_num:
            best_num = num
            best_url = href
    if best_url is None:
        return None
    return urljoin(base_url, best_url)


_UNLIMITED_PAGE_SAFETY_CAP = 5000  # hard safety ceiling even when max_candidates=0 (unlimited)
_LIMITED_PAGE_SAFETY_CAP = 200


def discover_distributor_candidates(
    product_url: str,
    *,
    market: str,
    max_candidates: int,
) -> list[KompassCompanyCandidate]:
    """Paginate through the Kompass product/distributor listing, collecting
    company-profile links. max_candidates=0 means NO upper limit — keeps
    paginating every listing URL variant until genuinely out of pages
    (bounded only by a generous safety cap to prevent a runaway loop, not a
    business-logic cap). Otherwise stops on: max_candidates reached, no next
    page, or a fetch failure (spec §15)."""
    unlimited = max_candidates <= 0
    page_safety_cap = _UNLIMITED_PAGE_SAFETY_CAP if unlimited else _LIMITED_PAGE_SAFETY_CAP
    countries = world_countries() if global_search_enabled() else [""]
    seen_urls: set[str] = set()
    out: list[KompassCompanyCandidate] = []

    listing_urls = _build_distributor_listing_urls(product_url)
    _log(
        f"sweeping {len(listing_urls)} listings: the global listing first, "
        f"then {len(listing_urls) - 1} per-country distributor listings"
    )
    for start_url in listing_urls:
        before = len(out)
        page_url: str | None = start_url
        pages_visited = 0
        while page_url and pages_visited < page_safety_cap and (unlimited or len(out) < max_candidates):
            page = kompass_client.fetch_kompass_page(page_url)
            pages_visited += 1
            if not page.alive:
                break
            links = _extract_company_links(page.html, page.final_url)
            new = 0
            for cand in links:
                key = cand.profile_url.rstrip("/").lower()
                if key in seen_urls:
                    continue
                seen_urls.add(key)
                out.append(cand)
                new += 1
                if not unlimited and len(out) >= max_candidates:
                    break
            if new == 0 and pages_visited > 1:
                break  # no new companies on this page — stop paginating it
            page_url = _find_next_page(page.html, page.final_url)

        found_here = len(out) - before
        if found_here:
            # Label by what the listing IS, not by its subdomain: the global
            # listing's /z/ww/ path makes it worldwide regardless of which
            # country subdomain served it, so printing e.g. "lb.kompass.com"
            # for it wrongly reads as "Lebanon only".
            if "/z/ww/" in start_url:
                source = "GLOBAL listing (/z/ww/, all countries)"
            elif "/x/distributor/" in start_url:
                source = f"{urlparse(start_url).netloc.split('.')[0].upper()} distributors"
            else:
                source = urlparse(start_url).netloc
            _sub(
                f"{source}: +{found_here} companies "
                f"({pages_visited} page(s)) — running total {len(out)}"
            )
        # Every listing URL is a DIFFERENT country's distributor listing
        # (they return disjoint company sets), so we must visit them all to
        # get global coverage — never stop early just because one country
        # already returned results. Only a max_candidates ceiling stops us.
        if not unlimited and len(out) >= max_candidates:
            break

    _log(
        f"discovered {len(out)} raw distributor candidates "
        f"across {len(listing_urls)} country listings"
    )
    return out if unlimited else out[:max_candidates]


# --------------------------------------------------------------------------
# Step 5-8: open each company profile, verify, extract fields
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class ActivityItem:
    classification_id: str
    classification_url: str
    text: str
    # Kompass's own legend: "supplier" = Producer/Manufacturer role,
    # "distributor" = Distributor role, "service" = Service providers.
    roles: tuple[str, ...]  # e.g. ("distributor",), ("supplier", "service")

    def is_distributor(self) -> bool:
        return "distributor" in self.roles

    def is_manufacturer(self) -> bool:
        """"supplier" is Kompass's internal role name for the Producer/
        Manufacturer legend entry — see the "Activities" legend markup."""
        return "supplier" in self.roles

    def is_pure_distributor(self) -> bool:
        """Distributor role WITHOUT any manufacturer/OEM/producer role on
        this same classification entry — user explicitly wants no
        manufacturer/OEM involvement at all, not even combo roles."""
        return self.is_distributor() and not self.is_manufacturer()


def _extract_activities(html: str) -> list[ActivityItem]:
    """Parse the per-classification activity tree — see _ACTIVITY_ITEM_RE's
    docstring comment above for why this is the authoritative source for
    market-specific distributor verification, not just a company-wide badge."""
    items: list[ActivityItem] = []
    for cid, url, inner in _ACTIVITY_ITEM_RE.findall(html):
        text_m = _ACTIVITY_TEXT_RE.search(inner)
        if not text_m:
            continue
        roles = tuple(dict.fromkeys(r.lower() for r in _ACTIVITY_CIRCLE_RE.findall(inner)))
        items.append(
            ActivityItem(
                classification_id=cid,
                classification_url=url,
                text=text_m.group(1).strip(),
                roles=roles,
            )
        )
    return items


def _matching_distributor_activities(activities: list[ActivityItem], market_terms: list[str]) -> list[ActivityItem]:
    """Activities that are a PURE distributor role (no manufacturer/OEM role
    on the same entry) AND textually match the requested market — the
    strongest possible evidence per spec §5/§6/§26, with the added rule that
    manufacturer/OEM involvement disqualifies the match entirely."""
    out = []
    for a in activities:
        if not a.is_pure_distributor():
            continue
        text_l = a.text.lower()
        if any(term in text_l for term in market_terms):
            out.append(a)
    return out


_JSON_LD_RE = re.compile(
    r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>', re.S | re.I
)


def _extract_organization_json_ld(html: str, expected_name: str = "") -> dict | None:
    """Kompass profiles embed a schema.org Organization JSON-LD block with
    clean, unambiguous structured data (name, identifier, foundingDate,
    numberOfEmployees, sameAs website list, address, contact info) — far more
    reliable than scraping visible text/markup. This is preferred whenever
    present; regex-on-HTML functions below are the fallback for pages that
    lack it. When multiple JSON-LD Organization blocks exist on a page (e.g.
    related/sidebar companies), prefer the one whose name matches the
    profile we actually requested, to avoid misattributing another
    company's data."""
    import json as _json

    best: dict | None = None
    for raw in _JSON_LD_RE.findall(html):
        try:
            data = _json.loads(raw)
        except Exception:
            continue
        if not isinstance(data, dict) or data.get("@type") != "Organization":
            continue
        name = str(data.get("name") or "")
        if expected_name and name.strip().lower() == expected_name.strip().lower():
            return data
        if best is None and name:
            best = data
    return best


def _extract_company_type(html: str, activities: list[ActivityItem] | None = None) -> str:
    """Overall company-type label. When the per-activity classification tree
    is present, this is the union of roles seen across all activities (e.g.
    "Distributor, Supplier, Service") — a company-level summary only; the
    market-specific distributor check does NOT rely on this string, it
    checks _matching_distributor_activities() instead (spec §5/§6/§26)."""
    if activities:
        roles_seen: list[str] = []
        for a in activities:
            for r in a.roles:
                label = {"supplier": "Manufacturer", "distributor": "Distributor", "service": "Service"}.get(
                    r, r.title()
                )
                if label not in roles_seen:
                    roles_seen.append(label)
        if roles_seen:
            return " ".join(roles_seen)
    m = _COMPANY_TYPE_BLOCK_RE.search(html)
    return re.sub(r"\s+", " ", m.group(0)).strip() if m else ""


def _extract_founded_year(html: str, org: dict | None = None) -> str:
    if org and org.get("foundingDate"):
        m = re.search(r"\d{4}", str(org["foundingDate"]))
        if m:
            return m.group(0)
    m = _FOUNDED_RE.search(html)
    return m.group(1) if m else ""


def _extract_employees(html: str, org: dict | None = None) -> str:
    if org:
        emp = org.get("numberOfEmployees")
        if isinstance(emp, dict) and emp.get("value") is not None:
            return str(emp["value"])
        if isinstance(emp, (int, str)) and str(emp).strip():
            return str(emp).strip()
    m = _EMPLOYEES_RANGE_RE.search(html)
    if m:
        return re.sub(r"\s*-\s*", "-", m.group(1)) + " Employees"
    m = _EMPLOYEES_EXACT_RE.search(html)
    if m:
        return m.group(1)
    return ""


def _extract_official_website(html: str, company_slug: str, org: dict | None = None) -> str:
    """Pick the most likely OFFICIAL site among URLs on the profile — prefer
    a domain that shares tokens with the company slug over unrelated/brand
    partner links (spec §12). The JSON-LD ``sameAs`` array, when present, is
    Kompass's own curated list of the company's real external sites and is
    preferred over scraping arbitrary URLs out of the page body."""
    slug_tokens = [t for t in re.split(r"[-_]", company_slug.lower()) if len(t) > 2]

    def score(u: str) -> int:
        host = urlparse(u).netloc.lower()
        return sum(1 for t in slug_tokens if t in host)

    if org:
        same_as = org.get("sameAs")
        urls = [u for u in same_as if isinstance(u, str)] if isinstance(same_as, list) else []
        urls = [u for u in urls if "kompass.com" not in u.lower()]
        if urls:
            urls_sorted = sorted(dict.fromkeys(urls), key=score, reverse=True)
            return urls_sorted[0]

    urls = [u.rstrip("/.,;)") for u in _WEBSITE_RE.findall(html)]
    urls = [u for u in urls if "kompass.com" not in u.lower()]
    if not urls:
        return ""
    urls_sorted = sorted(dict.fromkeys(urls), key=score, reverse=True)
    return urls_sorted[0]


def _extract_email(html: str, org: dict | None = None) -> str:
    if org and org.get("email"):
        return str(org["email"]).strip()
    m = _EMAIL_RE.search(html)
    return m.group(0) if m else ""


def _extract_phone(html: str, org: dict | None = None) -> str:
    if org and org.get("telephone"):
        return str(org["telephone"]).strip()
    m = _PHONE_RE.search(html)
    return m.group(0).strip() if m else ""


def _extract_address(org: dict | None = None) -> tuple[str, str, str]:
    """Returns (full_address, city, country) from JSON-LD PostalAddress."""
    if not org:
        return "", "", ""
    addr = org.get("address")
    if not isinstance(addr, dict):
        return "", "", ""
    parts = [
        addr.get("streetAddress"),
        addr.get("addressLocality"),
        addr.get("postalCode"),
        addr.get("addressCountry"),
    ]
    full = ", ".join(str(p) for p in parts if p)
    city = str(addr.get("addressLocality") or "")
    country = str(addr.get("addressCountry") or "")
    return full, city, country


def _extract_description(org: dict | None = None) -> str:
    if org and org.get("description"):
        return str(org["description"]).strip()
    return ""


def _market_terms(market: str) -> list[str]:
    """Lowercased, singularized significant terms from the market name, for
    evidence matching. Deliberately does not include generic connector
    words. Trailing 's' is stripped (naive singularization) so a market name
    like "Smartphones" matches page text using either "smartphone" or
    "smartphones" — the regex re-adds \\w* to match either form back."""
    m = re.sub(r"\bmarket\b", "", market, flags=re.I)
    terms = [t.lower() for t in re.findall(r"[a-zA-Z]{4,}", m)]
    stop = {"global", "industry", "products", "solutions", "sector"}
    terms = [t for t in terms if t not in stop]
    return [t[:-1] if t.endswith("s") and len(t) > 4 else t for t in terms]


def _has_market_specific_distribution_evidence(
    html: str,
    market: str,
    org: dict | None = None,
    activities: list[ActivityItem] | None = None,
) -> tuple[bool, str]:
    """This is the core anti-false-positive check (spec §5/§6/§26): merely
    finding the market keyword anywhere on the page is NOT sufficient.
    Evidence sources, strongest first:
      (a) A per-activity classification entry that is BOTH tagged
          "distributor" AND textually matches the market — Kompass's own
          structured, role-specific tagging (see ActivityItem /
          _matching_distributor_activities). This is authoritative: it
          proves the company is a distributor FOR THAT SPECIFIC
          classification, not merely a distributor of something unrelated.
      (b) Prose co-occurrence: the market term(s) appearing within a short
          window of an explicit distribution-role word — e.g. "distributor
          of smartphones" — used only as a fallback when no activity tree
          was extractable from the page (older/different page layout).
    ``knowsAbout`` (JSON-LD) and prose mentions WITHOUT a nearby role word
    are deliberately NOT treated as sufficient evidence on their own — a
    company merely "knowing about" or mentioning a market term (e.g. in a
    software/service description) is exactly the false-positive pattern the
    spec calls out.
    """
    terms = _market_terms(market)
    if not terms:
        return False, "no_market_terms"

    if activities:
        matches = _matching_distributor_activities(activities, terms)
        if matches:
            names = ", ".join(a.text for a in matches[:3])
            return True, f"Kompass activity classification (distributor role): {names}"
        # Activity tree was present but had no distributor+market match —
        # do NOT fall through to weaker prose-matching; the structured
        # data is authoritative when available.
        return False, "no Kompass activity classification tags this company as a distributor for this market"

    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text).lower()

    role_words = r"(distribut\w*|wholesal\w*|import\w*|reseller\w*|channel partner\w*|authori[sz]ed dealer\w*)"
    for term in terms:
        # role word near the market term, either order, within ~80 chars
        pattern_fwd = rf"{role_words}[^.]{{0,80}}\b{re.escape(term)}\w*"
        pattern_bwd = rf"\b{re.escape(term)}\w*[^.]{{0,80}}{role_words}"
        if re.search(pattern_fwd, text) or re.search(pattern_bwd, text):
            return True, f"co-occurrence: '{term}' near a distribution-role word"

    # Software/service-only mentions must NOT count as evidence.
    software_only = any(
        re.search(rf"\bsoftware for\b[^.]{{0,40}}{re.escape(term)}", text)
        or re.search(rf"{re.escape(term)}[^.]{{0,40}}\bsoftware\b", text)
        for term in terms
    )
    if software_only:
        return False, "market keyword present only in a software/service context"

    return False, "market keyword not found in distribution-role context or Kompass classification"


def verify_and_score(
    *,
    market: str,
    company_type: str,
    market_evidence_found: bool,
    website_confirms: bool = False,
) -> tuple[bool, int]:
    """Spec §19 scoring:
    +40 Kompass explicitly says Distributor
    +30 Kompass product classification matches requested market
    +20 official website confirms distribution
    (The original spec's +10 "manufacturer/brand partner evidence" bonus is
    intentionally removed — the user explicitly excluded manufacturer/OEM
    involvement, so rewarding manufacturer-adjacent evidence would
    contradict that.)

    Pure-distributor rule: a company carrying ANY manufacturer/OEM/producer
    role — even alongside a Distributor role (e.g. "Manufacturer
    Distributor") — is rejected outright, regardless of score. No
    manufacturer/OEM involvement is accepted, per explicit user instruction
    (stricter than the original spec, which allowed Manufacturer+Distributor
    combos)."""
    is_manufacturer = bool(_MANUFACTURER_TYPE_RE.search(company_type or ""))
    is_distributor_type = bool(_DISTRIBUTOR_TYPE_RE.search(company_type or ""))

    score = 0
    if is_distributor_type:
        score += 40
    if market_evidence_found:
        score += 30
    if website_confirms:
        score += 20

    if is_manufacturer:
        return False, score
    accepted = is_distributor_type and market_evidence_found and score >= min_distributor_confidence()
    return accepted, score


def verify_company(
    candidate: KompassCompanyCandidate,
    *,
    market: str,
    kompass_product_id: str,
) -> VerifiedDistributor | None:
    """Open one company profile, verify distributor status, extract fields.
    Returns None only on a hard fetch failure (page truly unreachable) — a
    reachable-but-rejected company still returns a VerifiedDistributor with
    is_actual_market_distributor=False, so callers can log the rejection
    reason (spec §21/§22)."""
    page = kompass_client.fetch_kompass_page(candidate.profile_url)
    if not page.alive:
        _detail(f"profile unreachable: {candidate.profile_url} ({page.error})")
        return None

    company_id = kompass_client.extract_company_id(page.final_url)
    if not company_id:
        _detail(f"could not extract Kompass company id from {page.final_url} — skipping")
        return None

    slug = kompass_client.extract_company_slug(page.final_url) or ""
    fallback_name = candidate.name or slug.replace("-", " ").title() or company_id
    org = _extract_organization_json_ld(page.html, expected_name=fallback_name)
    name = str(org.get("name")).strip() if org and org.get("name") else fallback_name
    activities = _extract_activities(page.html)

    company_type = _extract_company_type(page.html, activities)
    market_evidence_found, evidence_reason = _has_market_specific_distribution_evidence(
        page.html, market, org, activities
    )
    website = _extract_official_website(page.html, slug, org)
    address, city, country = _extract_address(org)

    accepted, score = verify_and_score(
        market=market,
        company_type=company_type,
        market_evidence_found=market_evidence_found,
    )

    result = VerifiedDistributor(
        company_name=name,
        market=market,
        kompass_company_id=company_id,
        kompass_profile_url=page.final_url,
        kompass_product_id=kompass_product_id,
        kompass_company_type=company_type or NOT_DISCLOSED,
        is_actual_market_distributor=accepted,
        verification_confidence=score,
        rejection_reason="" if accepted else evidence_reason,
        website=website or NOT_DISCLOSED,
        founded_year=_extract_founded_year(page.html, org) or NOT_DISCLOSED,
        employees=_extract_employees(page.html, org) or NOT_DISCLOSED,
        email=_extract_email(page.html, org) or NOT_DISCLOSED,
        phone=_extract_phone(page.html, org) or NOT_DISCLOSED,
        address=address or NOT_DISCLOSED,
        city=city,
        country=country,
        description=_extract_description(org),
        verification_urls=[page.final_url],
    )
    return result


# --------------------------------------------------------------------------
# Step 9: dedup
# --------------------------------------------------------------------------


def _name_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def deduplicate(distributors: list[VerifiedDistributor]) -> list[VerifiedDistributor]:
    """Primary key: kompass_company_id. Secondary fallback: normalized name
    + country (spec §14)."""
    by_id: dict[str, VerifiedDistributor] = {}
    by_name_country: dict[str, VerifiedDistributor] = {}
    out: list[VerifiedDistributor] = []
    for d in distributors:
        id_key = d.kompass_company_id.strip().upper()
        if id_key and id_key in by_id:
            continue
        nc_key = f"{_name_key(d.company_name)}|{d.country.strip().lower()}"
        if nc_key.strip("|") and nc_key in by_name_country:
            continue
        if id_key:
            by_id[id_key] = d
        by_name_country[nc_key] = d
        out.append(d)
    return out


# --------------------------------------------------------------------------
# Orchestration: one market end-to-end
# --------------------------------------------------------------------------


# Every SLIM_COLUMNS field the ChatGPT scraper can research, in export
# order. Excluded on purpose:
#   Company        — already known (that's what we screened)
#   Country Code /
#   Region Code    — derived deterministically from Headquarters at export
#                    time, so asking an LLM would only introduce errors
#   Summary        — written separately from the row's own verified data
#                    (see company_enrichment.write_company_summary), so it
#                    never invents facts that aren't in the other columns
_CHATGPT_FILL_COLUMNS = [
    "Website",
    "Founded",
    "Headquarters",
    "Continent / Geography",
    "Operational Presence",
    "Ownership",
    "Employees",
    "Core Categories",
    "Specialty Focus",
    "Key Brands Represented",
    "Online / Offline / Both",
    "Contact Person",
    "Role",
    "Email",
    "LinkedIn",
    "Office No.",
    "Office Email",
]


def _screen_candidates_via_chatgpt(
    candidates: list[KompassCompanyCandidate], *, market: str, kompass_product_id: str
) -> tuple[list[VerifiedDistributor], int, int, set[str]]:
    """Classify candidates with the ChatGPT scraper instead of fetching each
    Kompass profile, then have it fill the export columns for the accepted
    ones. Returns (accepted, rejected_count, unknown_count).

    Candidates with no name are dropped — ChatGPT can't screen a bare URL,
    and accepting them unscreened would defeat the purpose."""
    from vendor_intel.discovery.chatgpt_company_screen import (
        ScreenedCompany,
        fill_fields_via_google_ai,
        screen_roles_via_google_ai,
    )

    named = [c for c in candidates if c.name.strip()]
    skipped = len(candidates) - len(named)
    if skipped:
        _log(f"skipping {skipped} candidates with no name from the listing")
    if not named:
        return [], 0, skipped, set()

    screened = [
        ScreenedCompany(name=c.name, profile_url=c.profile_url, country_hint=c.country_hint)
        for c in named
    ]
    # Keep the listing's own facts alongside, keyed by profile URL, so the
    # accepted rows can use them instead of re-fetching each profile.
    hints = {c.profile_url: c for c in named}
    # Google AI Overview, one company per query. It cannot answer a batched
    # "classify these 25" prompt (search-backed: that returns "About 0
    # results"), so concurrency rather than batching provides the speed.
    screen_roles_via_google_ai(screened, market=market)

    distributors = [s for s in screened if s.is_distributor]
    rejected = sum(1 for s in screened if s.verdict in ("manufacturer", "other"))
    unknown = sum(1 for s in screened if s.verdict == "unknown") + skipped

    # Companies Google refused to answer about because it was rate-limiting
    # us. They were never really screened, so the caller must NOT record
    # them as processed — otherwise the checkpoint would skip them forever
    # on resume and they would silently vanish from the market.
    blocked_urls = {s.profile_url for s in screened if s.verdict == "blocked"}
    if blocked_urls:
        _log(
            f"{len(blocked_urls)} companies hit Google bot detection and were "
            f"NOT screened — they stay queued for the next run"
        )

    if distributors:
        # Seed each company with what Kompass's listing already told us,
        # BEFORE asking Google AI Overview. Two reasons: we don't pay to
        # research a field we already have, and Kompass's own data is
        # first-party — it should never be overwritten by a researched
        # guess. fill_fields() then only asks about columns still blank.
        for s in distributors:
            hint = hints.get(s.profile_url)
            if not hint:
                continue
            if hint.city_hint and hint.country_hint:
                s.fields.setdefault("Headquarters", f"{hint.city_hint}, {hint.country_hint}")
            elif hint.city_hint:
                s.fields.setdefault("Headquarters", hint.city_hint)
            if hint.employees_hint:
                s.fields.setdefault("Employees", hint.employees_hint)
            if hint.phone_hint:
                s.fields.setdefault("Office No.", hint.phone_hint)

        fill_fields_via_google_ai(
            distributors, columns=_CHATGPT_FILL_COLUMNS, market=market
        )

    accepted: list[VerifiedDistributor] = []
    for s in distributors:
        hint = hints.get(s.profile_url)
        accepted.append(
            VerifiedDistributor(
                company_name=s.name,
                market=market,
                kompass_company_id=kompass_client.extract_company_id(s.profile_url) or "",
                kompass_profile_url=s.profile_url,
                kompass_product_id=kompass_product_id,
                # Recorded honestly: this verdict came from Google AI
                # Overview's search results, NOT from Kompass's own
                # per-activity role tags, so it must not be presented as
                # Kompass evidence.
                kompass_company_type="Distributor (per Google AI Overview screening)",
                is_actual_market_distributor=True,
                verification_confidence=60,
                website=s.fields.get("Website", ""),
                headquarters=s.fields.get("Headquarters", ""),
                country=s.country_hint,
                # City comes straight from the Kompass listing. The other
                # listing facts (employees, phone, HQ) were already seeded
                # into s.fields BEFORE the research step, so they're read
                # from there rather than re-applied here.
                city=(hint.city_hint if hint else ""),
                founded_year=s.fields.get("Founded", ""),
                employees=s.fields.get("Employees", ""),
                contact_person=s.fields.get("Contact Person", ""),
                email=s.fields.get("Email", ""),
                phone=s.fields.get("Office No.", ""),
                brands=s.fields.get("Key Brands Represented", ""),
                description=s.reason,
                # Everything else researched for this company, carried
                # through to the export without needing its own field.
                extra_columns={
                    k: v
                    for k, v in s.fields.items()
                    if k
                    not in (
                        "Website",
                        "Headquarters",
                        "Founded",
                        "Employees",
                        "Contact Person",
                        "Email",
                        "Office No.",
                    )
                },
            )
        )
    return accepted, rejected, unknown, blocked_urls


_CKPT_SAVE_EVERY = 5  # persist to disk every N candidates, not every one — I/O cost vs. resume granularity
# Companies to screen between checkpoint saves. Four batches of 25 — small
# enough that a crash loses minutes rather than hours, large enough that the
# checkpoint write isn't the bottleneck.
_SCREEN_CHECKPOINT_EVERY = 100


def _verified_distributor_to_dict(d: VerifiedDistributor) -> dict[str, Any]:
    from dataclasses import asdict

    return asdict(d)


def _verified_distributor_from_dict(data: dict[str, Any]) -> VerifiedDistributor:
    return VerifiedDistributor(**data)


def run_kompass_distributor_search(
    market: str, *, ckpt: _CheckpointLike | None = None
) -> MarketDistributorResult:
    """ckpt (optional): an ExpandCheckpoint-shaped object (see
    _CheckpointLike). When given, this function checkpoints after every
    _CKPT_SAVE_EVERY candidates and, on a resumed run where the checkpoint
    already has 'kompass_product_url'/'kompass_product_id' recorded, SKIPS
    AI page discovery entirely and resumes the candidate-verification loop
    from the first not-yet-processed candidate — a killed run on candidate
    #300 of 500 does not restart from candidate #1."""
    resumed_product_url = ckpt.data("kompass_product_url") if ckpt else None
    resumed_product_id = ckpt.data("kompass_product_id") if ckpt else None

    _log(f"MARKET: {market}")
    if resumed_product_url and resumed_product_id:
        product_url, product_id = resumed_product_url, resumed_product_id
        _log(f"AI SEARCH: RESUMED (checkpoint) — skipping page discovery: {product_url}")
        _log(f"PRODUCT ID: {product_id}")
    else:
        _log("AI SEARCH: searching for Kompass product page...")
        page_info = find_market_product_page(market)
        if page_info is None:
            _log("STATUS: market not found on Kompass")
            if ckpt:
                ckpt.bump("kompass_1_discover", "market_not_found", note="kompass market not found")
            return MarketDistributorResult(market=market, market_status="kompass_market_not_found")
        product_url, product_id = page_info
        _log(f"AI SEARCH: found: {product_url}")
        _log(f"PRODUCT ID: {product_id}")
        if ckpt:
            ckpt.bump(
                "kompass_1_discover", "page_found",
                kompass_product_url=product_url, kompass_product_id=product_id,
                note="kompass product page discovered",
            )

    try:
        cached_candidates = ckpt.data("kompass_candidates") if ckpt else None
        if cached_candidates:
            _log(f"DISTRIBUTOR SEARCH: RESUMED (checkpoint) — {len(cached_candidates)} candidates already discovered")
            candidates = [KompassCompanyCandidate(**c) for c in cached_candidates]
        else:
            _log("DISTRIBUTOR SEARCH: searching global Kompass distributor results...")
            candidates = discover_distributor_candidates(
                product_url, market=market, max_candidates=max_companies_per_market()
            )
            if ckpt:
                ckpt.bump(
                    "kompass_2_candidates", "candidates_found",
                    kompass_candidates=[
                        {"name": c.name, "profile_url": c.profile_url, "country_hint": c.country_hint}
                        for c in candidates
                    ],
                    note=f"{len(candidates)} candidates discovered",
                )
    except Exception as exc:  # noqa: BLE001
        _log(f"STATUS: kompass access error — {type(exc).__name__}: {exc}")
        if ckpt:
            ckpt.bump("kompass_2_candidates", "access_error", note=f"{type(exc).__name__}: {exc}")
        return MarketDistributorResult(
            market=market,
            market_status="kompass_access_error",
            kompass_product_id=product_id,
            kompass_product_url=product_url,
        )

    processed_urls: set[str] = set(ckpt.data("kompass_processed_urls", []) if ckpt else [])
    accepted_dicts: list[dict[str, Any]] = list(ckpt.data("kompass_accepted", []) if ckpt else [])
    accepted: list[VerifiedDistributor] = [_verified_distributor_from_dict(d) for d in accepted_dicts]
    rejected_count: int = int(ckpt.data("kompass_rejected_count", 0) if ckpt else 0)
    unreachable_count: int = int(ckpt.data("kompass_unreachable_count", 0) if ckpt else 0)
    if processed_urls:
        _log(
            f"DISTRIBUTOR SEARCH: RESUMED (checkpoint) — {len(processed_urls)}/{len(candidates)} "
            f"candidates already verified ({len(accepted)} accepted so far)"
        )

    # Alternative path: skip Kompass profile fetches entirely and let the
    # ChatGPT scraper classify companies in batches instead (opt in with
    # KOMPASS_CHATGPT_SCREEN=true). Much faster and immune to Kompass's
    # profile-page rate limiting, at the cost of trading Kompass's own role
    # tags for an LLM's general knowledge.
    if screening_enabled():
        pending = [c for c in candidates if c.profile_url.rstrip("/").lower() not in processed_urls]
        if processed_urls:
            _log(
                f"SCREENING: RESUMED (checkpoint) — {len(processed_urls)}/{len(candidates)} "
                f"already screened ({len(accepted)} distributors so far)"
            )

        # Screen in chunks and checkpoint after each one. Screening the whole
        # list in a single call and saving only at the end meant a crash
        # anywhere threw away everything: observed live, a run died at
        # company 1,225 and lost all 19 distributors it had already found,
        # because the only save happened after all 9,437 completed.
        for chunk_start in range(0, len(pending), _SCREEN_CHECKPOINT_EVERY):
            chunk = pending[chunk_start : chunk_start + _SCREEN_CHECKPOINT_EVERY]
            (
                chunk_accepted,
                chunk_rejected,
                chunk_unknown,
                chunk_blocked,
            ) = _screen_candidates_via_chatgpt(
                chunk, market=market, kompass_product_id=product_id
            )
            accepted.extend(chunk_accepted)
            rejected_count += chunk_rejected
            unreachable_count += chunk_unknown
            # Companies blocked by Google's bot detection were never really
            # screened. Leaving them OUT of processed_urls is what lets a
            # later run pick them up: marking them done would drop them from
            # the market permanently on the strength of a CAPTCHA page.
            blocked_keys = {u.rstrip("/").lower() for u in chunk_blocked}
            for c in chunk:
                key = c.profile_url.rstrip("/").lower()
                if key not in blocked_keys:
                    processed_urls.add(key)

            if ckpt:
                ckpt.bump(
                    "kompass_3_verify",
                    f"screened_{len(processed_urls)}_of_{len(candidates)}",
                    kompass_processed_urls=list(processed_urls),
                    kompass_accepted=[_verified_distributor_to_dict(d) for d in accepted],
                    kompass_rejected_count=rejected_count,
                    kompass_unreachable_count=unreachable_count,
                    note=f"screened {len(processed_urls)}/{len(candidates)}",
                )

            # Stop once enough VERIFIED distributors are in hand. Screening
            # the rest would cost hours and add nothing the user asked for.
            # This never lowers the bar to reach the number: a candidate is
            # still only accepted if it passes the same distributor and
            # market checks, so a market with few real distributors simply
            # finishes under target rather than padding with manufacturers.
            target = target_distributors_per_market()
            if target and len(accepted) >= target:
                _log(
                    f"target reached — {len(accepted)} verified distributors "
                    f"(cap {target}); stopping after {len(processed_urls)}/"
                    f"{len(candidates)} candidates screened"
                )
                break
            _log(
                f"CHECKPOINT: {len(processed_urls)}/{len(candidates)} screened, "
                f"{len(accepted)} distributors saved"
            )

        deduped = deduplicate(accepted)
        # Screening runs in chunks of 100, so the target can be overshot
        # within the final chunk. Trim AFTER dedup, or removing duplicates
        # later could drop the row count back below the requested number.
        target = target_distributors_per_market()
        if target and len(deduped) > target:
            _log(f"trimming {len(deduped)} -> {target} (target cap)")
            deduped = deduped[:target]
        _log(
            f"done: {len(deduped)} distributors via ChatGPT screening "
            f"({rejected_count} rejected, {unreachable_count} unknown)"
        )
        return MarketDistributorResult(
            market=market,
            market_status="ok",
            kompass_product_id=product_id,
            kompass_product_url=product_url,
            distributors=deduped,
            rejected_count=rejected_count,
        )

    # Default path: verify against Kompass profile pages, one at a time on
    # purpose. Concurrent profile fetches were tried (10 at a time) and
    # Kompass throttled hard: 7 of 10 profiles came back HTTP 405 or
    # redirected to account.kompass.com/login, and each blocked fetch then
    # fell back to launching a browser — so the "faster" version was slower
    # AND lost companies. Sequential fetching stays within what Kompass
    # tolerates.
    since_save = 0
    for cand in candidates:
        key = cand.profile_url.rstrip("/").lower()
        if key in processed_urls:
            continue
        _log(f"CANDIDATE: {cand.name or cand.profile_url}")
        result = verify_company(cand, market=market, kompass_product_id=product_id)
        processed_urls.add(key)
        if result is None:
            unreachable_count += 1
        else:
            _log(f"KOMPASS ID: {result.kompass_company_id}")
            _log(f"COMPANY TYPE: {result.kompass_company_type}")
            if result.is_actual_market_distributor:
                _log(f"MARKET VERIFICATION: {market} distributor evidence: PASS")
                _log(f"FOUNDED: {result.founded_year}")
                _log(f"EMPLOYEES: {result.employees}")
                _log("STATUS: ACCEPTED")
                accepted.append(result)
            else:
                rejected_count += 1
                _log(f"MARKET VERIFICATION: {market} distributor evidence: FAIL")
                _log("STATUS: REJECTED")
                _log(f"REASON: {result.rejection_reason}")

        since_save += 1
        if ckpt and since_save >= _CKPT_SAVE_EVERY:
            ckpt.bump(
                "kompass_3_verify", f"verified_{len(processed_urls)}_of_{len(candidates)}",
                kompass_processed_urls=list(processed_urls),
                kompass_accepted=[_verified_distributor_to_dict(d) for d in accepted],
                kompass_rejected_count=rejected_count,
                kompass_unreachable_count=unreachable_count,
                note=f"verified {len(processed_urls)}/{len(candidates)} candidates",
            )
            since_save = 0

    if ckpt and since_save > 0:
        ckpt.bump(
            "kompass_3_verify", "verify_loop_done",
            kompass_processed_urls=list(processed_urls),
            kompass_accepted=[_verified_distributor_to_dict(d) for d in accepted],
            kompass_rejected_count=rejected_count,
            kompass_unreachable_count=unreachable_count,
            note="candidate verification loop complete",
        )

    deduped = deduplicate(accepted)
    _log(
        f"done: {len(deduped)} verified distributors "
        f"({rejected_count} rejected, {unreachable_count} unreachable)"
    )
    return MarketDistributorResult(
        market=market,
        market_status="ok",
        kompass_product_id=product_id,
        kompass_product_url=product_url,
        distributors=deduped,
        rejected_count=rejected_count,
    )


# --------------------------------------------------------------------------
# Export: SLIM_COLUMNS-shaped rows (Kompass metadata packed into Summary,
# per the existing 21-column schema used by the ChatGPT expand pipeline)
# --------------------------------------------------------------------------


def _country_display_name(iso_or_name: str) -> str:
    """JSON-LD addressCountry is typically a 2-letter ISO code (e.g. "MD")
    — resolve to a display name when possible, else pass the raw value
    through (better than silently dropping an unrecognized code)."""
    from vendor_intel.export.landscape_slim import iso_to_country_name

    val = (iso_or_name or "").strip()
    if len(val) == 2 and val.isalpha():
        name = iso_to_country_name(val.upper())
        return name if name != val.upper() else val
    return val


def to_slim_row(d: VerifiedDistributor) -> dict[str, str]:
    row = {h: "" for h in SLIM_COLUMNS}
    row["Company"] = d.company_name
    row["Website"] = d.website
    row["Founded"] = d.founded_year
    country_display = _country_display_name(d.country)
    if d.city and country_display:
        row["Headquarters"] = f"{d.city}, {country_display}"
    else:
        row["Headquarters"] = d.headquarters or country_display or NOT_DISCLOSED
    row["Employees"] = d.employees
    row["Key Brands Represented"] = d.brands or NOT_DISCLOSED
    row["Core Categories"] = d.market
    row["Contact Person"] = d.contact_person
    row["Email"] = d.email
    row["Office No."] = d.phone
    # Researched values for columns with no dedicated dataclass field.
    # Applied last so they win over the defaults set above (e.g. a real
    # "Key Brands Represented" replaces the NOT_DISCLOSED placeholder), but
    # only when non-empty — never blanks out a value we already have.
    # Country/Region Code are DERIVED from Headquarters, not researched.
    # to_slim_row() never set them, so both sat at 0/431 in a live export
    # even though Headquarters was 100% filled — pure computation that was
    # simply missing.
    _hq = str(row.get("Headquarters") or "").strip()
    if _hq and _hq != NOT_DISCLOSED:
        try:
            from vendor_intel.export.landscape_slim import country_and_region_codes

            cc, rc = country_and_region_codes(_hq)
            if cc:
                row["Country Code"] = cc
            if rc:
                row["Region Code"] = rc
        except Exception:
            pass  # codes are a convenience; never fail an export over them

    for column, value in (d.extra_columns or {}).items():
        text = str(value or "").strip()
        if text and column in row:
            row[column] = text
    summary_parts = [
        f"Kompass Company ID: {d.kompass_company_id}",
        f"Kompass Product ID: {d.kompass_product_id}",
        f"Kompass Profile: {d.kompass_profile_url}",
        f"Company Type: {d.kompass_company_type}",
        f"Market: {d.market}",
        f"Distributor Verification Score: {d.verification_confidence}/100",
        f"Verified Market Distributor: {'yes' if d.is_actual_market_distributor else 'no'}",
    ]
    if d.description:
        summary_parts.append(f"Description: {d.description}")
    row["Summary"] = " | ".join(summary_parts)
    return row


def market_result_to_rows(result: MarketDistributorResult) -> list[dict[str, str]]:
    return [to_slim_row(d) for d in result.distributors]


# --------------------------------------------------------------------------
# Final verification gate — no manufacturer/OEM, ever, regardless of source
# --------------------------------------------------------------------------

_SUMMARY_COMPANY_TYPE_RE = re.compile(r"Company Type:\s*([^|]+)", re.I)
# Ownership/description phrasing that indicates a manufacturer/OEM role even
# when no clean "Company Type:" field exists (e.g. rows from the ChatGPT/
# DeepSeek expand-pipeline fallback, which has no such field at all — that
# path only had a soft LLM-prompt instruction, never a deterministic check).
_MANUFACTURER_TEXT_RE = re.compile(
    r"\b(manufactures?|manufacturing|manufacturer|original equipment manufacturer|\boem\b|"
    r"producer of|produces|in-house production)\b",
    re.I,
)
_DISTRIBUTOR_TEXT_RE = re.compile(
    r"\b(distributor|distributes|distribution partner|wholesaler|reseller|"
    r"channel partner|authori[sz]ed dealer|importer)\b",
    re.I,
)


def row_company_type_text(row: dict[str, str]) -> str:
    """Best-effort recovery of a company-type signal from a SLIM_COLUMNS row,
    regardless of which source produced it. SLIM_COLUMNS has no dedicated
    "Company Type" column, so this checks every place that information can
    show up: Kompass rows embed "Company Type: ..." in Summary (see
    to_slim_row); ChatGPT/DeepSeek/Smart-Crawl rows have no such field, so
    Ownership/Ownership-adjacent free text and Summary itself are scanned
    for manufacturer/distributor language instead."""
    summary = str(row.get("Summary") or "")
    m = _SUMMARY_COMPANY_TYPE_RE.search(summary)
    if m:
        return m.group(1).strip()
    return " ".join(
        str(row.get(c) or "") for c in ("Ownership", "Summary", "Specialty Focus", "Core Categories")
    )


def is_manufacturer_or_oem_row(row: dict[str, str]) -> bool:
    """Deterministic, source-agnostic version of verify_and_score()'s
    manufacturer/OEM veto (spec: pure distributors only, no manufacturer/OEM
    involvement at all). Unlike verify_and_score(), which only runs inside
    the Kompass path, this is meant to be the FINAL gate applied to every
    row before export, regardless of which of the 5 enrichment sources
    (Kompass, ChatGPT scraper, Google AI, DeepSeek, Smart Crawl) produced or
    touched it — closing the gap where the ChatGPT/DeepSeek expand-pipeline
    fallback only had a soft prompt instruction, not a deterministic check."""
    type_text = row_company_type_text(row)
    if not type_text.strip():
        return False  # no signal at all — do not reject on absence of evidence
    has_manufacturer = bool(_MANUFACTURER_TYPE_RE.search(type_text)) or bool(
        _MANUFACTURER_TEXT_RE.search(type_text)
    )
    if not has_manufacturer:
        return False
    # A row can legitimately mention "manufacturer" only in reference to
    # brands/products it distributes (e.g. "distributes products from various
    # manufacturers") — that is NOT the company itself being a manufacturer.
    # Require the manufacturer language to NOT be immediately explained away
    # by an adjacent distributor phrase in the same short window; when the
    # structured Kompass "Company Type:" field is available, trust it as-is
    # (already role-only text, no such ambiguity).
    if _SUMMARY_COMPANY_TYPE_RE.search(str(row.get("Summary") or "")):
        return has_manufacturer
    return has_manufacturer and not _DISTRIBUTOR_TEXT_RE.search(type_text)


def apply_final_distributor_gate(
    rows: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """Run the final manufacturer/OEM exclusion over a full row set,
    regardless of source. Returns (kept, rejected). Call this once, right
    before a distributor-only market's rows are considered final — e.g. at
    the end of run_kompass_backend(), after both the Kompass pass and any
    ChatGPT/DeepSeek fallback top-up have run."""
    kept: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []
    for row in rows:
        if is_manufacturer_or_oem_row(row):
            rejected.append(row)
            _log(
                f"FINAL VERIFY: REJECTED {row.get('Company', '(unnamed)')!r} — "
                f"manufacturer/OEM signal detected post-enrichment"
            )
        else:
            kept.append(row)
    _log(
        f"FINAL VERIFY: {len(kept)} kept, {len(rejected)} rejected "
        f"(manufacturer/OEM exclusion, all sources)"
    )
    return kept, rejected
