"""Detect and merge missing firmographic fields (gap-fill only)."""
from __future__ import annotations

import re
from typing import Any

from vendor_intel.export.fmcg_landscape import crawl_snapshot, is_plausible_location

PLACEHOLDERS = frozenset(
    {
        "",
        "not publicly disclosed",
        "not specified",
        "unknown",
        "n/a",
        "na",
        "-",
        "none",
        "tbd",
        "null",
    }
)


def norm(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_empty_value(value: Any) -> bool:
    return norm(value).lower() in PLACEHOLDERS


def _crawl(row: dict[str, Any]) -> dict[str, Any]:
    c = row.get("_crawl")
    return c if isinstance(c, dict) else {}


def has_headquarters(row: dict[str, Any]) -> bool:
    crawl = _crawl(row)
    for key in ("headquarters", "contact_address"):
        val = norm(crawl.get(key))
        if val and is_plausible_location(val, allow_country_only=True):
            return True
    return False


def has_founded(row: dict[str, Any]) -> bool:
    return not is_empty_value(_crawl(row).get("founded_year"))


def has_employees(row: dict[str, Any]) -> bool:
    return not is_empty_value(_crawl(row).get("employee_count"))


def has_revenue(row: dict[str, Any]) -> bool:
    return not is_empty_value(_crawl(row).get("revenue"))


def has_offices(row: dict[str, Any]) -> bool:
    crawl = _crawl(row)
    for key in ("offices", "manufacturing_locations"):
        val = norm(crawl.get(key))
        if val and is_plausible_location(val, allow_country_only=False):
            return True
    return False


def has_linkedin(row: dict[str, Any]) -> bool:
    crawl = _crawl(row)
    for key in ("linkedin",):
        val = norm(crawl.get(key) or row.get(key))
        if val and "linkedin.com" in val.lower():
            return True
    raw = row.get("_crawl_raw") or {}
    contact = (raw.get("data") or {}).get("contact") or {}
    val = norm(contact.get("linkedin"))
    return bool(val and "linkedin.com" in val.lower())


def has_email(row: dict[str, Any]) -> bool:
    val = norm(_crawl(row).get("email"))
    return bool(val and "@" in val and not is_empty_value(val))


def has_phone(row: dict[str, Any]) -> bool:
    val = norm(_crawl(row).get("phone"))
    digits = re.sub(r"\D", "", val)
    return len(digits) >= 8


def has_contact_person(row: dict[str, Any]) -> bool:
    val = norm(_crawl(row).get("contact_person"))
    return bool(val and len(val) >= 3 and not is_empty_value(val))


def has_core_categories(row: dict[str, Any]) -> bool:
    val = norm(_crawl(row).get("core_categories") or row.get("key_products"))
    return bool(val and len(val) >= 3 and not is_empty_value(val))


def has_specialty_focus(row: dict[str, Any]) -> bool:
    val = norm(_crawl(row).get("specialty_focus"))
    if not val or is_empty_value(val):
        return False
    # Placeholder defaults from export are not "filled"
    low = val.lower()
    if low in {"industry manufacturer", "service provider", "distributor"}:
        return False
    return len(val) >= 3


def has_key_brands(row: dict[str, Any]) -> bool:
    val = norm(_crawl(row).get("key_brands_represented"))
    return bool(val and len(val) >= 2 and not is_empty_value(val))


def normalize_turnover_range(value: Any) -> str:
    """Convert a single revenue figure into a range band.

    Examples:
      $100M  → $80–100M
      50 million USD → $40–50M
      $1.2B → $960M–1.2B (or $0.96–1.2B)
      already '$80–100M' → kept
    """
    raw = norm(value)
    if not raw or is_empty_value(raw):
        return ""
    # Already a range
    if re.search(r"\d[\d,.]*.{0,8}(?:–|-|to)\s*\d", raw, re.I):
        return raw

    text = raw.replace(",", "").strip()
    m = re.search(
        r"(?P<sym>\$|USD|US\$|€|£)?\s*(?P<num>\d+(?:\.\d+)?)\s*(?P<unit>billion|million|bn|m|k|b)?",
        text,
        re.I,
    )
    if not m:
        return raw

    num = float(m.group("num"))
    unit_raw = (m.group("unit") or "").lower()
    sym = (m.group("sym") or "$").upper().replace("US$", "$").replace("USD", "$")
    if sym not in {"$", "€", "£"}:
        sym = "$"

    mult = 1.0
    unit_label = "M"
    if unit_raw in ("billion", "bn", "b"):
        mult = 1000.0  # work in millions
        unit_label = "B"
        # keep num in billions for display if >= 1B
    elif unit_raw in ("million", "m"):
        mult = 1.0
        unit_label = "M"
    elif unit_raw in ("k",):
        # express as millions if large, else keep K
        if num >= 1000:
            num = num / 1000.0
            unit_label = "M"
        else:
            unit_label = "K"
    else:
        # bare number — assume millions if >= 1 and looks like revenue
        unit_label = "M"

    # Convert to a common "millions" scale for banding when unit is B
    if unit_label == "B":
        high = num
        low = num * 0.8
        def _fmt(x: float) -> str:
            if x >= 10:
                return f"{x:.0f}"
            if x >= 1:
                return f"{x:.1f}".rstrip("0").rstrip(".")
            return f"{x:.2f}".rstrip("0").rstrip(".")
        return f"{sym}{_fmt(low)}–{_fmt(high)}B"

    if unit_label == "K":
        high = num
        low = num * 0.8
        return f"{sym}{low:.0f}–{high:.0f}K"

    # Millions: user wants e.g. 100M → 80–100
    high = num
    low = num * 0.8

    def _fmt_m(x: float) -> str:
        if abs(x - round(x)) < 0.05:
            return str(int(round(x)))
        return f"{x:.1f}".rstrip("0").rstrip(".")

    return f"{sym}{_fmt_m(low)}–{_fmt_m(high)}M"


CONTACT_SOURCES = frozenset({"contact_page", "leadership"})


def uses_contact_sources(sources: tuple[str, ...] | None) -> bool:
    return bool(sources and CONTACT_SOURCES.intersection(sources))


def missing_contact_fields(row: dict[str, Any]) -> list[str]:
    gaps: list[str] = []
    if not has_email(row):
        gaps.append("email")
    if not has_phone(row):
        gaps.append("phone")
    if not has_contact_person(row):
        gaps.append("contact_person")
    return gaps


def missing_fields(row: dict[str, Any]) -> list[str]:
    gaps: list[str] = []
    if not has_headquarters(row):
        gaps.append("headquarters")
    if not has_founded(row):
        gaps.append("founded_year")
    if not has_employees(row):
        gaps.append("employee_count")
    if not has_revenue(row):
        gaps.append("revenue")
    if not has_offices(row):
        gaps.append("offices")
    if not has_linkedin(row):
        gaps.append("linkedin")
    if not has_core_categories(row):
        gaps.append("core_categories")
    if not has_specialty_focus(row):
        gaps.append("specialty_focus")
    if not has_key_brands(row):
        gaps.append("key_brands_represented")
    if not has_contact_person(row):
        gaps.append("contact_person")
    return gaps


def row_has_gaps(row: dict[str, Any], *, sources: tuple[str, ...] | None = None) -> bool:
    gaps = list(missing_fields(row))
    if uses_contact_sources(sources):
        for field in missing_contact_fields(row):
            if field not in gaps:
                gaps.append(field)
    return bool(gaps)


def sync_crawl_from_raw(row: dict[str, Any]) -> None:
    snap = crawl_snapshot(row.get("_crawl_raw"))
    row["_crawl"] = {**snap, **{k: v for k, v in (_crawl(row)).items() if v}}


def merge_extracted(row: dict[str, Any], extracted: dict[str, Any], *, source: str) -> list[str]:
    """Merge LLM/search extracted fields into row (gap-fill only). Returns field names updated."""
    if not extracted:
        return []
    crawl = dict(_crawl(row))
    updates: list[str] = []

    hq = norm(extracted.get("headquarters"))
    if hq and not has_headquarters(row) and is_plausible_location(hq, allow_country_only=True):
        crawl["headquarters"] = hq
        updates.append("headquarters")

    founded = extracted.get("founded_year")
    if founded and not has_founded(row):
        try:
            crawl["founded_year"] = int(str(founded).strip()[:4])
            updates.append("founded_year")
        except (TypeError, ValueError):
            pass

    emp = norm(extracted.get("employee_count"))
    if emp and not is_empty_value(emp) and not has_employees(row):
        crawl["employee_count"] = emp
        updates.append("employee_count")

    rev = norm(extracted.get("revenue") or extracted.get("turnover"))
    if rev and not is_empty_value(rev) and not has_revenue(row):
        crawl["revenue"] = normalize_turnover_range(rev) or rev
        updates.append("revenue")

    offices = norm(extracted.get("offices") or extracted.get("cities_regions"))
    if offices and not has_offices(row) and is_plausible_location(offices, allow_country_only=False):
        crawl["offices"] = offices
        updates.append("offices")

    linkedin = norm(extracted.get("linkedin_url") or extracted.get("linkedin"))
    if linkedin and not has_linkedin(row) and "linkedin.com" in linkedin.lower():
        if not linkedin.startswith("http"):
            linkedin = "https://" + linkedin.lstrip("/")
        crawl["linkedin"] = linkedin
        row["linkedin"] = linkedin
        updates.append("linkedin")

    core = norm(extracted.get("core_categories") or extracted.get("products"))
    if core and not has_core_categories(row):
        crawl["core_categories"] = core
        updates.append("core_categories")

    specialty = norm(extracted.get("specialty_focus"))
    if specialty and not has_specialty_focus(row):
        crawl["specialty_focus"] = specialty
        updates.append("specialty_focus")

    brands = norm(
        extracted.get("key_brands_represented")
        or extracted.get("key_brands")
        or extracted.get("brands")
    )
    if brands and not has_key_brands(row):
        crawl["key_brands_represented"] = brands
        updates.append("key_brands_represented")

    person = norm(extracted.get("contact_person"))
    if person and len(person) >= 3 and not has_contact_person(row):
        crawl["contact_person"] = person
        updates.append("contact_person")

    role = norm(extracted.get("contact_role") or extracted.get("role"))
    if role and not norm(crawl.get("contact_role")):
        crawl["contact_role"] = role
        updates.append("contact_role")

    if not updates:
        return []

    row["_crawl"] = crawl
    raw = dict(row.get("_crawl_raw") or {})
    data = dict(raw.get("data") or {})
    company = dict(data.get("company") or {})
    location = dict(data.get("location") or {})
    financials = dict(data.get("financials") or {})
    contact = dict(data.get("contact") or {})
    people = dict(data.get("people") or {}) if isinstance(data.get("people"), dict) else {}

    if "headquarters" in updates:
        location["headquarters"] = crawl["headquarters"]
    if "founded_year" in updates:
        company["founded_year"] = crawl["founded_year"]
    if "employee_count" in updates:
        financials["employee_count"] = crawl["employee_count"]
    if "revenue" in updates:
        financials["revenue"] = crawl["revenue"]
    if "offices" in updates:
        loc_off = location.get("offices")
        if isinstance(loc_off, list):
            loc_off = list(loc_off)
        elif loc_off:
            loc_off = [str(loc_off)]
        else:
            loc_off = []
        for part in crawl["offices"].split(","):
            p = part.strip()
            if p and p not in loc_off:
                loc_off.append(p)
        location["offices"] = loc_off
    if "linkedin" in updates:
        contact["linkedin"] = crawl["linkedin"]
    if "core_categories" in updates:
        company["core_categories"] = crawl["core_categories"]
        company["products"] = crawl["core_categories"]
    if "specialty_focus" in updates:
        company["specialty_focus"] = crawl["specialty_focus"]
    if "key_brands_represented" in updates:
        company["key_brands_represented"] = crawl["key_brands_represented"]
    if "contact_person" in updates or "contact_role" in updates:
        leadership = list(people.get("leadership") or [])
        if not isinstance(leadership, list):
            leadership = []
        entry = {
            "name": crawl.get("contact_person") or "",
            "title": crawl.get("contact_role") or "",
        }
        if entry["name"] and not any(
            isinstance(p, dict) and norm(p.get("name")).lower() == entry["name"].lower()
            for p in leadership
        ):
            leadership.insert(0, entry)
        people["leadership"] = leadership
        contact["name"] = crawl.get("contact_person") or contact.get("name") or ""

    data["company"] = company
    data["location"] = location
    data["financials"] = financials
    data["contact"] = contact
    data["people"] = people
    raw["data"] = data
    row["_crawl_raw"] = raw

    src = str(row.get("enrichment_source") or "")
    if source and source not in src:
        row["enrichment_source"] = f"{src}+{source}".strip("+")

    sync_crawl_from_raw(row)
    return updates


def merge_contact_extracted(row: dict[str, Any], extracted: dict[str, Any], *, source: str) -> list[str]:
    """Merge contact / leadership fields (gap-fill only)."""
    if not extracted:
        return []
    crawl = dict(_crawl(row))
    updates: list[str] = []

    email = norm(extracted.get("email"))
    if email and "@" in email and not has_email(row):
        crawl["email"] = email.lower()
        updates.append("email")

    phone = norm(extracted.get("phone"))
    if phone and len(re.sub(r"\D", "", phone)) >= 8 and not has_phone(row):
        crawl["phone"] = phone
        updates.append("phone")

    person = norm(extracted.get("contact_person"))
    if person and len(person) >= 3 and not has_contact_person(row):
        crawl["contact_person"] = person
        updates.append("contact_person")

    role = norm(extracted.get("contact_role"))
    if role and not norm(crawl.get("contact_role")):
        crawl["contact_role"] = role
        updates.append("contact_role")

    addr = norm(extracted.get("contact_address"))
    if addr and not norm(crawl.get("contact_address")) and is_plausible_location(addr, allow_country_only=True):
        crawl["contact_address"] = addr
        updates.append("contact_address")

    for url_key in ("contact_source_url", "leadership_source_url"):
        url = norm(extracted.get(url_key))
        if url:
            crawl[url_key] = url

    if not updates:
        return []

    row["_crawl"] = crawl
    raw = dict(row.get("_crawl_raw") or {})
    data = dict(raw.get("data") or {})
    contact = dict(data.get("contact") or {})
    people = dict(data.get("people") or {}) if isinstance(data.get("people"), dict) else {}

    if "email" in updates:
        contact["email"] = crawl["email"]
    if "phone" in updates:
        contact["phone"] = crawl["phone"]
    if "contact_address" in updates:
        contact["address"] = crawl["contact_address"]
    if "contact_person" in updates or "contact_role" in updates:
        leadership = list(people.get("leadership") or [])
        if not isinstance(leadership, list):
            leadership = []
        entry = {
            "name": crawl.get("contact_person") or "",
            "title": crawl.get("contact_role") or "",
        }
        if entry["name"] and not any(
            isinstance(p, dict) and norm(p.get("name")).lower() == entry["name"].lower() for p in leadership
        ):
            leadership.insert(0, entry)
        people["leadership"] = leadership

    data["contact"] = contact
    data["people"] = people
    raw["data"] = data
    row["_crawl_raw"] = raw

    src = str(row.get("enrichment_source") or "")
    if source and source not in src:
        row["enrichment_source"] = f"{src}+{source}".strip("+")

    sync_crawl_from_raw(row)
    return updates


def is_apollo_protected_row(row: dict[str, Any], protected_keys: frozenset[str]) -> bool:
    """Row already filled by a prior Apollo enrich run — do not call Apollo again."""
    if company_key(row) in protected_keys:
        return True
    crawl = _crawl(row)
    if norm(crawl.get("apollo_person_id")):
        return True
    return False


def is_apollo_found_contact(row: dict[str, Any], cache_entry: dict[str, Any] | None = None) -> bool:
    """Row has a contact person that Apollo identified (this run or prior enrich)."""
    if not has_contact_person(row):
        return False
    crawl = _crawl(row)
    if norm(crawl.get("apollo_person_id")):
        return True
    if "apollo" in str(row.get("enrichment_source") or "").lower():
        return True
    cached = cache_entry if isinstance(cache_entry, dict) else {}
    if cached.get("contact_person") and not cached.get("error") and not cached.get("skipped"):
        return True
    return False


def needs_apollo_fields(row: dict[str, Any], fields: frozenset[str]) -> bool:
    if "contact_person" in fields and not has_contact_person(row):
        return True
    if "email" in fields and not has_email(row):
        return True
    if "linkedin" in fields and not has_linkedin(row):
        return True
    if "contact_role" in fields and is_empty_value(_crawl(row).get("contact_role")):
        return True
    if "phone" in fields and not has_phone(row):
        return True
    return False


def merge_apollo_contact(
    row: dict[str, Any],
    extracted: dict[str, Any],
    *,
    source: str = "apollo",
    name_only: bool = False,
    fill_empty_only: bool = True,
    fields: frozenset[str] | None = None,
) -> list[str]:
    """Merge Apollo contact data into a company row.

    When ``name_only=True``, only ``contact_person`` is written.
    ``fields`` limits which columns are written (e.g. contact_person + email + linkedin).
    When ``fill_empty_only=True`` (default), existing values are not overwritten.
    """
    if not extracted or extracted.get("error"):
        return []
    if fields is None:
        fields = frozenset({"contact_person"}) if name_only else frozenset(
            {"contact_person", "contact_role", "email", "linkedin", "phone"}
        )
    crawl = dict(_crawl(row))
    updates: list[str] = []
    existing_person = norm(crawl.get("contact_person"))

    if "contact_person" in fields:
        person = norm(extracted.get("contact_person"))
        allow_obfuscated = extracted.get("apollo_name_obfuscated") == "true"
        if person and len(person) >= 3 and (allow_obfuscated or "*" not in person):
            if fill_empty_only and has_contact_person(row):
                # Gap-fill: never downgrade full → partial; do upgrade partial → full.
                if "*" not in existing_person:
                    person = ""
                elif "*" in existing_person and "*" in person:
                    person = ""
            elif existing_person and "*" not in existing_person and "*" in person:
                person = ""
            if person:
                crawl["contact_person"] = person
                updates.append("contact_person")

    if name_only or fields == frozenset({"contact_person"}):
        if not updates:
            return []
        row["_crawl"] = crawl
        raw = dict(row.get("_crawl_raw") or {})
        data = dict(raw.get("data") or {})
        people = dict(data.get("people") or {}) if isinstance(data.get("people"), dict) else {}
        leadership = [
            {
                "name": crawl["contact_person"],
                "title": crawl.get("contact_role") or "",
                "source": source,
            }
        ]
        people["leadership"] = leadership
        people["primary_contact"] = leadership[0]
        data["people"] = people
        raw["data"] = data
        row["_crawl_raw"] = raw
        src = str(row.get("enrichment_source") or "")
        tag = source or "apollo"
        if tag not in src.split("+"):
            row["enrichment_source"] = f"{src}+{tag}".strip("+")
        sync_crawl_from_raw(row)
        return updates

    if "contact_role" in fields:
        role = norm(extracted.get("contact_role"))
        if role and (not fill_empty_only or is_empty_value(crawl.get("contact_role"))):
            crawl["contact_role"] = role
            updates.append("contact_role")

    if "linkedin" in fields:
        linkedin = norm(extracted.get("linkedin"))
        if linkedin and "linkedin.com" in linkedin.lower():
            if not fill_empty_only or not has_linkedin(row):
                if not linkedin.startswith("http"):
                    linkedin = "https://" + linkedin.lstrip("/")
                crawl["linkedin"] = linkedin
                row["linkedin"] = linkedin
                updates.append("linkedin")

    if "email" in fields:
        email = norm(extracted.get("email"))
        if email and "@" in email:
            if not fill_empty_only or not has_email(row):
                crawl["email"] = email.lower()
                updates.append("email")

    if "phone" in fields:
        phone = norm(extracted.get("phone"))
        if phone and len(re.sub(r"\D", "", phone)) >= 8:
            if not fill_empty_only or not has_phone(row):
                crawl["phone"] = phone
                updates.append("phone")

    apollo_id = norm(extracted.get("apollo_person_id"))
    if apollo_id:
        crawl["apollo_person_id"] = apollo_id
    apollo_url = norm(extracted.get("apollo_source_url"))
    if apollo_url:
        crawl["apollo_source_url"] = apollo_url
        crawl["contact_source_url"] = apollo_url

    if not updates:
        return []

    row["_crawl"] = crawl
    raw = dict(row.get("_crawl_raw") or {})
    data = dict(raw.get("data") or {})
    contact = dict(data.get("contact") or {})
    people = dict(data.get("people") or {}) if isinstance(data.get("people"), dict) else {}

    if "email" in updates:
        contact["email"] = crawl["email"]
    if "phone" in updates:
        contact["phone"] = crawl["phone"]
    if "linkedin" in updates:
        contact["linkedin"] = crawl["linkedin"]
    if "contact_person" in updates or "contact_role" in updates:
        leadership = [
            {
                "name": crawl.get("contact_person") or "",
                "title": crawl.get("contact_role") or "",
                "source": source,
            }
        ]
        people["leadership"] = leadership
        people["primary_contact"] = leadership[0]

    data["contact"] = contact
    data["people"] = people
    raw["data"] = data
    row["_crawl_raw"] = raw

    src = str(row.get("enrichment_source") or "")
    tag = source or "apollo"
    if tag not in src.split("+"):
        row["enrichment_source"] = f"{src}+{tag}".strip("+")

    sync_crawl_from_raw(row)
    return updates


def company_key(row: dict[str, Any]) -> str:
    return norm(row.get("company") or row.get("brand")).lower()


def domain_from_row(row: dict[str, Any]) -> str:
    raw = norm(row.get("domain") or row.get("website") or "")
    return raw.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0]
