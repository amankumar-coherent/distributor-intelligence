"""21-column output schema for the market research pipeline.

Distinct from ``vendor_intel.export.landscape_slim.SLIM_COLUMNS`` (a different,
already-consumed schema used elsewhere in this repo) — this schema matches the
user's exact spec: a leading row number, "Key Products / Technologies" instead
of "Key Brands Represented", a separate "Office Website" column, a combined
"Country Code / Region Code" column, and no "Summary" column.

Geo helpers (``hq_to_iso``, ``iso_to_region``, ``format_continent_geography``,
``parse_country_list``, ``operational_presence_text``) are imported from
``landscape_slim`` rather than reimplemented — they already encode "only
derive from real HQ/geo text, never invent" per that module's docstring.
"""
from __future__ import annotations

import re
from typing import Any

from vendor_intel.export.landscape_slim import (
    format_continent_geography,
    hq_to_iso,
    iso_to_region,
    operational_presence_text,
    parse_country_list,
)

COLUMNS: tuple[str, ...] = (
    "#",
    "Company",
    "Website",
    "Founded",
    "Headquarters",
    "Continent / Geography",
    "Operational Presence",
    "Ownership",
    "Employees",
    "Core Categories",
    "Specialty Focus",
    "Key Products / Technologies",
    "Online / Offline / Both",
    "Contact Person",
    "Role",
    "Email",
    "LinkedIn",
    "Office No.",
    "Office Website",
    "Office Email",
    "Country Code / Region Code",
)

# The six contact fields may legitimately be EMPTY.
#
# They were previously required, which forced the discovery prompt to
# invent them ("best plausible" contact, info@<domain> email, a LinkedIn
# slug built from a name, a phone number from the country calling code) so
# that rows would not be dropped. The result was measurable fabrication:
# 94% of exported Office Email values were info@/sales@ at the company's
# own domain. Per explicit user instruction ("no guessed, only filled real
# data on these columns"), an unverified contact field is now left blank
# instead — a wrong phone number or email is worse than an empty cell,
# while the firmographic columns stay mandatory so a row is still a
# complete company profile.
CONTACT_FIELDS: tuple[str, ...] = (
    "Contact Person",
    "Role",
    "Email",
    "LinkedIn",
    "Office No.",
    "Office Email",
)

# Every column except "#" (assigned by the batch runner) and the contact
# fields must be populated with a real value or the company is dropped.
REQUIRED_FIELDS: tuple[str, ...] = tuple(
    c for c in COLUMNS if c != "#" and c not in CONTACT_FIELDS
)

# --- market-dependent columns -------------------------------------------

# "Online / Offline / Both" describes a SALES CHANNEL that only exists for
# physical goods: a distributor either ships stock from a warehouse, sells
# over the counter, or does both. A CPaaS or SaaS reseller has no such
# distinction -- the product is delivered over the internet by definition,
# so every row in a software market either says "Online" (telling you
# nothing) or gets invented to fill a mandatory cell.
#
# Dropping the column for these markets is better than carrying a constant:
# it is one fewer field the prompt must ask for, and one fewer required
# field that can drop an otherwise-good company for failing to answer a
# question that has no meaningful answer.
_CHANNEL_COLUMN = "Online / Offline / Both"


def is_software_market(market: str) -> bool:
    """True for markets where goods are delivered digitally.

    Shares discovery's hint list rather than keeping a second copy, so a
    market routed to the software DISTRIBUTOR definition is always the
    same set routed away from the physical-channel column.
    """
    from vendor_intel.pipeline.market_research.discovery import (
        _SOFTWARE_MARKET_HINTS,
    )

    low = (market or "").lower()
    return any(hint in low for hint in _SOFTWARE_MARKET_HINTS)


def columns_for(market: str) -> tuple[str, ...]:
    """Column list for this market, minus any that do not apply to it."""
    if is_software_market(market):
        return tuple(c for c in COLUMNS if c != _CHANNEL_COLUMN)
    return COLUMNS


def required_fields_for(market: str) -> tuple[str, ...]:
    """Mandatory fields for this market.

    A column that is not exported must not be able to drop a row.
    """
    cols = columns_for(market)
    return tuple(c for c in cols if c != "#" and c not in CONTACT_FIELDS)

_PLACEHOLDER_VALUES = frozenset(
    {
        "",
        "n/a",
        "na",
        "none",
        "null",
        "unknown",
        "not available",
        "not disclosed",
        "not publicly disclosed",
        "tbd",
        "-",
        "--",
    }
)


def is_placeholder(value: Any) -> bool:
    """True when `value` is empty/N/A/TBD/etc. — never acceptable in a final row."""
    if value is None:
        return True
    return str(value).strip().lower() in _PLACEHOLDER_VALUES


# ISO 3166-1 alpha-2 -> international calling code. Public, static, factual
# data (not a guess) — used only to format the "Country Code / Region Code"
# column for countries the row's own verified Headquarters/Operational
# Presence already resolved via landscape_slim.hq_to_iso/parse_country_list.
# Never used to invent a country that isn't already in the row's real data.
CALLING_CODE: dict[str, str] = {
    "US": "+1", "CA": "+1", "GB": "+44", "IE": "+353", "FR": "+33", "DE": "+49",
    "IT": "+39", "ES": "+34", "PT": "+351", "NL": "+31", "BE": "+32", "LU": "+352",
    "CH": "+41", "AT": "+43", "SE": "+46", "NO": "+47", "DK": "+45", "FI": "+358",
    "IS": "+354", "PL": "+48", "CZ": "+420", "SK": "+421", "HU": "+36", "RO": "+40",
    "BG": "+359", "GR": "+30", "TR": "+90", "RU": "+7", "UA": "+380", "BY": "+375",
    "LT": "+370", "LV": "+371", "EE": "+372", "HR": "+385", "SI": "+386", "RS": "+381",
    "MT": "+356", "CY": "+357",
    "CN": "+86", "JP": "+81", "KR": "+82", "TW": "+886", "HK": "+852", "MO": "+853",
    "IN": "+91", "PK": "+92", "BD": "+880", "LK": "+94", "NP": "+977",
    "SG": "+65", "MY": "+60", "TH": "+66", "VN": "+84", "PH": "+63", "ID": "+62",
    "KH": "+855", "LA": "+856", "MM": "+95", "BN": "+673",
    "AU": "+61", "NZ": "+64",
    "AE": "+971", "SA": "+966", "QA": "+974", "KW": "+965", "BH": "+973",
    "OM": "+968", "IL": "+972", "JO": "+962", "LB": "+961", "IQ": "+964",
    "EG": "+20", "ZA": "+27", "NG": "+234", "KE": "+254", "MA": "+212",
    "DZ": "+213", "TN": "+216", "GH": "+233", "ET": "+251", "TZ": "+255",
    "UG": "+256", "SN": "+221", "CI": "+225", "CM": "+237",
    "MX": "+52", "BR": "+55", "AR": "+54", "CL": "+56", "CO": "+57",
    "PE": "+51", "VE": "+58", "EC": "+593", "UY": "+598", "PY": "+595",
    "BO": "+591", "CR": "+506", "PA": "+507", "GT": "+502", "DO": "+1",
}


def format_country_region_code(countries: list[str]) -> str:
    """"+1 / US; +44 / GB; ..." for each country ISO code we have a calling
    code for, in the order given. Countries with no known calling code are
    silently skipped (never invented) — if that empties the result the
    caller's finalize_row treats the whole field as unverified."""
    parts: list[str] = []
    for iso in countries:
        code = CALLING_CODE.get(iso.upper())
        if code:
            parts.append(f"{code} / {iso.upper()}")
    return "; ".join(parts)


# A draft answer that only says this (no actual country names alongside it)
# is a vague summary, not a real list — must be rejected rather than passed
# through as if it were verified geography/presence data. Matches the "not
# just Global/Worldwide/APAC/EMEA — name every country" instruction given to
# the research prompt; this is the enforcement side of that same rule.
_VAGUE_GEO_RE = re.compile(
    r"^\s*("
    r"global|worldwide|international|multinational|"
    r"na|eu|me|af|latam|apac|emea|"
    r"north america|europe|asia|africa|middle east|latin america|"
    r"\d+\+?\s*countries?"
    r")\s*$",
    re.I,
)


def _is_vague_geo_text(raw: str) -> bool:
    """True for a bare region name / "Global" / "50+ countries" with no
    actual country names extractable from it — parse_country_list already
    strips region words on its own, but a string that becomes fully empty
    after that strip (rather than "" to begin with) still needs to be
    treated as vague, not silently left as its original unhelpful text."""
    s = raw.strip()
    if not s:
        return False
    return bool(_VAGUE_GEO_RE.match(s)) or not parse_country_list(s)


def derive_geo_fields(row: dict[str, Any]) -> dict[str, Any]:
    """Fill Continent / Geography, Operational Presence, and Country Code /
    Region Code from whatever real Headquarters/geography text the row
    already has. Mirrors landscape_slim.finalize_slim_row's approach but
    targets this schema's column names/shape.

    A draft value that is vague ("Global", "Worldwide", "50+ countries", a
    bare region name) rather than an actual list of countries is blanked
    out here rather than passed through — finalize_row then rejects the row
    for that field being unverified, per the explicit "name every country,
    don't summarize" requirement."""
    out = dict(row)
    hq = str(out.get("Headquarters") or "").strip()
    if is_placeholder(hq):
        hq = ""

    geo_raw = out.get("Continent / Geography") or ""
    op_raw = out.get("Operational Presence") or ""

    if _is_vague_geo_text(str(geo_raw)):
        geo_raw = ""
        out["Continent / Geography"] = ""
    if _is_vague_geo_text(str(op_raw)):
        op_raw = ""
        out["Operational Presence"] = ""

    countries = parse_country_list(geo_raw)
    if not countries and hq:
        countries = parse_country_list(hq)
    if countries:
        out["Continent / Geography"] = format_continent_geography(countries)

    if hq or str(op_raw).strip() or str(geo_raw).strip():
        derived_op = operational_presence_text(hq, op_raw, geo_raw)
        if derived_op and not is_placeholder(derived_op):
            out["Operational Presence"] = derived_op

    isos: list[str] = []
    for country in countries or ([hq] if hq else []):
        iso = hq_to_iso(country)
        if iso and iso not in isos:
            isos.append(iso)
    if not isos and hq:
        iso = hq_to_iso(hq)
        if iso:
            isos.append(iso)

    if isos:
        combined = format_country_region_code(isos)
        if combined:
            out["Country Code / Region Code"] = combined

    return out


def finalize_row(candidate: dict[str, Any], *, index: int,
                 market: str = "") -> dict[str, str] | None:
    """Assemble one final row from a researched+verified candidate dict.

    Returns None (company excluded from the batch) if any of the 21 required
    fields is missing or still a placeholder after geo derivation — per spec
    section 8, accuracy takes priority over reaching a full batch of 20.
    """
    row = derive_geo_fields(candidate)
    out: dict[str, str] = {"#": str(index)}
    # Market-aware: a software market has no "Online / Offline / Both"
    # column, so a blank one must not drop the row. Defaults to the full
    # list when no market is passed, keeping existing callers unchanged.
    required = required_fields_for(market) if market else REQUIRED_FIELDS
    for col in required:
        value = row.get(col)
        if is_placeholder(value):
            return None
        out[col] = str(value).strip()

    # Contact fields are optional and must never carry a placeholder
    # string: an unverified value is normalised to "" so a blank cell
    # plainly means "not verified" rather than "N/A"/"unknown"/invented.
    for col in CONTACT_FIELDS:
        value = row.get(col)
        out[col] = "" if is_placeholder(value) else str(value).strip()
    return out
