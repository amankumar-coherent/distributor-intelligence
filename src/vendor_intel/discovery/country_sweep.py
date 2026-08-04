"""Country-wise discovery sweep for global markets.

Builds queries like:
  "{market} distributors in India"
  "{market} wholesalers channel partners Germany"

Enabled when DISCOVERY_COUNTRY_SWEEP=true (default true for global runs via caller).
Stops early once unique company count reaches the configured minimum (default 300).
"""
from __future__ import annotations

import os
from typing import Any

# Priority order: major commercial markets first, then remaining ISO-style countries.
# Names match how search engines expect them in natural-language queries.
PRIORITY_COUNTRIES: list[str] = [
    "United States",
    "India",
    "China",
    "United Kingdom",
    "Germany",
    "Japan",
    "France",
    "Canada",
    "Australia",
    "Brazil",
    "South Korea",
    "Italy",
    "Spain",
    "Netherlands",
    "Mexico",
    "Indonesia",
    "Turkey",
    "Saudi Arabia",
    "United Arab Emirates",
    "Singapore",
    "Switzerland",
    "Sweden",
    "Poland",
    "Belgium",
    "Thailand",
    "Vietnam",
    "Malaysia",
    "Philippines",
    "South Africa",
    "Nigeria",
    "Egypt",
    "Israel",
    "Argentina",
    "Chile",
    "Colombia",
    "Peru",
    "Austria",
    "Denmark",
    "Finland",
    "Norway",
    "Ireland",
    "Portugal",
    "Czech Republic",
    "Greece",
    "Hungary",
    "Romania",
    "New Zealand",
    "Taiwan",
    "Hong Kong",
    "Pakistan",
    "Bangladesh",
    "Ukraine",
    "Russia",
    "Kenya",
    "Ghana",
    "Morocco",
    "Qatar",
    "Kuwait",
    "Oman",
    "Bahrain",
    "Jordan",
    "Lebanon",
]

# Remaining countries (alphabetical) to approach ~195 total
_REST: list[str] = [
    "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda",
    "Armenia", "Azerbaijan", "Bahamas", "Barbados", "Belarus", "Belize", "Benin",
    "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brunei", "Bulgaria",
    "Burkina Faso", "Burundi", "Cambodia", "Cameroon", "Cape Verde",
    "Central African Republic", "Chad", "Comoros", "Congo", "Costa Rica",
    "Croatia", "Cuba", "Cyprus", "Djibouti", "Dominica", "Dominican Republic",
    "Ecuador", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia",
    "Eswatini", "Ethiopia", "Fiji", "Gabon", "Gambia", "Georgia", "Grenada",
    "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Honduras",
    "Iceland", "Iraq", "Jamaica", "Kazakhstan", "Kiribati", "Kyrgyzstan", "Laos",
    "Latvia", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania",
    "Luxembourg", "Madagascar", "Malawi", "Maldives", "Mali", "Malta",
    "Marshall Islands", "Mauritania", "Mauritius", "Micronesia", "Moldova",
    "Monaco", "Mongolia", "Montenegro", "Mozambique", "Myanmar", "Namibia",
    "Nauru", "Nepal", "Nicaragua", "Niger", "North Macedonia", "Palau", "Panama",
    "Papua New Guinea", "Paraguay", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia",
    "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe",
    "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Slovakia", "Slovenia",
    "Solomon Islands", "Somalia", "South Sudan", "Sri Lanka", "Sudan", "Suriname",
    "Syria", "Tajikistan", "Tanzania", "Timor-Leste", "Togo", "Tonga",
    "Trinidad and Tobago", "Tunisia", "Turkmenistan", "Tuvalu", "Uganda", "Uruguay",
    "Uzbekistan", "Vanuatu", "Venezuela", "Yemen", "Zambia", "Zimbabwe",
]


def _env_true(name: str, default: bool = False) -> bool:
    raw = (os.getenv(name) or "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "on")


def country_sweep_enabled(*, default: bool = True) -> bool:
    """Master switch. Default ON so global runs sweep countries unless disabled."""
    return _env_true("DISCOVERY_COUNTRY_SWEEP", default)


def country_sweep_min_companies() -> int:
    try:
        return max(50, int(os.getenv("DISCOVERY_COUNTRY_SWEEP_MIN_COMPANIES", "300") or "300"))
    except ValueError:
        return 300


def country_sweep_max_countries() -> int:
    try:
        return max(1, int(os.getenv("DISCOVERY_COUNTRY_SWEEP_MAX_COUNTRIES", "195") or "195"))
    except ValueError:
        return 195


def country_sweep_queries_per_country() -> int:
    """How many query templates per country (1–6)."""
    try:
        return max(1, min(6, int(os.getenv("DISCOVERY_COUNTRY_SWEEP_QUERIES_PER_COUNTRY", "2") or "2")))
    except ValueError:
        return 2


def world_countries(*, max_countries: int | None = None) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for name in PRIORITY_COUNTRIES + _REST:
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(name)
    cap = max_countries if max_countries is not None else country_sweep_max_countries()
    return out[:cap]


def _market_label(market: str) -> str:
    m = (market or "").strip()
    # Drop trailing "Market" for cleaner search queries
    if m.lower().endswith(" market"):
        m = m[: -len(" market")].strip()
    return m or "market"


def build_country_distributor_prompts(
    market: str,
    country: str,
    *,
    queries_per_country: int | None = None,
    start_index: int = 0,
) -> list[dict[str, str]]:
    """Build distributor/channel-partner prompts for one country."""
    m = _market_label(market)
    c = (country or "").strip()
    if not m or not c:
        return []
    n = queries_per_country if queries_per_country is not None else country_sweep_queries_per_country()
    templates = [
        f"{m} distributors in {c}",
        f"{m} wholesalers channel partners {c}",
        f"{m} authorized distributors {c}",
        f"{m} importers distributors {c}",
        f"{m} bulk oil distributors wholesalers {c}",
        f"{m} reseller distributor supplier {c}",
    ]
    prompts: list[dict[str, str]] = []
    for i, text in enumerate(templates[:n]):
        prompts.append(
            {
                "id": f"CS{start_index + i + 1}",
                "level": "country_sweep",
                "text": text,
                "sub_sector": "distributor",
                "country": c,
            }
        )
    return prompts


def build_country_sweep_prompt_batches(
    market: str,
    *,
    max_countries: int | None = None,
    queries_per_country: int | None = None,
) -> list[tuple[str, list[dict[str, str]]]]:
    """Return [(country, prompts), ...] in priority order."""
    batches: list[tuple[str, list[dict[str, str]]]] = []
    idx = 0
    for country in world_countries(max_countries=max_countries):
        prompts = build_country_distributor_prompts(
            market,
            country,
            queries_per_country=queries_per_country,
            start_index=idx,
        )
        if prompts:
            batches.append((country, prompts))
            idx += len(prompts)
    return batches


def should_run_country_sweep(scope: dict[str, Any] | None) -> bool:
    if not country_sweep_enabled(default=True):
        return False
    from vendor_intel.pipeline.geo_limits import is_global_geography

    geo = ""
    if isinstance(scope, dict):
        geos = scope.get("geographies") or ["global"]
        geo = str(geos[0] if geos else "global")
    return is_global_geography(geo)
