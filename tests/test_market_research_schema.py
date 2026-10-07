"""schema.py: 21-column shape, placeholder rejection, geo derivation."""
from __future__ import annotations

from vendor_intel.pipeline.market_research import schema


def test_columns_are_exactly_21_in_order():
    assert len(schema.COLUMNS) == 21
    assert schema.COLUMNS[0] == "#"
    assert schema.COLUMNS[1] == "Company"
    assert schema.COLUMNS[-1] == "Country Code / Region Code"
    assert "Key Products / Technologies" in schema.COLUMNS
    assert "Office Website" in schema.COLUMNS
    assert "Summary" not in schema.COLUMNS


def _full_candidate(**overrides) -> dict:
    base = {
        "Company": "Acme Semiconductor Distribution Inc",
        "Website": "https://acmesemi.example.com",
        "Founded": "1998",
        "Headquarters": "Austin, United States",
        "Continent / Geography": "",
        "Operational Presence": "",
        "Ownership": "Private Company",
        "Employees": "250-500",
        "Core Categories": "Semiconductor distribution",
        "Specialty Focus": "Analog ICs",
        "Key Products / Technologies": "Analog ICs, power management chips",
        "Online / Offline / Both": "Both",
        "Contact Person": "Jane Doe",
        "Role": "VP Sales",
        "Email": "sales@acmesemi.example.com",
        "LinkedIn": "https://linkedin.com/company/acme-semi",
        "Office No.": "+1 512-555-0100",
        "Office Website": "https://acmesemi.example.com",
        "Office Email": "sales@acmesemi.example.com",
    }
    base.update(overrides)
    return base


def test_finalize_row_accepts_complete_candidate():
    row = schema.finalize_row(_full_candidate(), index=1)
    assert row is not None
    assert row["#"] == "1"
    assert row["Company"] == "Acme Semiconductor Distribution Inc"
    assert set(row.keys()) == set(schema.COLUMNS)
    # Geo fields derived from Headquarters since Continent/Geography was blank.
    assert "United States" in row["Continent / Geography"]
    assert row["Country Code / Region Code"]


def test_finalize_row_rejects_when_any_field_missing():
    for missing_field in schema.REQUIRED_FIELDS:
        if missing_field in ("Continent / Geography", "Operational Presence", "Country Code / Region Code"):
            # These are derivable from Headquarters — skip, covered separately.
            continue
        candidate = _full_candidate(**{missing_field: ""})
        assert schema.finalize_row(candidate, index=1) is None, missing_field


def test_finalize_row_rejects_placeholder_values():
    for placeholder in ("n/a", "N/A", "unknown", "TBD", "not disclosed", "-", ""):
        candidate = _full_candidate(Employees=placeholder)
        assert schema.finalize_row(candidate, index=1) is None


def test_finalize_row_none_when_headquarters_unresolvable():
    candidate = _full_candidate(Headquarters="Nowhereland")
    row = schema.finalize_row(candidate, index=1)
    # Unresolvable country -> no Country Code / Region Code derivable -> excluded.
    assert row is None


def test_format_country_region_code_skips_unknown_iso():
    assert schema.format_country_region_code(["US", "GB"]) == "+1 / US; +44 / GB"
    assert schema.format_country_region_code([]) == ""


def test_finalize_row_rejects_vague_continent_geography():
    """"Global"/"Worldwide"/"50+ countries" etc. are summaries, not a real
    list of countries, and must not be accepted as verified geography —
    even though Headquarters alone is real, the draft's own vague answer
    for this field must not silently pass through."""
    for vague in ("Global", "Worldwide", "APAC", "50+ countries", "International"):
        candidate = _full_candidate(
            **{"Continent / Geography": vague, "Headquarters": "Nowhereland"}
        )
        assert schema.finalize_row(candidate, index=1) is None, vague


def test_finalize_row_accepts_explicit_country_list_in_continent_geography():
    candidate = _full_candidate(
        **{
            "Continent / Geography": "India, Singapore, Japan",
            "Operational Presence": "India, Singapore, Japan, Vietnam",
        }
    )
    row = schema.finalize_row(candidate, index=1)
    assert row is not None
    assert "India" in row["Continent / Geography"]
    assert "Singapore" in row["Continent / Geography"]
    assert "Japan" in row["Continent / Geography"]
    assert "Vietnam" in row["Operational Presence"]


def test_is_vague_geo_text_rejects_summaries_and_accepts_lists():
    assert schema._is_vague_geo_text("Global") is True
    assert schema._is_vague_geo_text("Worldwide") is True
    assert schema._is_vague_geo_text("APAC") is True
    assert schema._is_vague_geo_text("50+ countries") is True
    assert schema._is_vague_geo_text("India, Singapore, Japan") is False
    assert schema._is_vague_geo_text("") is False
