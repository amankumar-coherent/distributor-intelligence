"""The VERIFIED workbook is the one handed on, so it needs the banner too.

Found by opening the finished Dry Eyes verified file: 373 companies,
header on row 1, no definition banner, and wrap_text on every cell.

Both are inconsistencies with the discovery export, and both matter:

  * The banner states which rule selected these rows. "Distributor" means
    different things in different markets -- in CPaaS the channel partners
    ARE the distributors, in a physical-goods market they are explicitly
    excluded -- so a sheet without it cannot be audited. It also carries
    the product scope, which is what separates a dry-eye DRUG distributor
    from an ophthalmic EQUIPMENT distributor.
  * wrap_text is the "expanded cell" rendering already removed from the
    discovery export: with the per-country geography columns populated it
    makes every row ~6 lines tall, and 373 rows of that cannot be
    scrolled or compared.

The banner text is derived from export_xlsx._definition_banner -- the same
functions that built the discovery prompt -- so the two workbooks cannot
state different criteria.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import openpyxl
import pytest

_ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location(
    "verify_distributors_ai_mode", _ROOT / "verify_distributors_ai_mode.py"
)
vd = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(vd)

_HEADERS = ["#", "Company", "Website", "Continent / Geography",
            "Operational Presence"]
_ROWS = [
    {
        "#": i,
        "Company": f"Distributor {i}",
        "Website": f"https://d{i}.example",
        "Continent / Geography": "APAC(India, Singapore); EU(Germany)",
        "Operational Presence": "India; Singapore; Germany; Kenya",
    }
    for i in range(1, 4)
]


def _write(tmp_path: Path, market_name: str = "Global Dry Eyes Treatment Market"):
    out = tmp_path / "verified.xlsx"
    vd.write_output(_ROWS, _HEADERS, "Companies", out, market_name=market_name)
    return openpyxl.load_workbook(out).active


def _header_row(ws) -> int:
    for r in range(1, 26):
        if any(str(c.value or "") == "Company" for c in ws[r]):
            return r
    raise AssertionError("no header row found")


# --- the banner -----------------------------------------------------------


def test_the_market_name_is_the_first_line(tmp_path):
    ws = _write(tmp_path)
    assert "Dry Eyes" in str(ws.cell(1, 1).value)


def test_the_distributor_definition_is_stated(tmp_path):
    ws = _write(tmp_path)
    text = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, 6))
    assert "DISTRIBUTOR DEFINITION" in text


def test_the_product_scope_is_stated(tmp_path):
    """This is what distinguishes a dry-eye DRUG distributor from an
    ophthalmic EQUIPMENT distributor -- the whole reason the scope exists.
    Without it in the sheet the inclusion criteria are invisible."""
    ws = _write(tmp_path)
    text = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, 6))
    assert "PRODUCT SCOPE" in text
    assert '"dry eye" AND (' in text


def test_the_geography_columns_are_explained(tmp_path):
    """The two geography columns mean different things and the difference
    is not guessable from the headers."""
    ws = _write(tmp_path)
    text = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, 6))
    assert "physical site" in text
    assert "WIDER" in text


def test_the_software_market_gets_the_inverted_definition(tmp_path):
    """CPaaS is the case that proves the banner is derived rather than
    hardcoded: there the channel partners ARE the distributors."""
    ws = _write(tmp_path, "Global CPaaS (Communication Platform as a Service) Market")
    text = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, 6))
    assert "software" in text.lower() or "CHANNEL PARTNERS ARE THE DISTRIBUTORS" in text


def test_the_banner_matches_the_discovery_export(tmp_path):
    """Derived from the same function, so the two workbooks cannot state
    different criteria for the same market."""
    from vendor_intel.pipeline.market_research.export_xlsx import _definition_banner

    market = "Global Polyurethane Market - Distributors Only"
    ws = _write(tmp_path, market)
    sheet_text = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, 8))
    for text, _is_title in _definition_banner(market):
        head = " ".join(str(text).split())[:60]
        assert head in sheet_text, head


# --- the table itself -----------------------------------------------------


def test_every_kept_row_is_written(tmp_path):
    """Formatting must never cost data."""
    ws = _write(tmp_path)
    hdr = _header_row(ws)
    assert ws.max_row - hdr == len(_ROWS)


def test_long_columns_wrap_and_short_ones_do_not(tmp_path):
    """Every column must be readable WITHOUT expanding the cell.

    Two mechanisms, chosen per column. Short factual values (Founded, #)
    get a width that fits them and stay single-line. Long free text
    (geography lists reach 423 characters, office contacts 223 now that
    they hold every published number) would need a 400-wide column to fit
    on one line, which pushes every other field off the screen -- so those
    wrap into a taller row instead.

    This replaced a blanket "nothing wraps" assertion, which was correct
    when every column was single-line but became wrong once the columns
    had per-column widths.
    """
    ws = _write(tmp_path)
    hdr = _header_row(ws)
    from openpyxl.utils import get_column_letter

    for c, header in enumerate(_HEADERS, 1):
        width = ws.column_dimensions[get_column_letter(c)].width or 8.43
        wraps = bool(ws.cell(hdr + 1, c).alignment.wrap_text)
        longest = max(
            (len(str(ws.cell(r, c).value or ""))
             for r in range(hdr + 1, ws.max_row + 1)),
            default=0,
        )
        assert wraps or longest <= width, (
            f"{header}: {longest} chars in a {width:.0f}-wide column that "
            "does not wrap — the value is cut off"
        )


def test_the_autofilter_anchors_on_the_real_header(tmp_path):
    """Anchoring at A1 would treat the banner as the header and break
    sorting -- the same off-by-one that broke the discovery export."""
    ws = _write(tmp_path)
    hdr = _header_row(ws)
    assert ws.auto_filter.ref is not None
    assert ws.auto_filter.ref.startswith(f"A{hdr}:")


def test_the_panes_freeze_below_the_header(tmp_path):
    ws = _write(tmp_path)
    hdr = _header_row(ws)
    assert ws.freeze_panes == f"B{hdr + 1}"


def test_embedded_newlines_are_collapsed(tmp_path):
    """One stray newline in a value makes Excel render that row
    multi-line even with wrap_text off."""
    out = tmp_path / "nl.xlsx"
    vd.write_output(
        [{"#": 1, "Company": "Acme\nDistribution", "Website": "https://a.example",
          "Continent / Geography": "APAC(India)\nEU(Germany)",
          "Operational Presence": ""}],
        _HEADERS, "Companies", out, market_name="Global Polyurethane Market",
    )
    ws = openpyxl.load_workbook(out).active
    hdr = _header_row(ws)
    for c in range(1, len(_HEADERS) + 1):
        assert "\n" not in str(ws.cell(hdr + 1, c).value or "")


# --- robustness -----------------------------------------------------------


def test_no_market_name_still_writes_the_data(tmp_path):
    """Callers that pass no market name must still get their rows."""
    out = tmp_path / "nobanner.xlsx"
    vd.write_output(_ROWS, _HEADERS, "Companies", out)
    ws = openpyxl.load_workbook(out).active
    assert _header_row(ws) == 1
    assert ws.max_row == len(_ROWS) + 1


def test_an_unknown_market_still_writes_the_data(tmp_path):
    """A market with no product scope must not lose its banner or rows."""
    ws = _write(tmp_path, "Global Widget Market")
    hdr = _header_row(ws)
    assert ws.max_row - hdr == len(_ROWS)


def test_a_banner_failure_never_costs_the_data(tmp_path, monkeypatch):
    """The banner is presentation; the verified companies are the point."""
    import vendor_intel.pipeline.market_research.export_xlsx as ex

    def _boom(_market):
        raise RuntimeError("banner exploded")

    monkeypatch.setattr(ex, "_definition_banner", _boom)
    out = tmp_path / "safe.xlsx"
    vd.write_output(_ROWS, _HEADERS, "Companies", out,
                    market_name="Global Dry Eyes Treatment Market")
    ws = openpyxl.load_workbook(out).active
    hdr = _header_row(ws)
    assert ws.max_row - hdr == len(_ROWS)


def test_empty_rows_do_not_raise(tmp_path):
    """A market where everything was dropped still produces a readable
    file rather than a crash."""
    out = tmp_path / "empty.xlsx"
    vd.write_output([], _HEADERS, "Companies", out,
                    market_name="Global Dry Eyes Treatment Market")
    ws = openpyxl.load_workbook(out).active
    assert _header_row(ws) >= 1


def test_the_verifier_passes_the_market_name_through():
    """Otherwise every verified sheet silently loses its banner."""
    src = (_ROOT / "verify_distributors_ai_mode.py").read_text(encoding="utf-8")
    assert 'market_name=market.get("market_name", "")' in src
