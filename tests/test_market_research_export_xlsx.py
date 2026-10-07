"""export_xlsx.py: final dedup + completeness re-check before writing."""
from __future__ import annotations

from openpyxl import load_workbook

from vendor_intel.pipeline.market_research import export_xlsx, schema
from vendor_intel.pipeline.market_research.state import MarketResearchState


def _header_row(ws) -> int:
    """Row index of the real column header.

    The sheet now opens with a market/definition banner (requested so the
    inclusion rule is visible in the file itself), so the header is no
    longer row 1 and tests must locate it rather than assume.
    """
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "#" and ws.cell(r, 2).value == "Company":
            return r
    raise AssertionError("column header row not found")


def _headers(ws) -> list:
    hr = _header_row(ws)
    return [c.value for c in ws[hr]]


def _data_rows(ws) -> list:
    hr = _header_row(ws)
    return [
        [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        for r in range(hr + 1, ws.max_row + 1)
        if ws.cell(r, 2).value
    ]


def _complete_row(company: str, **overrides) -> dict:
    # Contact fields are no longer in REQUIRED_FIELDS but still belong in a
    # realistic row, so populate both sets here.
    row = {
        col: f"{company}-{col}"
        for col in (*schema.REQUIRED_FIELDS, *schema.CONTACT_FIELDS)
    }
    row["Company"] = company
    row["Website"] = f"https://{company.lower().replace(' ', '')}.example.com"
    row.update(overrides)
    return row


def test_blank_contact_fields_do_not_drop_a_company(tmp_path, monkeypatch):
    """The six contact fields are optional now (real data or nothing), so
    a company with none of them must still export. If this regressed, the
    whole dataset would vanish the moment fabrication was switched off."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Blank Contacts Market")
    row = _complete_row("Acme Inc")
    for col in schema.CONTACT_FIELDS:
        row[col] = ""
    state.add_batch([row])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Blank Contacts Market"))["Companies"]
    hdr = _headers(ws)
    assert len(_data_rows(ws)) == 1, "the company must survive with blank contacts"
    assert _data_rows(ws)[0][hdr.index("Company")] == "Acme Inc"
    assert _data_rows(ws)[0][hdr.index("Email")] in (None, "")


def test_missing_firmographic_field_still_drops_the_company(tmp_path, monkeypatch):
    """Relaxing the contact fields must not relax the rest of the schema."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Firmographic Market")
    bad = _complete_row("No HQ Co")
    bad["Employees"] = ""
    state.add_batch([_complete_row("Good Co"), bad])
    state.save()

    wb = load_workbook(export_xlsx.export_xlsx("Firmographic Market"))
    ws = wb["Companies"]
    hdr = _headers(ws)
    names = [r[hdr.index("Company")] for r in _data_rows(ws)]
    assert names == ["Good Co"]
    assert "Export Audit" in wb.sheetnames


def test_rows_are_single_line_height(tmp_path, monkeypatch):
    """Rows must NOT be pre-expanded. Every row was previously forced to
    108px (~6 lines) with wrap_text on, so 493 rows were unscrollable and
    each cell looked "expanded". Requested explicitly: proper format, no
    expanded cells."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Height Market")
    state.add_batch([_complete_row("Acme Inc"), _complete_row("Beta Corp")])
    state.save()

    path = export_xlsx.export_xlsx("Height Market")
    ws = load_workbook(path)["Companies"]

    for r in range(_header_row(ws) + 1, ws.max_row + 1):
        height = ws.row_dimensions[r].height
        assert height is None or height <= 20, f"row {r} is {height}px tall"


def test_cells_do_not_wrap(tmp_path, monkeypatch):
    """wrap_text is what makes Excel render a tall multi-line cell."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Wrap Market")
    state.add_batch([_complete_row("Acme Inc")])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Wrap Market"))["Companies"]

    # The banner deliberately wraps (it is prose); the TABLE must not.
    for row in ws.iter_rows(min_row=_header_row(ws), max_row=ws.max_row):
        for cell in row:
            assert not cell.alignment.wrap_text, f"{cell.coordinate} wraps"


def test_embedded_newlines_are_collapsed(tmp_path, monkeypatch):
    """A stray newline inside a value makes Excel show that row multi-line
    even with wrapping off, which reintroduces the expanded look. AI Mode
    answers do sometimes carry them, so they must be normalised."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Newline Market")
    row = _complete_row("Acme Inc")
    row["Operational Presence"] = "NA(United States,\nCanada);\n EU(Germany)"
    state.add_batch([row])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Newline Market"))["Companies"]
    hdr = _headers(ws)
    col = hdr.index("Operational Presence")
    value = _data_rows(ws)[0][col]

    assert "\n" not in value
    assert value == "NA(United States, Canada); EU(Germany)"


def test_index_column_is_numeric(tmp_path, monkeypatch):
    """Written as a number so Excel sorts 2 < 10 rather than "10" < "2"."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Index Market")
    state.add_batch([_complete_row(f"Co {i}") for i in range(3)])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Index Market"))["Companies"]
    assert [r[0] for r in _data_rows(ws)[:3]] == [1, 2, 3]


def test_no_column_is_absurdly_wide(tmp_path, monkeypatch):
    """A column sized to the longest country list would push every other
    field off-screen — the opposite failure from too-narrow columns."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Width Market")
    state.add_batch([_complete_row("Acme Inc")])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Width Market"))["Companies"]
    for letter, dim in ws.column_dimensions.items():
        if dim.width:
            assert dim.width <= 50, f"column {letter} is {dim.width} wide"


def test_dedupe_and_validate_drops_duplicates(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    rows = [
        _complete_row("Acme Inc"),
        _complete_row("Acme Inc"),  # exact duplicate
        _complete_row("Acme"),  # dedupe-key collides with "Acme Inc" (suffix stripped)
        _complete_row("Beta Corp"),
    ]
    clean, dropped = export_xlsx._dedupe_and_validate(rows)
    assert [r["Company"] for r in clean] == ["Acme Inc", "Beta Corp"]
    assert len(dropped) == 2


def test_dedupe_and_validate_drops_incomplete_rows():
    incomplete = _complete_row("Gamma Ltd")
    incomplete["Employees"] = "N/A"
    rows = [_complete_row("Delta Co"), incomplete]
    clean, dropped = export_xlsx._dedupe_and_validate(rows)
    assert [r["Company"] for r in clean] == ["Delta Co"]
    assert any("Gamma Ltd" in d and "Employees" in d for d in dropped)


def test_export_xlsx_writes_clean_rows_and_audit_sheet(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Export Market")
    # A missing FIRMOGRAPHIC field still drops the row. (Deliberately not
    # Email: the six contact fields are legitimately blank now, so a blank
    # Email must NOT drop a company — see test_blank_contact_fields_*.)
    incomplete = _complete_row("Broken Co")
    incomplete["Headquarters"] = ""
    state.add_batch([_complete_row("Good Co", **{"#": "1"}), incomplete])
    state.save()

    out_path = tmp_path / "out.xlsx"
    result_path = export_xlsx.export_xlsx("Test Export Market", path=out_path)
    assert result_path == out_path
    assert out_path.exists()

    wb = load_workbook(out_path)
    ws = wb["Companies"]
    header = _headers(ws)
    assert header == list(schema.COLUMNS)
    company_col = header.index("Company")
    companies = [r[company_col] for r in _data_rows(ws)]
    assert companies == ["Good Co"]

    assert "Export Audit" in wb.sheetnames
    audit_rows = [r[0].value for r in wb["Export Audit"].iter_rows(min_row=2)]
    assert any("Broken Co" in (r or "") for r in audit_rows)


def test_export_xlsx_no_audit_sheet_when_nothing_dropped(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Clean Market")
    state.add_batch([_complete_row("Only Co")])
    state.save()

    out_path = tmp_path / "clean.xlsx"
    export_xlsx.export_xlsx("Clean Market", path=out_path)
    wb = load_workbook(out_path)
    assert "Export Audit" not in wb.sheetnames


# --- definition banner ----------------------------------------------------
#
# Requested because "distributor" does not mean the same thing in every
# market: for CPaaS the user stated that "channel partners ARE known as
# distributors over here", while in a physical-goods market they are
# explicitly excluded. Without the rule written in the file, the inclusion
# criteria are invisible and the sheet cannot be audited.


def _banner_text(ws) -> str:
    hr = _header_row(ws)
    return " ".join(str(ws.cell(r, 1).value or "") for r in range(1, hr))


def test_banner_names_the_market(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Global Polyurethane Market - Distributors Only")
    state.add_batch([_complete_row("Acme Inc")])
    state.save()

    ws = load_workbook(
        export_xlsx.export_xlsx("Global Polyurethane Market - Distributors Only")
    )["Companies"]
    assert "Polyurethane" in _banner_text(ws)


def test_physical_market_banner_excludes_channel_partners(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Global Silicon Carbide Market - Distributors Only")
    state.add_batch([_complete_row("Acme Inc")])
    state.save()

    ws = load_workbook(
        export_xlsx.export_xlsx("Global Silicon Carbide Market - Distributors Only")
    )["Companies"]
    text = _banner_text(ws)
    assert "physical goods" in text.lower()
    assert "EXCLUDED" in text and "channel partners" in text.lower()


def test_cpaas_banner_says_channel_partners_ARE_distributors(tmp_path, monkeypatch):
    """The opposite rule, per explicit instruction for this market."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    m = "Global CPaaS (Communication Platform as a Service) Market - Distributors Only"
    state = MarketResearchState.load(m)
    state.add_batch([_complete_row("Acme Inc")])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx(m))["Companies"]
    text = _banner_text(ws)
    assert "CHANNEL PARTNERS ARE THE DISTRIBUTORS" in text
    assert "software" in text.lower() or "cloud" in text.lower()


def test_banner_explains_the_two_geography_columns(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Banner Geo Market")
    state.add_batch([_complete_row("Acme Inc")])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Banner Geo Market"))["Companies"]
    text = _banner_text(ws)
    assert "physical site" in text
    assert "WIDER" in text
    assert "NPD" in text


def test_banner_does_not_break_the_table(tmp_path, monkeypatch):
    """The header must still be a real header: filterable and frozen."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Banner Table Market")
    state.add_batch([_complete_row(f"Co {i}") for i in range(3)])
    state.save()

    ws = load_workbook(export_xlsx.export_xlsx("Banner Table Market"))["Companies"]
    hr = _header_row(ws)
    assert hr > 1, "banner should sit above the header"
    assert _headers(ws) == list(schema.COLUMNS)
    assert len(_data_rows(ws)) == 3
    # the filter must start at the header, not at the banner
    assert ws.auto_filter.ref.startswith(f"A{hr}:")
    assert ws.freeze_panes == f"B{hr + 1}"


# --- definition first, companies second ----------------------------------
#
# Requested ordering: define the market, write it into the Excel, and only
# then search for companies and fill them in. Previously the file appeared
# only after a run finished, so during a multi-hour run there was nothing
# to inspect and no record of which inclusion rule was in force.


def test_definition_is_written_before_any_company_exists(tmp_path, monkeypatch):
    """Exporting an empty market must still produce the header sheet."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)

    path = export_xlsx.export_xlsx("Global Empty Market - Distributors Only")
    ws = load_workbook(path)["Companies"]

    hr = _header_row(ws)
    assert hr > 1, "the definition banner must sit above the header"
    assert "DISTRIBUTOR DEFINITION" in str(ws.cell(2, 1).value)
    assert _data_rows(ws) == [], "no companies yet"


def test_companies_are_appended_under_the_same_definition(tmp_path, monkeypatch):
    """Re-exporting after a batch must keep the banner and add rows below,
    so the file is usable while the run is still going."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    market = "Global Growing Market - Distributors Only"

    export_xlsx.export_xlsx(market)          # header only

    state = MarketResearchState.load(market)
    state.add_batch([_complete_row("Acme Inc")])
    state.save()
    ws = load_workbook(export_xlsx.export_xlsx(market))["Companies"]

    hr = _header_row(ws)
    assert "DISTRIBUTOR DEFINITION" in str(ws.cell(2, 1).value)
    assert len(_data_rows(ws)) == 1
    assert ws.cell(hr + 1, 2).value == "Acme Inc"


def test_header_row_does_not_drift_as_companies_are_added(tmp_path, monkeypatch):
    """The banner is a fixed height, so the header must land on the same
    row every refresh - otherwise the autofilter and freeze-panes would
    move under the user mid-run."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    market = "Global Stable Header Market - Distributors Only"

    ws0 = load_workbook(export_xlsx.export_xlsx(market))["Companies"]
    first = _header_row(ws0)

    state = MarketResearchState.load(market)
    for i in range(3):
        state.add_batch([_complete_row(f"Co {i}")])
        state.save()
        ws = load_workbook(export_xlsx.export_xlsx(market))["Companies"]
        assert _header_row(ws) == first, "header row moved between refreshes"
