"""Software/cloud markets omit the "Online / Offline / Both" column.

"Online / Offline / Both" describes a PHYSICAL sales channel: stock
shipped from a warehouse, sold over a counter, or both. A CPaaS or SaaS
reseller delivers its product over the internet by definition, so the
column can only ever say "Online" -- a constant that carries no
information -- or be invented to fill a cell the model was told was
mandatory.

Three things have to agree for the column to really be gone, and each is
a separate failure if it drifts:

  1. the exported workbook must not have the column;
  2. the required-field gate must not drop a row for leaving it blank
     (otherwise removing it from the sheet silently destroys companies);
  3. the discovery prompt must stop asking for it.
"""
from __future__ import annotations

from vendor_intel.pipeline.market_research import discovery, schema

SOFTWARE = "Global CPaaS (Communication Platform as a Service) Market - Distributors Only"
PHYSICAL = "Global Polyurethane Market - Distributors Only"
CHANNEL = "Online / Offline / Both"


# --- which markets count as software -------------------------------------


def test_cpaas_is_a_software_market():
    assert schema.is_software_market(SOFTWARE)


def test_a_chemicals_market_is_not():
    assert not schema.is_software_market(PHYSICAL)


def test_the_hint_list_is_shared_with_the_distributor_definition():
    """The same markets that get the software DISTRIBUTOR rule must get
    the software COLUMN rule -- two separate lists would drift."""
    for market in ("Global SaaS Market", "Global Cloud Platform Market",
                   "Global API Management Market"):
        assert schema.is_software_market(market)


# --- 1. the column list ---------------------------------------------------


def test_software_market_columns_exclude_the_channel():
    assert CHANNEL not in schema.columns_for(SOFTWARE)


def test_physical_market_columns_keep_the_channel():
    assert CHANNEL in schema.columns_for(PHYSICAL)


def test_only_the_channel_column_is_removed():
    """Removing one column must not disturb any other."""
    removed = set(schema.columns_for(PHYSICAL)) - set(schema.columns_for(SOFTWARE))
    assert removed == {CHANNEL}


def test_column_order_is_otherwise_unchanged():
    expected = [c for c in schema.COLUMNS if c != CHANNEL]
    assert list(schema.columns_for(SOFTWARE)) == expected


# --- 2. the required-field gate -------------------------------------------


def test_the_channel_is_not_required_in_a_software_market():
    """The point of the whole change: a blank channel must not drop a row
    in a market where the column does not exist."""
    assert CHANNEL not in schema.required_fields_for(SOFTWARE)


def test_the_channel_is_still_required_in_a_physical_market():
    assert CHANNEL in schema.required_fields_for(PHYSICAL)


def test_contact_fields_stay_optional_in_both():
    for market in (SOFTWARE, PHYSICAL):
        required = schema.required_fields_for(market)
        for field in schema.CONTACT_FIELDS:
            assert field not in required


def _candidate(**over):
    """A complete row in finalize_row's own key space.

    Note the DISPLAY names: finalize_row consumes the mapped row that
    batch_runner._build_row produces, not the snake_case draft straight
    from the model.
    """
    row = {
        "Company": "Example Reseller Ltd", "Website": "example.com",
        "Founded": "2010", "Headquarters": "Lisbon, Portugal",
        "Continent / Geography": "Europe(Portugal)",
        "Operational Presence": "Portugal, Spain",
        "Ownership": "Private Company", "Employees": "50-100",
        "Core Categories": "Cloud communications",
        "Specialty Focus": "SMS APIs",
        "Key Products / Technologies": "Twilio, Vonage",
        "Office Website": "example.com",
        "Country Code / Region Code": "+351 / PT",
        CHANNEL: "",
    }
    row.update(over)
    return row


def test_a_software_row_survives_without_a_channel_value():
    """The point of the change, end to end: one candidate with a blank
    channel is KEPT for a software market and DROPPED for a physical one,
    proving the gate is what changed and nothing else."""
    kept = schema.finalize_row(_candidate(), index=1, market=SOFTWARE)
    assert kept is not None, "a software row must survive a blank channel"
    assert not kept.get(CHANNEL), "the channel must not be invented"

    dropped = schema.finalize_row(_candidate(), index=1, market=PHYSICAL)
    assert dropped is None, "a physical market still requires the channel"


def test_a_physical_row_with_a_channel_is_kept():
    """Guards against the gate being loosened for everyone."""
    kept = schema.finalize_row(
        _candidate(**{CHANNEL: "Both"}), index=1, market=PHYSICAL)
    assert kept is not None
    assert kept[CHANNEL] == "Both"


# --- 3. the discovery prompt ----------------------------------------------


def test_the_prompt_stops_asking_for_the_channel_in_software_markets():
    prompt = discovery.build_draft_prompt(
        SOFTWARE, excluded=[], count=10, collected=[])
    assert "- online_offline_both:" not in prompt
    assert '"online_offline_both"' not in prompt


def test_the_prompt_still_asks_for_it_in_physical_markets():
    prompt = discovery.build_draft_prompt(
        PHYSICAL, excluded=[], count=10, collected=[])
    assert "- online_offline_both:" in prompt
    assert '"online_offline_both"' in prompt


def test_the_reply_template_stays_valid_after_the_field_is_removed():
    """The JSON template is the last thing in the prompt; a bad edit there
    costs the whole answer."""
    prompt = discovery.build_draft_prompt(
        SOFTWARE, excluded=[], count=10, collected=[])
    tail = prompt[prompt.rindex("[{"):]
    assert tail.count("{") == tail.count("}")
    assert '"company"' in tail and '"verdict"' in tail
