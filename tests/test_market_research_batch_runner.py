"""batch_runner.py: combined discover+draft flow, oversampling rounds, and
the run_until_target loop."""
from __future__ import annotations

from unittest.mock import patch

import pytest

from vendor_intel.pipeline.market_research import batch_runner
from vendor_intel.pipeline.market_research.state import MarketResearchState


def _draft(name: str, **overrides) -> dict:
    base = {
        "company": name,
        "website": f"https://{name.lower().replace(' ', '')}.example.com",
        "founded": "1998",
        "headquarters": "Austin, United States",
        "continent_geography": "",
        "operational_presence": "",
        "ownership": "Private Company",
        "employees": "250-500",
        "core_categories": "Semiconductor distribution",
        "specialty_focus": "Analog ICs",
        "key_products_technologies": "Analog ICs",
        "online_offline_both": "Both",
        "contact_person": "",
        "role": "",
        "linkedin": "",
        "verdict": "distributor",
        "verdict_reason": "Resells third-party semiconductor components.",
    }
    base.update(overrides)
    return base


def _mock_build_row_pass_through(item, *, market=""):
    """Stands in for _build_row, which now also takes the market so it can
    blank a category that merely echoes the market name."""
    return {
        "Company": item["company"],
        "Website": item["website"],
        "Founded": item["founded"],
        "Headquarters": item["headquarters"],
        "Continent / Geography": "APAC(India)",
        "Operational Presence": "India",
        "Ownership": item["ownership"],
        "Employees": item["employees"],
        "Core Categories": item["core_categories"],
        "Specialty Focus": item["specialty_focus"],
        "Key Products / Technologies": item["key_products_technologies"],
        "Online / Offline / Both": item["online_offline_both"],
        "Contact Person": "Jane Doe",
        "Role": "VP Sales",
        "Email": "sales@example.com",
        "LinkedIn": "https://linkedin.com/company/example",
        "Office No.": "+1 512-555-0100",
        "Office Website": item["website"],
        "Office Email": "sales@example.com",
    }


def test_run_one_batch_caps_request_size_per_round(tmp_path, monkeypatch):
    """A single discover_and_draft call for a full 20-company response
    measured at ~267s against the real proxy (over the 120s default
    timeout) — each round's request must be capped to
    _MAX_DRAFT_REQUEST_SIZE, never the full remaining batch_size, so
    multiple smaller/faster rounds fill a 20-company batch instead of one
    slow, timeout-prone call."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market One")

    def fake_discover(market, *, excluded, count, collected=None):
        assert count <= batch_runner._MAX_DRAFT_REQUEST_SIZE
        return [_draft(f"Company {len(excluded) + i}") for i in range(count)]

    with patch(
        "vendor_intel.pipeline.market_research.discovery.discover_and_draft",
        side_effect=fake_discover,
    ) as mock_discover, patch.object(
        batch_runner, "_build_row", side_effect=_mock_build_row_pass_through
    ):
        rows = batch_runner._run_one_batch(state, "Test Market One", batch_size=20)

    assert len(rows) == 20
    # 20 companies at _MAX_DRAFT_REQUEST_SIZE=10 per round takes 2 rounds.
    assert mock_discover.call_count == 20 // batch_runner._MAX_DRAFT_REQUEST_SIZE


def test_run_one_batch_reasks_when_first_round_short(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Two")

    round1 = [_draft("Company A"), _draft("Company B", verdict="manufacturer")]  # 1 survives
    round2 = [_draft("Company C")]  # tops up to 2

    with patch(
        "vendor_intel.pipeline.market_research.discovery.discover_and_draft",
        side_effect=[round1, round2],
    ) as mock_discover, patch.object(
        batch_runner, "_build_row",
        side_effect=lambda item, **_kw: (
            None if item.get("verdict") != "distributor"
            else _mock_build_row_pass_through(item)
        ),
    ):
        rows = batch_runner._run_one_batch(state, "Test Market Two", batch_size=2)

    assert [r["Company"] for r in rows] == ["Company A", "Company C"]
    assert mock_discover.call_count == 2


def test_run_one_batch_stops_when_no_progress(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Three")

    with patch(
        "vendor_intel.pipeline.market_research.discovery.discover_and_draft",
        return_value=[_draft("Rejected Co", verdict="manufacturer")],
    ) as mock_discover, patch.object(batch_runner, "_build_row", return_value=None):
        rows = batch_runner._run_one_batch(state, "Test Market Three", batch_size=5)

    assert rows == []
    # Two no-progress rounds, not eight: the point of this test is that a
    # dead batch does not burn the whole _MAX_DRAFT_ROUNDS budget.
    #
    # One round of tolerance was added deliberately. A round that yields
    # nothing is equally consistent with a transient backend failure as
    # with an exhausted market, and every accepted company changes the
    # exclusion list, so the next round asks a different question. Live
    # cost of stopping on the first: Dry Eyes batch 10 ended at 6/20 with
    # six of eight rounds unused, dropping yield from +19 to +6.
    assert mock_discover.call_count == 2
    assert mock_discover.call_count < batch_runner._MAX_DRAFT_ROUNDS


def test_a_single_empty_round_is_retried(tmp_path, monkeypatch):
    """An empty answer is not proof of exhaustion.

    ask() has already exhausted its own internal retries by this point, so
    an empty result here means "that query produced nothing" -- which
    covers a timeout, a wedged browser and a refusal as well as a genuinely
    dry market. Retrying once recovers the first three.
    """
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Empty Then Good")

    calls = {"n": 0}

    def _empty_then_good(*_a, **_kw):
        calls["n"] += 1
        if calls["n"] == 1:
            return []          # the transient failure
        return [_draft(f"Recovered Co {calls['n']}",
                       website=f"https://r{calls['n']}.example.com")]

    with patch(
        "vendor_intel.pipeline.market_research.discovery.discover_and_draft",
        side_effect=_empty_then_good,
    ), patch.object(batch_runner, "_build_row", _mock_build_row_pass_through):
        rows = batch_runner._run_one_batch(
            state, "Test Market Empty Then Good", batch_size=2)

    assert rows, "an empty first round must not abandon the batch"
    assert calls["n"] >= 2, "must have asked again after the empty round"


def test_two_consecutive_empty_rounds_stop_the_batch(tmp_path, monkeypatch):
    """The exhaustion exit must survive: a genuinely dry market should not
    keep querying for all eight rounds."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Dry")

    with patch(
        "vendor_intel.pipeline.market_research.discovery.discover_and_draft",
        return_value=[],
    ) as mock_discover:
        rows = batch_runner._run_one_batch(state, "Test Market Dry", batch_size=20)

    assert rows == []
    assert mock_discover.call_count == 2, (
        "a dry market must stop after two empties, not burn every round"
    )


def test_the_empty_counter_resets_after_a_good_round(tmp_path, monkeypatch):
    """Otherwise two empties spread across a long batch would stop it even
    though productive rounds happened in between."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Alternating")

    sequence = [
        [],                                           # empty (1)
        [_draft("Alt One", website="https://a1.example.com")],
        [],                                           # empty (1 again, reset)
        [_draft("Alt Two", website="https://a2.example.com")],
        [],                                           # empty (1)
        [],                                           # empty (2) -> stop
    ]
    calls = {"n": 0}

    def _next(*_a, **_kw):
        i = calls["n"]
        calls["n"] += 1
        return sequence[i] if i < len(sequence) else []

    with patch(
        "vendor_intel.pipeline.market_research.discovery.discover_and_draft",
        side_effect=_next,
    ), patch.object(batch_runner, "_build_row", _mock_build_row_pass_through):
        rows = batch_runner._run_one_batch(
            state, "Test Market Alternating", batch_size=20)

    assert len(rows) == 2, "both productive rounds must be kept"
    assert calls["n"] == len(sequence), (
        "an interleaved empty round must not stop the batch early"
    )


def test_run_until_target_loops_batches_until_reached(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)

    call_count = {"n": 0}

    def fake_run_one_batch(state, market, *, batch_size):
        call_count["n"] += 1
        start = len(state.companies) + 1
        return [
            _mock_build_row_pass_through(_draft(f"Company {i}", website=f"https://c{i}.example.com"))
            for i in range(start, start + batch_size)
        ]

    with patch.object(batch_runner, "_run_one_batch", side_effect=fake_run_one_batch):
        summary = batch_runner.run_until_target("Test Market Target", target=45, batch_size=20)

    # 20 + 20 + 20 = 60, the first total >= target 45 (batches aren't trimmed
    # mid-batch to hit the target exactly).
    assert "60/45" in summary
    assert call_count["n"] == 3


def test_run_until_target_retries_before_giving_up_on_empty_batches(tmp_path, monkeypatch):
    """A batch returning 0 companies can mean genuine exhaustion OR a
    transient failure (proxy timeout + fallback also timing out) — this bit
    us for real: a --target 500 run stopped permanently at 90 companies
    because one batch happened to time out on both providers. Must retry
    several times (not give up on the very first empty batch) before
    concluding the market is genuinely exhausted."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    monkeypatch.setattr("time.sleep", lambda _: None)  # don't actually pause in tests

    with patch.object(batch_runner, "_run_one_batch", return_value=[]) as mock_run_one_batch:
        summary = batch_runner.run_until_target("Test Market Exhausted", target=500, batch_size=20)

    assert "0/500" in summary
    assert mock_run_one_batch.call_count == batch_runner._MAX_CONSECUTIVE_EMPTY_BATCHES


def test_run_until_target_recovers_after_a_transient_empty_batch(tmp_path, monkeypatch):
    """A single empty batch (e.g. one bad-luck timeout) must NOT end the
    run early if a later batch succeeds — this is the exact bug: previously
    ANY empty batch ended the whole run immediately."""
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    monkeypatch.setattr("time.sleep", lambda _: None)

    call_count = {"n": 0}

    def fake_run_one_batch(state, market, *, batch_size):
        call_count["n"] += 1
        if call_count["n"] == 1:
            return []  # transient failure on the first attempt
        start = len(state.companies) + 1
        return [
            _mock_build_row_pass_through(_draft(f"Company {i}", website=f"https://c{i}.example.com"))
            for i in range(start, start + batch_size)
        ]

    with patch.object(batch_runner, "_run_one_batch", side_effect=fake_run_one_batch):
        summary = batch_runner.run_until_target("Test Market Recovers", target=20, batch_size=20)

    assert "20/20" in summary


# --- live Excel refresh ---------------------------------------------------
#
# Requested: the definition goes in first, then companies get filled in.
# The sheet is therefore rewritten after each completed batch so it is
# inspectable DURING a run, not only at the end.


def test_export_refresh_is_called_after_each_batch(monkeypatch):
    from vendor_intel.pipeline.market_research import batch_runner as br

    refreshed = []
    monkeypatch.setattr(br, "_refresh_export", lambda m: refreshed.append(m))
    monkeypatch.setattr(br, "_backend_is_blocked", lambda: False)
    monkeypatch.setattr(br, "_backend_refused", lambda: False)
    monkeypatch.setattr(br.time, "sleep", lambda _s: None)

    batches = {"n": 0}

    def _one_then_done(_state, _market, **_kw):
        batches["n"] += 1
        return [{"Company": f"Co {batches['n']}"}] if batches["n"] <= 2 else []

    monkeypatch.setattr(br, "_run_one_batch", _one_then_done)

    class _State:
        companies: list = []

        @staticmethod
        def load(_m):
            return _State()

        def add_batch(self, rows):
            type(self).companies = list(self.companies) + list(rows)

        def save(self):
            pass

    _State.companies = []
    monkeypatch.setattr(br, "MarketResearchState", _State)

    br.run_until_target("Global Refresh Market", target=2, batch_size=1)

    assert refreshed, "the Excel must be refreshed as batches complete"
    assert all(m == "Global Refresh Market" for m in refreshed)


def test_a_failing_refresh_does_not_abort_discovery(monkeypatch, capsys):
    """Most likely cause is the workbook being open in Excel, which locks
    it on Windows. The companies are already in the state file, so losing
    a refresh must never cost the run."""
    from vendor_intel.pipeline.market_research import batch_runner as br

    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.export_xlsx.export_xlsx",
        lambda _m: (_ for _ in ()).throw(PermissionError("file is open in Excel")),
    )
    monkeypatch.delenv("MARKET_RESEARCH_LIVE_EXPORT", raising=False)

    br._refresh_export("Global Locked File Market")   # must not raise

    assert "could not refresh" in capsys.readouterr().out


def test_live_export_can_be_switched_off(monkeypatch):
    from vendor_intel.pipeline.market_research import batch_runner as br

    called = []
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.export_xlsx.export_xlsx",
        lambda _m: called.append(_m),
    )
    monkeypatch.setenv("MARKET_RESEARCH_LIVE_EXPORT", "false")

    br._refresh_export("Global Off Market")
    assert called == []


# --- the market name was being echoed back as the category --------------
#
# Measured on live data before this guard: 227 of 414 Dry Eyes rows (55%)
# and 33 of 80 Solar Rooftop rows (41%) had
#
#     Core Categories = "Global Solar Rooftop Market - Distributors Only"
#
# The column was worthless in over half the sheet, and worse, it hid
# whether the company sells the right product -- exactly what the product
# scope exists to establish.
#
# Cause: the prompt interpolated the market into the field it asked about
# ('core_categories: its core categories within "{market}"'), making the
# echo the path of least resistance. The prompt no longer restates the
# market; this guard backs that up, because wording alone is not a
# guarantee.


@pytest.mark.parametrize(
    "value",
    [
        "Global Solar Rooftop Market - Distributors Only",   # exact, as seen live
        "Global Solar Rooftop Market",                        # suffix dropped
        "global solar rooftop market",                        # lowercased
        "Global  Solar-Rooftop  Market!",                     # repunctuated
        "GLOBAL SOLAR ROOFTOP MARKET - DISTRIBUTORS ONLY",
    ],
)
def test_a_category_that_echoes_the_market_is_rejected(value):
    assert batch_runner._echoes_market(
        value, "Global Solar Rooftop Market - Distributors Only"
    ), value


@pytest.mark.parametrize(
    "value",
    [
        "PV modules; inverters; mounting systems",
        "Solar panel wholesale distribution",   # mentions solar, still real
        "Photovoltaic modules and balance-of-system components",
        "Ophthalmic pharmaceuticals",
        "",
    ],
)
def test_a_real_category_is_kept(value):
    """A false positive here silently discards a good row, so mentioning
    the market's subject matter must not be enough to trigger it."""
    assert not batch_runner._echoes_market(
        value, "Global Solar Rooftop Market - Distributors Only"
    ), value


def test_an_echoed_category_is_blanked_not_kept():
    market = "Global Dry Eyes Treatment Market - Distributors Only"
    assert batch_runner._clean_category(market, market) == ""
    assert batch_runner._clean_category("Artificial tears; OTC eye care", market) == (
        "Artificial tears; OTC eye care"
    )


def test_no_market_means_no_filtering():
    """_build_row is called directly by tests and other callers without a
    market; that must not start blanking their categories."""
    assert batch_runner._clean_category("anything at all", "") == "anything at all"


def test_build_row_blanks_an_echoed_category():
    """End-to-end through the real _build_row."""
    market = "Global Solar Rooftop Market - Distributors Only"
    item = _draft("Echo Co", website="https://echo.example.com")
    item["core_categories"] = market
    row = batch_runner._build_row(item, market=market)
    # Blanked rather than dropped: an echoed category is a bad VALUE, not
    # proof the company is wrong. Core Categories is a REQUIRED_FIELD, so
    # finalize_row() rejects the row afterwards and the export audit
    # records why.
    assert row is not None
    assert row["Core Categories"] == ""


def test_build_row_keeps_a_real_category():
    market = "Global Solar Rooftop Market - Distributors Only"
    item = _draft("Real Co", website="https://real.example.com")
    item["core_categories"] = "PV modules; string inverters"
    row = batch_runner._build_row(item, market=market)
    assert row["Core Categories"] == "PV modules; string inverters"


def test_the_prompt_no_longer_interpolates_the_market_into_the_category():
    """The originating cause. Asking for "categories within <market>" put
    the string the model then echoed directly into the instruction."""
    from vendor_intel.pipeline.market_research import discovery

    assert '- core_categories: its core categories within "{market}".' not in (
        discovery._DRAFT_PROMPT
    )
    assert "NEVER repeat the market name" in discovery._DRAFT_PROMPT


# --- a market can dry up without ever returning an EMPTY batch ----------
#
# _MAX_CONSECUTIVE_EMPTY_BATCHES only counts batches that produced zero
# companies, so a market trickling +1 and +2 never trips it.
#
# Observed on Polyurethane: batches of +1 and +2 while every round came
# back with one or two names already collected (211 website duplicates and
# 29 name duplicates rejected across the run). The remaining ~95
# companies were never going to arrive, and each batch spent queries the
# next market needed.
#
# Stopping is not abandoning the target. It reports the number the market
# actually contains -- the alternative is padding it with the duplicates
# the dedupe just rejected.


def test_repeated_low_yield_batches_stop_the_run(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)

    calls = {"n": 0}

    def _trickle(state, market, *, batch_size):
        calls["n"] += 1
        start = len(state.companies) + 1
        # One new company per batch: never empty, never enough.
        return [_mock_build_row_pass_through(
            _draft(f"Company {start}", website=f"https://c{start}.example.com"))]

    with patch.object(batch_runner, "_run_one_batch", side_effect=_trickle):
        batch_runner.run_until_target("Trickle Market", target=400, batch_size=20)

    assert calls["n"] <= batch_runner._MAX_LOW_YIELD_BATCHES + 1, (
        f"stopped after {calls['n']} batches; a trickling market must not "
        "run to the target"
    )


def test_a_single_weak_batch_does_not_stop_the_run(tmp_path, monkeypatch):
    """One weak batch is normal noise. Polyurethane produced +4, then +9,
    +13, +16 -- stopping at the first dip would have cost ~100 real
    companies."""
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)

    sizes = iter([1, 18, 18, 18])      # weak, then strong
    calls = {"n": 0}

    def _mixed(state, market, *, batch_size):
        calls["n"] += 1
        n = next(sizes, 0)
        start = len(state.companies) + 1
        return [
            _mock_build_row_pass_through(
                _draft(f"Company {i}", website=f"https://c{i}.example.com"))
            for i in range(start, start + n)
        ]

    with patch.object(batch_runner, "_run_one_batch", side_effect=_mixed):
        batch_runner.run_until_target("Mixed Market", target=50, batch_size=20)

    assert calls["n"] >= 3, "a recovery after one weak batch must be allowed"


def test_the_low_yield_counter_resets_on_a_good_batch(tmp_path, monkeypatch):
    """Two weak batches separated by a strong one are not exhaustion."""
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)

    sizes = iter([1, 1, 18, 1, 1, 18])
    calls = {"n": 0}

    def _alternating(state, market, *, batch_size):
        calls["n"] += 1
        n = next(sizes, 0)
        start = len(state.companies) + 1
        return [
            _mock_build_row_pass_through(
                _draft(f"Company {i}", website=f"https://c{i}.example.com"))
            for i in range(start, start + n)
        ]

    with patch.object(batch_runner, "_run_one_batch", side_effect=_alternating):
        batch_runner.run_until_target("Alternating Market", target=200,
                                      batch_size=20)

    assert calls["n"] >= 5, (
        "an interleaved strong batch must reset the counter, not be ignored"
    )


def test_the_threshold_is_a_small_fraction_of_a_batch():
    """Set too high, this would stop healthy markets early -- Solar
    Rooftop had legitimate +14 and +15 batches near the end."""
    assert batch_runner._LOW_YIELD_THRESHOLD <= 3
    assert batch_runner._MAX_LOW_YIELD_BATCHES >= 3
