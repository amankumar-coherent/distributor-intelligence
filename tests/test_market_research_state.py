"""state.py: dedup/persistence, "give me more 20" exclusion semantics."""
from __future__ import annotations

from vendor_intel.pipeline.market_research.state import MarketResearchState, market_slug, state_path


def test_market_slug_is_filesystem_safe():
    assert market_slug("Global Semiconductor Distributor Market") == "global-semiconductor-distributor-market"
    assert market_slug("  Weird!! Market??  ") == "weird-market"


def test_fresh_state_has_no_companies(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Alpha")
    assert state.last_index == 0
    assert state.companies == []
    assert state.excluded_names() == []
    assert not state.is_seen("Anything")


def test_add_batch_persists_and_dedupes(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    state = MarketResearchState.load("Test Market Beta")

    batch1 = [
        {"#": "1", "Company": "Acme Inc"},
        {"#": "2", "Company": "Beta Corp"},
    ]
    state.add_batch(batch1)
    state.save()

    assert state.last_index == 2
    assert state.is_seen("Acme Inc")
    assert state.is_seen("Acme")  # normalized dedupe key strips "Inc"

    reloaded = MarketResearchState.load("Test Market Beta")
    assert reloaded.last_index == 2
    assert reloaded.excluded_names() == ["Acme Inc", "Beta Corp"]

    # "give me more 20": a repeat name is excluded, numbering continues.
    reloaded.add_batch([
        {"#": "3", "Company": "Gamma Ltd"},
        {"#": "1", "Company": "Acme Inc"},  # duplicate, ignored even if numbered wrong
    ])
    assert reloaded.last_index == 3
    assert [c["Company"] for c in reloaded.companies] == ["Acme Inc", "Beta Corp", "Gamma Ltd"]


def test_state_path_uses_slug(tmp_path, monkeypatch):
    monkeypatch.setattr("vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    path = state_path("Global EV Battery Market")
    assert path.parent == tmp_path
    assert path.name == "global-ev-battery-market.json"


# --- the same company under a different name ----------------------------
#
# Measured on finished markets: 13% of Dry Eyes rows and 19% of Solar
# Rooftop rows were one company listed repeatedly --
#
#     powernsun.com   PowernSun / Power & Sun / Power n Sun /
#                     Power & Sun FZCO / Powernsun Group / ...  (7 rows)
#
# company_dedupe_key normalises punctuation and legal suffixes, but those
# names genuinely differ, so it cannot catch them. A company controls one
# primary domain, which is much stronger identity evidence than a name.


def test_name_variants_sharing_a_website_are_rejected(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    state = MarketResearchState(market="Website Dedupe Test")
    state.add_batch([{"Company": "PowernSun",
                      "Website": "https://powernsun.com",
                      "Headquarters": "Dubai, United Arab Emirates"}])
    state.add_batch([{"Company": "Power & Sun",
                      "Website": "https://www.powernsun.com",
                      "Headquarters": "Dubai, United Arab Emirates"}])
    state.add_batch([{"Company": "Powernsun Group",
                      "Website": "powernsun.com/",
                      "Headquarters": "Dubai, United Arab Emirates"}])

    assert len(state.companies) == 1, "three names, one company"


def test_national_subsidiaries_on_one_domain_are_all_kept(tmp_path, monkeypatch):
    """Krannich Solar runs five real national distributors on one domain,
    each with its own geography and contacts. Collapsing them by domain
    alone would delete four real companies -- which is why country is part
    of the key."""
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    state = MarketResearchState(market="Subsidiary Test")
    for company, hq in [
        ("Krannich Solar GmbH & Co. KG", "Weil der Stadt, Germany"),
        ("Krannich Solar Turkey", "Istanbul, Turkey"),
        ("Krannich Solar d.o.o.", "Zagreb, Croatia"),
        ("Krannich Solar Energy (Pty) Ltd", "Cape Town, South Africa"),
        ("Krannich Solar Private Limited", "Pune, India"),
    ]:
        state.add_batch([{"Company": company,
                          "Website": "https://krannich-solar.com",
                          "Headquarters": hq}])

    assert len(state.companies) == 5


def test_rows_without_a_website_are_never_collapsed(tmp_path, monkeypatch):
    """No domain means no identity evidence; falling back to "" as a
    shared key would merge every website-less company into one."""
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    state = MarketResearchState(market="No Website Test")
    state.add_batch([{"Company": "Alpha Trading", "Website": "",
                      "Headquarters": "Nairobi, Kenya"}])
    state.add_batch([{"Company": "Beta Trading", "Website": "",
                      "Headquarters": "Nairobi, Kenya"}])
    assert len(state.companies) == 2


def test_the_website_keys_work_on_a_state_file_written_before_this_check(
    tmp_path, monkeypatch
):
    """Several state files already on disk predate this dedupe. The keys
    are derived from `companies` rather than persisted, so a loaded file
    is protected immediately."""
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    old = MarketResearchState(
        market="Legacy",
        companies=[{"Company": "PowernSun", "Website": "https://powernsun.com",
                    "Headquarters": "Dubai, United Arab Emirates"}],
    )
    old.save()

    loaded = MarketResearchState.load("Legacy")
    loaded.add_batch([{"Company": "Power & Sun",
                       "Website": "https://www.powernsun.com",
                       "Headquarters": "Dubai, United Arab Emirates"}])
    assert len(loaded.companies) == 1


def test_is_seen_row_reports_both_kinds_of_duplicate(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "vendor_intel.pipeline.market_research.state.STATE_DIR", tmp_path)
    from vendor_intel.pipeline.market_research.state import MarketResearchState

    state = MarketResearchState(market="is_seen_row")
    state.add_batch([{"Company": "PowernSun",
                      "Website": "https://powernsun.com",
                      "Headquarters": "Dubai, United Arab Emirates"}])

    # by name
    assert state.is_seen_row({"Company": "PowernSun", "Website": "",
                              "Headquarters": ""})
    # by website+country
    assert state.is_seen_row({"Company": "Power & Sun",
                              "Website": "https://www.powernsun.com",
                              "Headquarters": "Dubai, United Arab Emirates"})
    # genuinely new
    assert not state.is_seen_row({"Company": "Segen",
                                  "Website": "https://segensolar.co.za",
                                  "Headquarters": "Cape Town, South Africa"})
