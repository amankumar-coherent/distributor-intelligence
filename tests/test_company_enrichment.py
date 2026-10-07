"""Unit tests for the Kompass -> ChatGPT -> Google AI -> DeepSeek -> Smart
Crawl per-field enrichment orchestrator. All external sources are mocked —
this tests the merge/skip/fallback control flow, not the sources themselves
(those already have their own test coverage elsewhere)."""
from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, patch

import pytest

from vendor_intel.pipeline import company_enrichment as ce


class TestIsMissing:
    @pytest.mark.parametrize(
        "value",
        [None, "", "   ", "N/A", "n/a", "NA", "Not Available", "Not Found", "Unknown", "None", "null"],
    )
    def test_recognizes_missing_sentinels(self, value):
        assert ce.is_missing(value) is True

    @pytest.mark.parametrize(
        "value",
        ["0", 0, "0 employees", "2000", "Not publicly disclosed", "Acme Inc", "  Acme  "],
    )
    def test_does_not_flag_legitimate_values(self, value):
        assert ce.is_missing(value) is False


class TestMergeCompanyData:
    def test_rule1_never_overwrite_valid_with_empty(self):
        existing = {"Founded": "2015"}
        merged = ce.merge_company_data(existing, {"Founded": ""}, source="chatgpt_scraper")
        assert merged["Founded"] == "2015"

    def test_rule2_fill_empty_cells(self):
        existing = {"Founded": ""}
        merged = ce.merge_company_data(existing, {"Founded": "2015"}, source="deepseek")
        assert merged["Founded"] == "2015"

    def test_rule4_does_not_replace_present_value_with_different_value(self):
        existing = {"Employees": "50"}
        merged = ce.merge_company_data(existing, {"Employees": "999"}, source="google_ai")
        assert merged["Employees"] == "50"

    def test_prefer_source_for_overrides_present_value(self):
        """Smart Crawl is documented as authoritative for website-derived
        fields — this is the one case a present value CAN be replaced."""
        existing = {"Key Brands Represented": "GuessedBrand"}
        merged = ce.merge_company_data(
            existing,
            {"Key Brands Represented": "Philips; Siemens"},
            source="smart_crawl",
            prefer_source_for=ce.WEBSITE_DERIVED_COLUMNS,
        )
        assert merged["Key Brands Represented"] == "Philips; Siemens"

    def test_prefer_source_for_does_not_affect_other_columns(self):
        existing = {"Founded": "2015", "Key Brands Represented": "GuessedBrand"}
        merged = ce.merge_company_data(
            existing,
            {"Founded": "1999", "Key Brands Represented": "Philips"},
            source="smart_crawl",
            prefer_source_for=ce.WEBSITE_DERIVED_COLUMNS,
        )
        assert merged["Founded"] == "2015"  # not in prefer_source_for -> untouched
        assert merged["Key Brands Represented"] == "Philips"

    def test_field_sources_tracked(self):
        field_sources: dict[str, str] = {}
        ce.merge_company_data(
            {"Founded": ""}, {"Founded": "2015"}, source="kompass", field_sources=field_sources
        )
        assert field_sources["Founded"] == "kompass"

    def test_ignores_unmanaged_columns(self):
        merged = ce.merge_company_data({}, {"NotARealColumn": "value"}, source="deepseek")
        assert "NotARealColumn" not in merged


class TestAcceptedContactRoles:
    @pytest.mark.parametrize(
        "role",
        [
            "VP of Sales", "Vice President of Sales", "Head of Sales", "Sales Director",
            "Director of Sales", "Sales Manager", "Chief Sales Officer", "CSO",
            "Head of Procurement", "Procurement Manager", "Procurement Director",
            "Purchasing Manager", "Purchasing Director", "Chief Procurement Officer", "CPO",
        ],
    )
    def test_accepts_sales_and_procurement_leadership(self, role):
        assert ce.is_accepted_contact_role(role) is True

    @pytest.mark.parametrize(
        "role",
        [
            "", "CEO", "Founder", "Software Engineer", "Marketing Manager",
            "Receptionist", "General Manager", "Accountant", "HR Manager",
            "Business Development Manager",  # deliberately excluded per user's scope decision
        ],
    )
    def test_rejects_non_matching_roles(self, role):
        assert ce.is_accepted_contact_role(role) is False


class TestContactRoleGate:
    """The sales/procurement-only rule for Contact Person/Role/LinkedIn —
    enforced universally in merge_company_data() so it applies regardless of
    which of the 5 enrichment sources found the contact."""

    def test_matching_role_merges_all_three_contact_fields(self):
        new_data = {"Contact Person": "Jane Doe", "Role": "VP of Sales", "LinkedIn": "linkedin.com/in/janedoe"}
        merged = ce.merge_company_data({}, new_data, source="chatgpt_scraper")
        assert merged["Contact Person"] == "Jane Doe"
        assert merged["Role"] == "VP of Sales"
        assert merged["LinkedIn"] == "linkedin.com/in/janedoe"

    def test_non_matching_role_strips_all_three_contact_fields(self):
        """A CEO or generic manager must be rejected entirely — not merged
        with a mismatched role."""
        new_data = {"Contact Person": "John Smith", "Role": "CEO", "LinkedIn": "linkedin.com/in/johnsmith"}
        merged = ce.merge_company_data({}, new_data, source="chatgpt_scraper")
        assert "Contact Person" not in merged
        assert "Role" not in merged
        assert "LinkedIn" not in merged

    def test_contact_person_with_no_role_at_all_is_rejected(self):
        """No role text means the type can't be confirmed — must not merge
        just a name with an unconfirmed title."""
        new_data = {"Contact Person": "Jane Doe", "LinkedIn": "linkedin.com/in/janedoe"}
        merged = ce.merge_company_data({}, new_data, source="google_ai")
        assert "Contact Person" not in merged
        assert "LinkedIn" not in merged

    def test_gate_does_not_affect_unrelated_fields_in_same_update(self):
        """A rejected contact must not block OTHER fields in the same
        source update from merging normally."""
        new_data = {"Founded": "2015", "Contact Person": "John Smith", "Role": "CEO"}
        merged = ce.merge_company_data({}, new_data, source="deepseek")
        assert merged["Founded"] == "2015"
        assert "Contact Person" not in merged

    def test_gate_applies_to_kompass_source_too(self):
        """Per explicit scope decision: the rule applies to ALL sources,
        not just the AI-search stages — including Kompass."""
        new_data = {"Contact Person": "Real Person", "Role": "Warehouse Supervisor"}
        merged = ce.merge_company_data({}, new_data, source="kompass")
        assert "Contact Person" not in merged

    def test_gate_applies_to_smart_crawl_source_too(self):
        new_data = {"Contact Person": "Real Person", "Role": "Receptionist"}
        merged = ce.merge_company_data({}, new_data, source="smart_crawl")
        assert "Contact Person" not in merged

    def test_existing_good_contact_not_overwritten_by_rejected_new_one(self):
        existing = {"Contact Person": "Jane Doe", "Role": "VP of Sales"}
        new_data = {"Contact Person": "John Smith", "Role": "CEO"}
        merged = ce.merge_company_data(existing, new_data, source="chatgpt_scraper")
        assert merged["Contact Person"] == "Jane Doe"
        assert merged["Role"] == "VP of Sales"


class TestContactRolePromptHint:
    def test_hint_included_when_contact_fields_requested(self):
        prompt = ce._field_search_prompt("Acme Inc", "", "", ["Contact Person", "Role"])
        assert "sales or procurement LEADERSHIP" in prompt

    def test_hint_omitted_when_contact_fields_not_requested(self):
        prompt = ce._field_search_prompt("Acme Inc", "", "", ["Founded", "Employees"])
        assert "sales or procurement LEADERSHIP" not in prompt


class TestGeographyPromptHint:
    def test_hint_included_when_geography_fields_requested(self):
        prompt = ce._field_search_prompt("Acme Inc", "", "", ["Continent / Geography"])
        assert "physically BASED" in prompt
        prompt2 = ce._field_search_prompt("Acme Inc", "", "", ["Operational Presence"])
        assert "AS A DISTRIBUTOR" in prompt2

    def test_hint_distinguishes_hq_from_distribution_territory(self):
        """The two columns must be presented as answering different
        questions, not one implying the other — per explicit distributor
        definition: HQ/offices vs. actual sales/distribution territory."""
        prompt = ce._field_search_prompt("Acme Inc", "", "", ["Operational Presence"])
        assert "NOT automatically the same as where it is headquartered" in prompt
        assert "not evidence it distributes internationally" in prompt

    def test_hint_omitted_when_geography_fields_not_requested(self):
        prompt = ce._field_search_prompt("Acme Inc", "", "", ["Founded", "Employees"])
        assert "physically BASED" not in prompt
        assert "AS A DISTRIBUTOR" not in prompt

    def test_both_hints_present_when_both_kinds_of_fields_requested(self):
        prompt = ce._field_search_prompt(
            "Acme Inc", "", "", ["Contact Person", "Continent / Geography"]
        )
        assert "sales or procurement LEADERSHIP" in prompt
        assert "physically BASED" in prompt


class TestMarketFieldUnderstanding:
    """ChatGPT scraper defines what Core Categories/Specialty Focus/Key
    Brands Represented mean for a SPECIFIC market -> DeepSeek distills that
    into prompt guidance -> reused for every company in that market."""

    def test_disabled_via_env_returns_empty(self):
        with patch.dict("os.environ", {"MARKET_FIELD_UNDERSTANDING_ENABLED": "false"}):
            assert ce.understand_market_fields("Flexible Packaging") == ""

    def test_empty_market_returns_empty(self):
        assert ce.understand_market_fields("") == ""

    def test_falls_back_to_empty_when_chatgpt_scraper_disabled(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=False), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search", return_value=None
        ):
            assert ce.understand_market_fields("Flexible Packaging") == ""

    def test_full_chain_definition_then_distillation(self):
        raw_explanation = (
            "Core Categories would include flexible films, pouches, and laminates. "
            "Specialty Focus might cover barrier films or biodegradable packaging. "
            "Key Brands would include companies like Amcor, Sealed Air, and Berry Global."
        )
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", return_value=raw_explanation
        ), patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-v4-flash"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json",
            return_value={
                "core_categories_guidance": "Flexible films, pouches, laminates.",
                "specialty_focus_guidance": "Barrier films, biodegradable packaging.",
                "key_brands_guidance": "Amcor, Sealed Air, Berry Global.",
            },
        ), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search", return_value=None
        ), patch("vendor_intel.clients.search_cache_disk.put_disk_cached_search") as mock_put:
            guidance = ce.understand_market_fields("Flexible Packaging")

        assert "Flexible films" in guidance
        assert "Amcor" in guidance
        mock_put.assert_called_once()

    def test_no_deepseek_key_falls_back_to_empty(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", return_value="Some explanation."
        ), patch("vendor_intel.clients.search_cache_disk.get_disk_cached_search", return_value=None):
            import os

            os.environ.pop("OPENAI_API_KEY", None)
            guidance = ce.understand_market_fields("Flexible Packaging")
        assert guidance == ""

    def test_uses_cache_on_second_call(self):
        with patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            return_value=[{"raw": "cached raw", "guidance": "cached guidance"}],
        ), patch("vendor_intel.scraping.google_ai_mode.ask_generative") as mock_ask:
            guidance = ce.understand_market_fields("Flexible Packaging")
        assert guidance == "cached guidance"
        mock_ask.assert_not_called()

    def test_guidance_included_in_prompt_for_market_field_columns(self):
        prompt = ce._field_search_prompt(
            "Acme Inc", "", "", ["Core Categories"], market_field_guidance="Some market guidance here."
        )
        assert "Some market guidance here." in prompt

    def test_guidance_omitted_when_not_a_market_field_column(self):
        prompt = ce._field_search_prompt(
            "Acme Inc", "", "", ["Founded"], market_field_guidance="Some market guidance here."
        )
        assert "Some market guidance here." not in prompt

    def test_chatgpt_scraper_stage_passes_market_to_understanding(self):
        row = {c: "" for c in ce.MANAGED_COLUMNS}
        row["Company"] = "Acme Inc"
        captured = {}

        def fake_understand(market):
            captured["market"] = market
            return "test guidance"

        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", return_value="{}"
        ), patch.object(ce, "understand_market_fields", side_effect=fake_understand):
            ce.enrich_from_chatgpt_scraper(row, market="Flexible Packaging")

        assert captured["market"] == "Flexible Packaging"


class TestWriteCompanySummary:
    def test_skipped_when_no_company_name(self):
        row = {"Company": ""}
        result_row, ran = ce.write_company_summary(row)
        assert ran is False

    def test_skipped_when_real_summary_already_present(self):
        row = {
            "Company": "Acme Inc",
            "Summary": "Acme Inc is a well-established distributor founded in 1990.",
            "Founded": "1990",
        }
        result_row, ran = ce.write_company_summary(row)
        assert ran is False

    def test_rewrites_kompass_metadata_dump_summary(self):
        row = {
            "Company": "Acme Inc",
            "Summary": "Kompass Company ID: MD001143 | Founded: 1996 | Market: Smartphones",
            "Founded": "1996",
            "Headquarters": "Chisinau, Moldova",
            "Employees": "83",
        }
        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-v4-flash"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json",
            return_value={"summary": "Acme Inc is a Moldova-based distributor with 83 employees, founded in 1996."},
        ):
            result_row, ran = ce.write_company_summary(row)
        assert ran is True
        assert "Moldova" in result_row["Summary"]
        assert "MD001143" not in result_row["Summary"]

    def test_skipped_when_not_enough_data(self):
        row = {"Company": "Acme Inc", "Founded": "1996"}  # only 1 fact
        result_row, ran = ce.write_company_summary(row)
        assert ran is False

    def test_skipped_when_no_api_key(self, monkeypatch):
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        row = {"Company": "Acme Inc", "Founded": "1996", "Employees": "83"}
        result_row, ran = ce.write_company_summary(row)
        assert ran is False

    def test_error_keeps_existing_summary(self):
        row = {"Company": "Acme Inc", "Founded": "1996", "Employees": "83", "Summary": ""}
        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-v4-flash"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json", side_effect=RuntimeError("api down")
        ):
            result_row, ran = ce.write_company_summary(row)
        assert ran is False
        assert result_row["Summary"] == ""

    def test_never_invents_facts_not_in_row_prompt_only_uses_given_facts(self):
        """Sanity check on the prompt itself: only facts actually present in
        the row are sent to DeepSeek — nothing else is available for it to
        reference, structurally preventing invention of unrelated facts."""
        row = {
            "Company": "Acme Inc", "Founded": "1996", "Employees": "83",
            "Headquarters": "Chisinau, Moldova",
        }
        captured = {}

        def fake_chat(client, model, system, user, **kwargs):
            captured["user"] = user
            return {"summary": "Acme Inc, founded 1996, based in Chisinau with 83 employees."}

        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-v4-flash"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json", side_effect=fake_chat
        ):
            ce.write_company_summary(row)

        assert "1996" in captured["user"]
        assert "Chisinau, Moldova" in captured["user"]
        assert "83" in captured["user"]


class TestGeographyUnion:
    def test_union_continent_geography_combines_regions(self):
        result = ce._union_geography_value(
            "Continent / Geography", "APAC(India)", "APAC(Singapore); EU(Germany)"
        )
        assert "India" in result
        assert "Singapore" in result
        assert "Germany" in result

    def test_union_operational_presence_combines_plain_lists(self):
        result = ce._union_geography_value("Operational Presence", "India", "Singapore; Germany")
        assert result == "India; Singapore; Germany"

    def test_union_no_new_countries_keeps_existing_unchanged(self):
        result = ce._union_geography_value("Operational Presence", "India; Singapore", "India")
        assert result == "India; Singapore"

    def test_union_unparseable_new_value_keeps_existing(self):
        result = ce._union_geography_value("Operational Presence", "India", "Not publicly disclosed")
        assert result == "India"

    def test_merge_company_data_unions_geography_instead_of_locking(self):
        """Regression: the default merge rule only fills EMPTY cells — a
        second source's additional countries must still be added even
        though Continent / Geography is already non-empty from an earlier
        source, per explicit instruction: write ALL locations found."""
        existing = {"Continent / Geography": "APAC(India)"}
        merged = ce.merge_company_data(
            existing, {"Continent / Geography": "APAC(Singapore)"}, source="chatgpt_scraper"
        )
        assert "India" in merged["Continent / Geography"]
        assert "Singapore" in merged["Continent / Geography"]

    def test_merge_company_data_unions_operational_presence_across_three_sources(self):
        row = {}
        row = ce.merge_company_data(row, {"Operational Presence": "India"}, source="kompass")
        row = ce.merge_company_data(row, {"Operational Presence": "Germany"}, source="chatgpt_scraper")
        row = ce.merge_company_data(row, {"Operational Presence": "Brazil"}, source="deepseek")
        assert "India" in row["Operational Presence"]
        assert "Germany" in row["Operational Presence"]
        assert "Brazil" in row["Operational Presence"]

    def test_merge_company_data_geography_union_tracks_combined_field_source(self):
        field_sources: dict[str, str] = {}
        row = ce.merge_company_data(
            {}, {"Operational Presence": "India"}, source="kompass", field_sources=field_sources
        )
        row = ce.merge_company_data(
            row, {"Operational Presence": "Germany"}, source="chatgpt_scraper", field_sources=field_sources
        )
        assert "kompass" in field_sources["Operational Presence"]
        assert "chatgpt_scraper" in field_sources["Operational Presence"]

    def test_other_columns_still_use_normal_fill_once_rule(self):
        """Sanity: the union behavior is scoped ONLY to the two geography
        columns — everything else keeps the original fill-once-then-lock
        rule (Rule 4)."""
        existing = {"Founded": "2015"}
        merged = ce.merge_company_data(existing, {"Founded": "1999"}, source="chatgpt_scraper")
        assert merged["Founded"] == "2015"


class TestMissingColumns:
    def test_returns_only_missing_managed_columns(self):
        row = {c: "value" for c in ce.MANAGED_COLUMNS}
        row["Founded"] = ""
        row["Employees"] = "N/A"
        gaps = ce.missing_columns(row)
        assert set(gaps) == {"Founded", "Employees"}

    def test_no_gaps_when_all_present(self):
        row = {c: "value" for c in ce.MANAGED_COLUMNS}
        assert ce.missing_columns(row) == []


def _all_filled_row(**overrides) -> dict[str, str]:
    row = {c: "value" for c in ce.MANAGED_COLUMNS}
    row.update(overrides)
    return row


class TestKompassNotPubliclyDisclosedFiller:
    """Regression: kd.to_slim_row() fills unfound Kompass fields with the
    literal string "Not publicly disclosed" as a display default — that
    filler must NOT be treated as a real found value, or every later stage
    (ChatGPT/Google/DeepSeek/SmartCrawl) would wrongly skip searching for it."""

    def _fake_verified_distributor(self, **overrides):
        from vendor_intel.discovery.kompass_distributors import VerifiedDistributor

        defaults = dict(
            company_name="ABC Medical",
            market="Smartphones",
            kompass_company_id="MD000001",
            kompass_profile_url="https://www.kompass.com/z/ww/c/abc-medical/md000001/",
            kompass_product_id="4066041",
            kompass_company_type="Distributor",
            is_actual_market_distributor=True,
            verification_confidence=90,
            website="https://abcmedical.com",
            founded_year="2015",
            employees="",  # genuinely not found by Kompass
            brands="",  # genuinely not found by Kompass
        )
        defaults.update(overrides)
        return VerifiedDistributor(**defaults)

    def test_kompass_filler_does_not_block_chatgpt_from_searching(self):
        """Kompass finds Founded/Website but not Employees or Key Brands —
        those two must still be asked of ChatGPT, not silently accepted as
        'Not publicly disclosed'."""
        row = _all_filled_row(
            Founded="", Website="", Employees="", **{"Key Brands Represented": ""}
        )
        row["Company"] = "ABC Medical"

        distributor = self._fake_verified_distributor()
        with patch.object(ce, "_kompass_row_for_company") as mock_lookup:
            mock_lookup.return_value = ce._raw_kompass_fields(distributor)
            merged, ran = ce.enrich_from_kompass(row, market="Smartphones")

        assert ran is True
        assert merged["Founded"] == "2015"
        assert merged["Website"] == "https://abcmedical.com"
        # These were genuinely not found by Kompass — must remain missing,
        # NOT "Not publicly disclosed", so ChatGPT/etc. still search for them.
        assert ce.is_missing(merged["Employees"])
        assert ce.is_missing(merged["Key Brands Represented"])

    def test_raw_kompass_fields_strips_not_publicly_disclosed_filler(self):
        distributor = self._fake_verified_distributor(employees="Not publicly disclosed", brands="Not publicly disclosed")
        fields = ce._raw_kompass_fields(distributor)
        assert fields["Employees"] == ""
        assert fields["Key Brands Represented"] == ""

    def test_raw_kompass_fields_keeps_real_values(self):
        distributor = self._fake_verified_distributor(employees="83", brands="Philips; Siemens")
        fields = ce._raw_kompass_fields(distributor)
        assert fields["Employees"] == "83"
        assert fields["Key Brands Represented"] == "Philips; Siemens"

    def test_kompass_never_overwrites_a_field_it_did_not_find(self):
        """enrich_from_kompass must merge, not replace the whole row —
        a field already present from an earlier stage must survive even
        though kd's raw fields dict has an empty string for it."""
        row = _all_filled_row(Employees="50", Founded="")  # Founded missing so Kompass actually runs
        row["Company"] = "ABC Medical"
        distributor = self._fake_verified_distributor(employees="")
        with patch.object(ce, "_kompass_row_for_company") as mock_lookup:
            mock_lookup.return_value = ce._raw_kompass_fields(distributor)
            merged, ran = ce.enrich_from_kompass(row, market="Smartphones")
        assert ran is True
        assert merged["Employees"] == "50"


class TestStageSkipping:
    """Each stage must be skipped (no API/network call) when there is
    nothing left to search for — the explicit performance requirement."""

    def test_kompass_skipped_when_no_gaps(self):
        row = _all_filled_row()
        with patch("vendor_intel.discovery.kompass_distributors.find_market_product_page") as mock_find:
            result_row, ran = ce.enrich_from_kompass(row, market="Smartphones")
        assert ran is False
        mock_find.assert_not_called()

    def test_chatgpt_scraper_skipped_when_no_gaps(self):
        row = _all_filled_row()
        with patch("vendor_intel.scraping.google_ai_mode.ask_generative") as mock_ask:
            result_row, ran = ce.enrich_from_chatgpt_scraper(row)
        assert ran is False
        mock_ask.assert_not_called()

    def test_chatgpt_scraper_skipped_when_disabled(self):
        row = _all_filled_row(Founded="")
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=False), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative"
        ) as mock_ask:
            result_row, ran = ce.enrich_from_chatgpt_scraper(row)
        assert ran is False
        mock_ask.assert_not_called()

    def test_google_ai_skipped_when_no_gaps(self):
        row = _all_filled_row()
        with patch(
            "vendor_intel.clients.google_ai_scraper.google_ai_gap_fill_enabled", return_value=True
        ), patch("vendor_intel.clients.google_ai_scraper.google_ai_ping", new=AsyncMock()) as mock_ping:
            result_row, ran = asyncio.run(ce.enrich_from_google_ai(row, market="Smartphones"))
        assert ran is False
        mock_ping.assert_not_called()

    def test_deepseek_skipped_when_no_gaps(self):
        row = _all_filled_row()
        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json"
        ) as mock_chat:
            result_row, ran = ce.enrich_from_deepseek(row, market="Smartphones")
        assert ran is False
        mock_chat.assert_not_called()

    def test_deepseek_skipped_when_no_api_key(self, monkeypatch):
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        row = _all_filled_row(Founded="")
        with patch("vendor_intel.pipeline.chatgpt_expand._chat_json") as mock_chat:
            result_row, ran = ce.enrich_from_deepseek(row, market="Smartphones")
        assert ran is False
        mock_chat.assert_not_called()

    def test_smart_crawl_skipped_when_no_website_derived_gaps(self):
        row = _all_filled_row(Website="https://example.com")
        with patch("crawler.smart_crawl.smart_crawl", new=AsyncMock()) as mock_crawl:
            result_row, ran = asyncio.run(ce.enrich_from_smart_crawl(row))
        assert ran is False
        mock_crawl.assert_not_called()

    def test_smart_crawl_skipped_when_no_website_known(self):
        row = _all_filled_row(Website="", **{"Key Brands Represented": ""})
        with patch("crawler.smart_crawl.smart_crawl", new=AsyncMock()) as mock_crawl:
            result_row, ran = asyncio.run(ce.enrich_from_smart_crawl(row))
        assert ran is False
        mock_crawl.assert_not_called()


class TestOnlyMissingFieldsRequested:
    def test_chatgpt_scraper_prompt_lists_only_missing_fields(self):
        row = _all_filled_row(Employees="", **{"Office No.": "", "Office Email": ""})
        row["Company"] = "ABC Medical"
        row["Founded"] = "2015"  # present — must NOT appear in the "find" list

        captured_prompt = {}

        def fake_ask(prompt, timeout=90):
            captured_prompt["text"] = prompt
            return "{}"

        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", side_effect=fake_ask
        ):
            ce.enrich_from_chatgpt_scraper(row)

        prompt = captured_prompt["text"]
        assert "Employees" in prompt
        assert "Office No." in prompt
        assert "Office Email" in prompt
        assert "Founded" not in prompt  # already present — must not be re-asked


class TestErrorHandlingNeverStopsPipeline:
    def test_kompass_exception_falls_through_to_next_stage(self):
        row = {c: "" for c in ce.MANAGED_COLUMNS}
        row["Company"] = "ABC Medical"

        with patch(
            "vendor_intel.discovery.kompass_distributors.find_market_product_page",
            side_effect=RuntimeError("kompass down"),
        ), patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=False), patch(
            "vendor_intel.clients.google_ai_scraper.google_ai_gap_fill_enabled", return_value=False
        ), patch.dict("os.environ", {}, clear=False):
            import os

            os.environ.pop("OPENAI_API_KEY", None)
            result = asyncio.run(ce.enrich_company(row, market="Smartphones"))

        assert "kompass" in result.stages_skipped or any("kompass" in e for e in result.errors)
        # Pipeline must complete without raising, regardless of Kompass failure.
        assert result.row["Company"] == "ABC Medical"

    def test_single_stage_exception_does_not_abort_remaining_stages(self):
        row = {c: "" for c in ce.MANAGED_COLUMNS}
        row["Company"] = "ABC Medical"

        async def raise_google_ai(r, market):
            raise RuntimeError("google ai relay down")

        with patch.object(ce, "enrich_from_kompass", return_value=(row, False)), patch.object(
            ce, "enrich_from_chatgpt_scraper", return_value=(row, False)
        ), patch.object(ce, "enrich_from_google_ai", side_effect=raise_google_ai), patch.object(
            ce, "enrich_from_deepseek", return_value=(row, False)
        ), patch.object(
            ce, "verify_website_stage", return_value=(row, False)
        ), patch.object(
            ce, "enrich_from_smart_crawl", new=AsyncMock(return_value=(row, False))
        ):
            result = asyncio.run(ce.enrich_company(row, market="Smartphones"))

        assert any("google_ai" in e for e in result.errors)
        assert result.row["Company"] == "ABC Medical"  # pipeline completed


class TestWebsiteVerifyStage:
    def test_skipped_when_no_company_name(self):
        row = {"Company": "", "Website": "https://example.com"}
        result_row, ran = ce.verify_website_stage(row)
        assert ran is False

    def test_corrects_website_and_reports_ran_true(self):
        from vendor_intel.pipeline.website_verification import WebsiteVerificationResult

        fake_result = WebsiteVerificationResult(
            website="https://real-acme.com", is_reachable=True, is_identity_match=True,
            was_corrected=True, source="chatgpt_scraper",
        )
        with patch(
            "vendor_intel.pipeline.website_verification.verify_and_fix_website",
            return_value=fake_result,
        ):
            row = {"Company": "Acme Inc", "Website": "https://dead-domain.com"}
            result_row, ran = ce.verify_website_stage(row)
        assert ran is True
        assert result_row["Website"] == "https://real-acme.com"

    def test_no_correction_needed_reports_ran_false(self):
        from vendor_intel.pipeline.website_verification import WebsiteVerificationResult

        fake_result = WebsiteVerificationResult(
            website="https://acme.com", is_reachable=True, is_identity_match=True,
            was_corrected=False, source="unchanged",
        )
        with patch(
            "vendor_intel.pipeline.website_verification.verify_and_fix_website",
            return_value=fake_result,
        ):
            row = {"Company": "Acme Inc", "Website": "https://acme.com"}
            result_row, ran = ce.verify_website_stage(row)
        assert ran is False
        assert result_row["Website"] == "https://acme.com"

    def test_dead_website_with_no_replacement_cleared(self):
        from vendor_intel.pipeline.website_verification import WebsiteVerificationResult

        fake_result = WebsiteVerificationResult(
            website="", is_reachable=False, is_identity_match=False,
            was_corrected=False, source="none_found",
        )
        with patch(
            "vendor_intel.pipeline.website_verification.verify_and_fix_website",
            return_value=fake_result,
        ):
            row = {"Company": "Acme Inc", "Website": "https://dead-domain.com"}
            result_row, ran = ce.verify_website_stage(row)
        assert ran is True
        assert result_row["Website"] == ""

    def test_error_keeps_existing_website(self):
        with patch(
            "vendor_intel.pipeline.website_verification.verify_and_fix_website",
            side_effect=RuntimeError("network down"),
        ):
            row = {"Company": "Acme Inc", "Website": "https://acme.com"}
            result_row, ran = ce.verify_website_stage(row)
        assert ran is False
        assert result_row["Website"] == "https://acme.com"

    def test_runs_between_deepseek_and_smart_crawl_in_full_pipeline(self):
        """End-to-end wiring check: verify_website_stage must run as part
        of enrich_company()'s stage sequence and be able to fix a website
        before Smart Crawl sees it."""
        base_row = {c: "" for c in ce.MANAGED_COLUMNS}
        base_row["Company"] = "ABC Medical"
        base_row["Website"] = "https://dead-old-site.com"

        from vendor_intel.pipeline.website_verification import WebsiteVerificationResult

        fixed = WebsiteVerificationResult(
            website="https://real-abcmedical.com", is_reachable=True, is_identity_match=True,
            was_corrected=True, source="chatgpt_scraper",
        )

        crawled_domain = {}

        async def fake_smart_crawl(row):
            crawled_domain["website"] = row.get("Website")
            return row, False

        with patch.object(ce, "enrich_from_kompass", return_value=(base_row, False)), patch.object(
            ce, "enrich_from_chatgpt_scraper", return_value=(base_row, False)
        ), patch.object(
            ce, "enrich_from_google_ai", new=AsyncMock(side_effect=lambda r, market: (r, False))
        ), patch.object(ce, "enrich_from_deepseek", return_value=(base_row, False)), patch(
            "vendor_intel.pipeline.website_verification.verify_and_fix_website", return_value=fixed
        ), patch.object(ce, "enrich_from_smart_crawl", new=AsyncMock(side_effect=fake_smart_crawl)):
            result = asyncio.run(ce.enrich_company(base_row, market="Smartphones"))

        assert result.row["Website"] == "https://real-abcmedical.com"
        assert crawled_domain["website"] == "https://real-abcmedical.com"
        assert "website_verify" in result.stages_run


class TestFieldSourceTracking:
    def test_field_sources_recorded_per_stage(self):
        base_row = {c: "" for c in ce.MANAGED_COLUMNS}
        base_row["Company"] = "ABC Medical"

        def fake_kompass(row, market):
            out = dict(row)
            out["Founded"] = "2015"
            return out, True

        async def fake_smart_crawl(row):
            out = dict(row)
            out["Office Email"] = "info@abcmedical.com"
            return out, True

        with patch.object(ce, "enrich_from_kompass", side_effect=fake_kompass), patch.object(
            ce, "enrich_from_chatgpt_scraper", return_value=(base_row, False)
        ), patch.object(
            ce, "enrich_from_google_ai", new=AsyncMock(side_effect=lambda r, market: (r, False))
        ), patch.object(ce, "enrich_from_deepseek", return_value=(base_row, False)), patch.object(
            ce, "verify_website_stage", return_value=(base_row, False)
        ), patch.object(
            ce, "enrich_from_smart_crawl", new=AsyncMock(side_effect=fake_smart_crawl)
        ):
            result = asyncio.run(ce.enrich_company(base_row, market="Smartphones"))

        assert result.field_sources.get("Founded") == "kompass"
        assert result.field_sources.get("Office Email") == "smart_crawl"
        assert "kompass" in result.stages_run
        assert "smart_crawl" in result.stages_run
        assert "chatgpt_scraper" in result.stages_skipped


class TestPolishTextFields:
    def test_skipped_when_nothing_needs_polish(self):
        row = {"Specialty Focus": "Clean text.", "Core Categories": "Smartphones"}
        with patch("vendor_intel.pipeline.chatgpt_expand._chat_json") as mock_chat:
            result_row, polished = ce.polish_text_fields(row)
        assert polished is False
        mock_chat.assert_not_called()

    def test_skipped_when_no_api_key(self, monkeypatch):
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        row = {"Specialty Focus": "messy   text\nwith breaks"}
        with patch("vendor_intel.pipeline.chatgpt_expand._chat_json") as mock_chat:
            result_row, polished = ce.polish_text_fields(row)
        assert polished is False
        mock_chat.assert_not_called()

    def test_polishes_only_flagged_fields(self):
        row = {
            "Specialty Focus": "messy   text\nwith breaks",
            "Core Categories": "Smartphones",  # clean — should not be sent
            "Website": "https://example.com",  # not a POLISH_COLUMNS field at all
        }
        captured = {}

        def fake_chat(client, model, system, user, **kwargs):
            captured["user"] = user
            return {"Specialty Focus": "Clean, well-formatted text."}

        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-chat"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json", side_effect=fake_chat
        ):
            result_row, polished = ce.polish_text_fields(row)

        assert polished is True
        assert result_row["Specialty Focus"] == "Clean, well-formatted text."
        assert result_row["Core Categories"] == "Smartphones"  # untouched
        assert "Core Categories" not in captured["user"]  # never sent — was already clean
        assert "Website" not in captured["user"]  # not a polish column at all

    def test_never_lets_polish_empty_out_a_field(self):
        row = {"Specialty Focus": "messy   text\nwith breaks"}
        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-chat"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json", return_value={"Specialty Focus": ""}
        ):
            result_row, polished = ce.polish_text_fields(row)
        # Model returned empty for a field that had real content — must keep the original.
        assert result_row["Specialty Focus"] == "messy   text\nwith breaks"

    def test_error_keeps_unpolished_text(self):
        row = {"Specialty Focus": "messy   text\nwith breaks"}
        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-chat"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json", side_effect=RuntimeError("api down")
        ):
            result_row, polished = ce.polish_text_fields(row)
        assert polished is False
        assert result_row["Specialty Focus"] == "messy   text\nwith breaks"

    def test_ignores_unrequested_keys_in_response(self):
        """If the model hallucinates an extra key not in POLISH_COLUMNS or
        not one we asked about, it must be ignored, not merged in."""
        row = {"Specialty Focus": "messy   text\nwith breaks"}
        with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-chat"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json",
            return_value={"Specialty Focus": "Cleaned.", "Founded": "1999"},
        ):
            result_row, polished = ce.polish_text_fields(row)
        assert result_row["Specialty Focus"] == "Cleaned."
        assert "Founded" not in result_row


class TestEnrichmentSummary:
    def test_records_website_discovery(self):
        summary = ce.EnrichmentSummary()
        enriched = ce.EnrichedCompany(
            row={c: "value" for c in ce.MANAGED_COLUMNS},
            field_sources={"Website": "kompass"},
            stages_run=["kompass"],
        )
        summary.record(enriched, had_website_before=False)
        assert summary.websites_discovered == 1
        assert summary.total_companies == 1
        assert summary.stage_run_counts["kompass"] == 1

    def test_does_not_count_website_discovery_when_already_present(self):
        summary = ce.EnrichmentSummary()
        enriched = ce.EnrichedCompany(row={c: "value" for c in ce.MANAGED_COLUMNS})
        summary.record(enriched, had_website_before=True)
        assert summary.websites_discovered == 0


class _FakeCkpt:
    """Minimal in-memory stand-in for ExpandCheckpoint (bump/data/save)."""

    def __init__(self, initial_data: dict | None = None):
        self._data: dict = dict(initial_data or {})
        self.bump_calls: list[tuple[str, str]] = []

    def bump(self, step: str, substep: str, *, note: str = "", **data) -> None:
        self.bump_calls.append((step, substep))
        self._data.update(data)

    def data(self, key: str, default=None):
        return self._data.get(key, default)

    def save(self, note: str = "") -> None:
        pass


class TestEnrichCompanyRowsCheckpointResume:
    """enrich_company_rows(rows, ckpt=...) must resume at COMPANY-BOUNDARY
    granularity: a killed run skips every already-completed company and
    only redoes companies that never finished their full chain."""

    def _rows(self, n: int) -> list[dict[str, str]]:
        return [{**{c: "" for c in ce.MANAGED_COLUMNS}, "Company": f"Company {i}"} for i in range(n)]

    def test_fresh_run_processes_every_company_and_saves_progress(self):
        rows = self._rows(25)
        calls: list[str] = []

        async def fake_enrich_company(row, *, market, company_index=None, total_companies=None):
            calls.append(row["Company"])
            return ce.EnrichedCompany(row=dict(row), stages_run=["kompass"])

        ckpt = _FakeCkpt()
        with patch.object(ce, "enrich_company", side_effect=fake_enrich_company):
            out, summary = asyncio.run(ce.enrich_company_rows(rows, market="Smartphones", ckpt=ckpt))

        assert len(calls) == 25
        assert len(out) == 25
        assert summary.total_companies == 25
        assert len(ckpt.data("enrich_done_keys")) == 25
        # saved at least twice for 25 companies at save-every-10 cadence
        assert len(ckpt.bump_calls) >= 2

    def test_resume_skips_already_completed_companies(self):
        rows = self._rows(5)
        calls: list[str] = []

        async def fake_enrich_company(row, *, market, company_index=None, total_companies=None):
            calls.append(row["Company"])
            return ce.EnrichedCompany(row=dict(row), stages_run=["kompass"])

        prior_state = {
            "enrich_done_keys": ["1:company 0", "2:company 1", "3:company 2"],
            "enrich_out_rows": [dict(rows[0]), dict(rows[1]), dict(rows[2])],
            "enrich_summary_state": {
                "total_companies": 3,
                "stage_run_counts": {},
                "stage_skip_counts": {},
                "fields_filled_by_source": {},
                "still_missing_total": 0,
                "errors": [],
                "websites_discovered": 0,
            },
        }
        ckpt = _FakeCkpt(prior_state)

        with patch.object(ce, "enrich_company", side_effect=fake_enrich_company):
            out, summary = asyncio.run(ce.enrich_company_rows(rows, market="Smartphones", ckpt=ckpt))

        # only companies 3 and 4 (the not-yet-done ones) should have run
        assert calls == ["Company 3", "Company 4"]
        assert len(out) == 5
        assert summary.total_companies == 5  # 3 carried over + 2 new

    def test_no_checkpoint_behaves_like_before(self):
        rows = self._rows(3)

        async def fake_enrich_company(row, *, market, company_index=None, total_companies=None):
            return ce.EnrichedCompany(row=dict(row), stages_run=["kompass"])

        with patch.object(ce, "enrich_company", side_effect=fake_enrich_company):
            out, summary = asyncio.run(ce.enrich_company_rows(rows, market="Smartphones"))

        assert len(out) == 3
        assert summary.total_companies == 3
