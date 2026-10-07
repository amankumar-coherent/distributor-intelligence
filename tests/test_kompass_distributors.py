"""Unit tests for the Kompass distributor scraper's deterministic logic —
URL/ID extraction, market-specific evidence matching, scoring, and dedup.
No network calls; AI discovery and HTTP fetch are exercised separately via
mocks where needed.
"""
from __future__ import annotations

import pytest

from vendor_intel.clients import kompass as kompass_client
from vendor_intel.discovery import kompass_distributors as kd


class TestFinalDistributorGate:
    """The deterministic final gate applied to EVERY row before export,
    regardless of source (Kompass, ChatGPT scraper, Google AI, DeepSeek,
    Smart Crawl) — closes the gap where non-Kompass rows previously only
    had a soft LLM-prompt instruction to avoid manufacturers/OEMs."""

    def test_kompass_row_with_structured_company_type_pure_distributor_kept(self):
        row = {"Company": "Exida Ltd", "Summary": "Company Type: Distributor | Market: Smartphones"}
        assert kd.is_manufacturer_or_oem_row(row) is False

    def test_kompass_row_with_structured_company_type_manufacturer_rejected(self):
        row = {"Company": "Acme Corp", "Summary": "Company Type: Manufacturer Distributor | Market: Smartphones"}
        assert kd.is_manufacturer_or_oem_row(row) is True

    def test_kompass_row_pure_manufacturer_rejected(self):
        row = {"Company": "Acme Corp", "Summary": "Company Type: Manufacturer | Market: Smartphones"}
        assert kd.is_manufacturer_or_oem_row(row) is True

    def test_non_kompass_row_with_manufacturer_free_text_rejected(self):
        """A ChatGPT/DeepSeek-sourced row has no 'Company Type:' field at
        all — the gate must still catch manufacturer language in whatever
        free text is available."""
        row = {
            "Company": "Global Plastics Inc",
            "Ownership": "Privately owned; manufactures plastic packaging films in-house",
        }
        assert kd.is_manufacturer_or_oem_row(row) is True

    def test_non_kompass_row_pure_distributor_language_kept(self):
        row = {
            "Company": "Regional Distributors LLC",
            "Ownership": "Independent distributor of packaging materials",
        }
        assert kd.is_manufacturer_or_oem_row(row) is False

    def test_row_with_no_type_signal_at_all_not_rejected(self):
        """Absence of evidence is not evidence of manufacturing — must not
        reject a row just because no company-type text exists anywhere."""
        row = {"Company": "Unknown Co", "Website": "https://example.com"}
        assert kd.is_manufacturer_or_oem_row(row) is False

    def test_mentioning_manufacturers_as_suppliers_not_self_manufacturer(self):
        """'Distributes products from various manufacturers' describes WHO
        they distribute for, not that the company itself manufactures —
        must not be rejected (only applies when there's no structured
        Company Type: field to trust as authoritative)."""
        row = {
            "Company": "Multi-Brand Distributors Ltd",
            "Ownership": "Distributor representing several manufacturers across the region",
        }
        assert kd.is_manufacturer_or_oem_row(row) is False

    def test_apply_final_gate_splits_kept_and_rejected(self):
        rows = [
            {"Company": "Good Co", "Summary": "Company Type: Distributor"},
            {"Company": "Bad Co", "Summary": "Company Type: Manufacturer"},
            {"Company": "Combo Co", "Summary": "Company Type: Manufacturer Distributor"},
            {"Company": "Neutral Co", "Website": "https://example.com"},
        ]
        kept, rejected = kd.apply_final_distributor_gate(rows)
        kept_names = {r["Company"] for r in kept}
        rejected_names = {r["Company"] for r in rejected}
        assert kept_names == {"Good Co", "Neutral Co"}
        assert rejected_names == {"Bad Co", "Combo Co"}

    def test_apply_final_gate_preserves_row_order_within_each_group(self):
        rows = [
            {"Company": "A", "Summary": "Company Type: Distributor"},
            {"Company": "B", "Summary": "Company Type: Manufacturer"},
            {"Company": "C", "Summary": "Company Type: Distributor"},
        ]
        kept, rejected = kd.apply_final_distributor_gate(rows)
        assert [r["Company"] for r in kept] == ["A", "C"]
        assert [r["Company"] for r in rejected] == ["B"]


class TestCompanyCountLimits:
    def test_max_companies_default_is_unlimited(self, monkeypatch):
        monkeypatch.delenv("MAX_COMPANIES_PER_MARKET", raising=False)
        assert kd.max_companies_per_market() == 0

    def test_max_companies_zero_env_is_unlimited(self, monkeypatch):
        monkeypatch.setenv("MAX_COMPANIES_PER_MARKET", "0")
        assert kd.max_companies_per_market() == 0

    def test_max_companies_positive_env_is_respected(self, monkeypatch):
        monkeypatch.setenv("MAX_COMPANIES_PER_MARKET", "250")
        assert kd.max_companies_per_market() == 250

    def test_min_companies_default_is_500(self, monkeypatch):
        monkeypatch.delenv("MIN_COMPANIES_PER_MARKET", raising=False)
        assert kd.min_companies_per_market() == 500

    def test_min_companies_env_override(self, monkeypatch):
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "800")
        assert kd.min_companies_per_market() == 800


class TestPagination:
    """Kompass paginates by PATH (".../23200/page-2/") via numbered
    "pageBtn" links and has NO "Next"/"»" link. Matching only a next-link
    (the original behaviour) silently stopped after page 1 — observed live
    as 77 companies discovered out of 37,528 listed."""

    _REAL_MARKUP = (
        '<div class="pagination blockPagination">'
        '<a href="https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/"'
        ' class="pageBtn active">1</a>'
        '<a href="https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/page-2/"'
        ' class="pageBtn ">2</a>'
        '<a href="https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/page-3/"'
        ' class="pageBtn ">3</a>'
        '<a href="https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/page-5/"'
        ' class="pageBtn ">5</a>'
        "</div>"
    )
    _BASE = "https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/"

    def test_finds_page_2_from_numbered_links_with_no_next_link(self):
        assert kd._find_next_page(self._REAL_MARKUP, self._BASE) == (
            "https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/page-2/"
        )

    def test_advances_one_page_at_a_time_from_a_numbered_page(self):
        nxt = kd._find_next_page(self._REAL_MARKUP, self._BASE + "page-2/")
        assert nxt == (
            "https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/page-3/"
        )

    def test_skips_gaps_when_intermediate_pages_are_not_linked(self):
        """The listing links a sparse subset (2, 3, 5, ...) — from page 3 the
        next linked page is 5, not 4."""
        nxt = kd._find_next_page(self._REAL_MARKUP, self._BASE + "page-3/")
        assert nxt == (
            "https://lb.kompass.com/z/ww/a/medical-equipment-and-instruments/23200/page-5/"
        )

    def test_returns_none_on_the_last_page(self):
        assert kd._find_next_page(self._REAL_MARKUP, self._BASE + "page-5/") is None

    def test_never_goes_backwards_or_repeats_current_page(self):
        for page in ("", "page-2/", "page-3/"):
            nxt = kd._find_next_page(self._REAL_MARKUP, self._BASE + page)
            assert nxt != self._BASE + page

    def test_conventional_next_link_still_wins_when_present(self):
        html = '<a href="/z/ww/a/x/1/page-9/">Next</a>' + self._REAL_MARKUP
        assert kd._find_next_page(html, self._BASE).endswith("/page-9/")

    def test_no_pagination_markup_returns_none(self):
        assert kd._find_next_page("<div>no pages here</div>", self._BASE) is None


class TestScreeningCheckpointResume:
    """Screening must checkpoint incrementally, not only at the very end.

    Regression: the original version screened all 9,437 companies in one
    call and saved once afterward, so a crash at company 1,225 discarded
    every distributor found up to that point.
    """

    @pytest.fixture(autouse=True)
    def _use_screening_path(self, monkeypatch):
        # KOMPASS_GOOGLE_AI_SCREEN is the live gate name; it takes priority
        # over the legacy KOMPASS_CHATGPT_SCREEN alias, so it must be the one
        # set here or conftest's "false" default wins and these tests
        # silently exercise the Kompass profile-fetch path instead.
        monkeypatch.setenv("KOMPASS_GOOGLE_AI_SCREEN", "true")

    def _candidates(self, n):
        return [
            kd.KompassCompanyCandidate(
                name=f"Company {i}",
                profile_url=f"https://www.kompass.com/z/ww/c/company-{i}/md{i:06d}/",
            )
            for i in range(n)
        ]

    def test_saves_progress_before_the_whole_list_is_screened(self, monkeypatch):
        candidates = self._candidates(250)
        monkeypatch.setattr(kd, "find_market_product_page", lambda m: ("https://k/p/1", "1"))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)

        saves = []

        def fake_screen(chunk, *, market, kompass_product_id):
            saves.append(len(chunk))
            return [], len(chunk), 0, set()

        monkeypatch.setattr(kd, "_screen_candidates_via_chatgpt", fake_screen)

        ckpt = _FakeCheckpoint()
        kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)

        # 250 companies at 100 per checkpoint => 3 chunks, so progress was
        # persisted twice BEFORE the run finished.
        assert len(saves) == 3
        assert len(ckpt.data("kompass_processed_urls")) == 250

    def test_resume_skips_already_screened_companies(self, monkeypatch):
        candidates = self._candidates(150)
        monkeypatch.setattr(kd, "find_market_product_page", lambda m: ("https://k/p/1", "1"))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)

        screened_names = []

        def fake_screen(chunk, *, market, kompass_product_id):
            screened_names.extend(c.name for c in chunk)
            return [], len(chunk), 0, set()

        monkeypatch.setattr(kd, "_screen_candidates_via_chatgpt", fake_screen)

        # Pretend the first 100 were screened before a crash.
        already = [c.profile_url.rstrip("/").lower() for c in candidates[:100]]
        ckpt = _FakeCheckpoint(
            {
                "kompass_product_url": "https://k/p/1",
                "kompass_product_id": "1",
                "kompass_candidates": [
                    {"name": c.name, "profile_url": c.profile_url, "country_hint": ""}
                    for c in candidates
                ],
                "kompass_processed_urls": already,
                "kompass_accepted": [],
                "kompass_rejected_count": 100,
                "kompass_unreachable_count": 0,
            }
        )

        kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)

        # Only the remaining 50 get re-screened — not all 150.
        assert len(screened_names) == 50
        assert screened_names[0] == "Company 100"

    def test_distributors_found_before_a_crash_survive_in_the_checkpoint(self, monkeypatch):
        """The actual loss in the live failure: 19 distributors discarded."""
        candidates = self._candidates(250)
        monkeypatch.setattr(kd, "find_market_product_page", lambda m: ("https://k/p/1", "1"))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)

        calls = {"n": 0}

        def fake_screen(chunk, *, market, kompass_product_id):
            calls["n"] += 1
            if calls["n"] == 2:
                raise RuntimeError("browser died mid-run")
            found = [
                kd.VerifiedDistributor(
                    company_name=chunk[0].name,
                    market=market,
                    kompass_company_id="MD000001",
                    kompass_profile_url=chunk[0].profile_url,
                    kompass_product_id=kompass_product_id,
                    kompass_company_type="Distributor (per ChatGPT screening)",
                    is_actual_market_distributor=True,
                    verification_confidence=60,
                )
            ]
            return found, len(chunk) - 1, 0, set()

        monkeypatch.setattr(kd, "_screen_candidates_via_chatgpt", fake_screen)

        ckpt = _FakeCheckpoint()
        with pytest.raises(RuntimeError):
            kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)

        # The first chunk's distributor was saved before the crash.
        assert len(ckpt.data("kompass_accepted", [])) == 1
        assert len(ckpt.data("kompass_processed_urls", [])) == 100


class TestUrlAndIdExtraction:
    def test_company_id_extracted_from_company_url(self):
        url = "https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/"
        assert kompass_client.extract_company_id(url) == "MD001143"

    def test_company_slug_extracted(self):
        url = "https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/"
        assert kompass_client.extract_company_slug(url) == "accent-electronic-sa"

    def test_product_id_extracted_from_product_url(self):
        url = "https://www.kompass.com/en/b/telecoms/mobile-devices/smartphones/4066041/"
        assert kompass_client.extract_product_id(url) == "4066041"

    def test_product_id_is_none_for_company_url(self):
        """A company URL must never be misread as a product page — product id
        extraction should refuse to run on it at all."""
        url = "https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/"
        assert kompass_client.extract_product_id(url) is None

    def test_company_id_is_none_for_product_url(self):
        """The reverse mistake: a product URL must never yield a company id."""
        url = "https://www.kompass.com/en/b/telecoms/mobile-devices/smartphones/4066041/"
        assert kompass_client.extract_company_id(url) is None

    def test_product_id_and_company_id_never_equal_for_real_examples(self):
        company_url = "https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/"
        product_url = "https://www.kompass.com/en/b/telecoms/mobile-devices/smartphones/4066041/"
        cid = kompass_client.extract_company_id(company_url)
        pid = kompass_client.extract_product_id(product_url)
        assert cid != pid
        assert cid == "MD001143"
        assert pid == "4066041"

    def test_non_kompass_url_rejected(self):
        assert kompass_client.is_kompass_url("https://example.com/z/ww/c/foo/md001143/") is False
        assert kompass_client.extract_company_id("https://example.com/z/ww/c/foo/md001143/") is None

    def test_global_region_code_extracted(self):
        url = "https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/"
        assert kompass_client.extract_global_region_code(url) == "ww"

    def test_global_region_code_none_when_absent(self):
        url = "https://www.kompass.com/en/b/telecoms/mobile-devices/smartphones/4066041/"
        assert kompass_client.extract_global_region_code(url) is None


class TestMarketSpecificEvidence:
    def test_accepts_explicit_distribution_of_market_term(self):
        html = "<p>ABC Electronics is a distributor of smartphones across the region.</p>"
        found, reason = kd._has_market_specific_distribution_evidence(html, "Smartphones")
        assert found is True

    def test_rejects_keyword_only_in_software_context(self):
        """Spec §5/§26: 'software for smartphones' must NOT count as
        distribution evidence, even though the market keyword is present."""
        html = "<p>JNC Sistemas provides software for smartphones and mobile devices.</p>"
        found, reason = kd._has_market_specific_distribution_evidence(html, "Smartphones")
        assert found is False

    def test_rejects_keyword_with_no_distribution_context_at_all(self):
        html = "<p>Founded in 1996. We manufacture industrial cables and connectors.</p>"
        found, reason = kd._has_market_specific_distribution_evidence(html, "Smartphones")
        assert found is False

    def test_rejects_generic_it_distributor_unrelated_to_market(self):
        """A company that is a general IT distributor (printers, computers)
        with no smartphone-specific evidence must be rejected."""
        html = (
            "<p>We are an authorized distributor of printers, computers, and "
            "office equipment across Eastern Europe.</p>"
        )
        found, reason = kd._has_market_specific_distribution_evidence(html, "Smartphones")
        assert found is False

    def test_accepts_reversed_order_term_then_role_word(self):
        html = "<p>Our smartphone wholesaler network covers 12 countries.</p>"
        found, _ = kd._has_market_specific_distribution_evidence(html, "Smartphones")
        assert found is True

    def test_no_market_terms_returns_false(self):
        found, reason = kd._has_market_specific_distribution_evidence("<p>anything</p>", "Market")
        assert found is False
        assert reason == "no_market_terms"


class TestVerificationScoring:
    def test_full_evidence_yields_high_score_and_acceptance(self):
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="Distributor Service",
            market_evidence_found=True,
            website_confirms=True,
        )
        assert score == 90
        assert accepted is True

    def test_distributor_type_without_market_evidence_is_rejected(self):
        """Spec's core rule: distributor status ALONE is not enough."""
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="Distributor",
            market_evidence_found=False,
        )
        assert accepted is False
        assert score == 40  # only the distributor-type point

    def test_non_distributor_type_is_always_rejected_even_with_market_evidence(self):
        """A pure 'Service' or 'Manufacturer' company type, even with strong
        market evidence, must not be accepted — Kompass must say Distributor."""
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="Service",
            market_evidence_found=True,
        )
        assert accepted is False

    def test_below_confidence_threshold_is_rejected(self, monkeypatch):
        monkeypatch.setenv("MIN_DISTRIBUTOR_CONFIDENCE", "90")
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="Distributor",
            market_evidence_found=True,
        )
        assert score == 70  # 40 + 30
        assert accepted is False  # below 90 threshold

    def test_manufacturer_distributor_combo_rejected(self):
        """User explicitly excludes ANY manufacturer/OEM involvement, even
        combined with a Distributor role — stricter than the original spec,
        which would have accepted this combo."""
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="Manufacturer Distributor",
            market_evidence_found=True,
        )
        assert accepted is False

    def test_oem_distributor_combo_rejected(self):
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="OEM Distributor",
            market_evidence_found=True,
        )
        assert accepted is False

    def test_pure_distributor_with_no_manufacturer_role_accepted(self):
        accepted, score = kd.verify_and_score(
            market="Smartphones",
            company_type="Distributor Service",
            market_evidence_found=True,
        )
        assert accepted is True
        assert score == 70


class TestDeduplication:
    def _mk(self, name, company_id, country="India"):
        return kd.VerifiedDistributor(
            company_name=name,
            market="Smartphones",
            kompass_company_id=company_id,
            kompass_profile_url=f"https://www.kompass.com/z/ww/c/{name}/{company_id}/",
            kompass_product_id="4066041",
            kompass_company_type="Distributor",
            is_actual_market_distributor=True,
            verification_confidence=90,
            country=country,
        )

    def test_dedup_by_company_id(self):
        rows = [self._mk("Acme", "MD001143"), self._mk("Acme Inc", "MD001143")]
        out = kd.deduplicate(rows)
        assert len(out) == 1

    def test_dedup_by_name_and_country_when_no_id_collision(self):
        a = self._mk("Acme Electronics", "MD000001", country="India")
        b = kd.VerifiedDistributor(
            company_name="Acme Electronics",
            market="Smartphones",
            kompass_company_id="",
            kompass_profile_url="https://www.kompass.com/z/ww/c/acme/x/",
            kompass_product_id="4066041",
            kompass_company_type="Distributor",
            is_actual_market_distributor=True,
            verification_confidence=90,
            country="India",
        )
        out = kd.deduplicate([a, b])
        assert len(out) == 1

    def test_no_false_dedup_across_different_countries(self):
        a = self._mk("Acme Electronics", "MD000001", country="India")
        b = self._mk("Acme Electronics", "MD000002", country="Germany")
        out = kd.deduplicate([a, b])
        assert len(out) == 2


class TestActivityTreeExtraction:
    """Regression coverage using real Kompass markup shapes, captured live
    against https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/
    (the exact company from the integration spec's example)."""

    _REAL_ACTIVITY_HTML = """
    <div class="classCat cat1">
        <a id="4066041" class="activityLink" href="https://www.kompass.com/z/ww/a/smartphones/4066041/">
            <i class="circle distributor" ></i>
            <span class="text">Smartphones</span>
        </a>
    </div>
    <div class="separatCat"></div>
    <div class="classCat cat1">
        <a id="15320" class="activityLink" href="https://www.kompass.com/z/ww/a/office-furniture/15320/">
            <i class="circle distributor" ></i>
            <span class="text">Office furniture</span>
        </a>
    </div>
    <div class="separatCat"></div>
    <div class="classCat cat1">
        <a id="99" class="activityLink" href="https://www.kompass.com/z/ww/a/consulting/99/">
            <i class="circle service" ></i>
            <span class="text">IT consulting</span>
        </a>
    </div>
    <div class="separatCat"></div>
    <div class="classCat cat1">
        <a id="777" class="activityLink" href="https://www.kompass.com/z/ww/a/solar-panels/777/">
            <i class="circle supplier" ></i>
            <i class="circle distributor" ></i>
            <span class="text">Solar panels</span>
        </a>
    </div>
    """

    def test_extracts_all_activity_items_with_roles(self):
        items = kd._extract_activities(self._REAL_ACTIVITY_HTML)
        assert len(items) == 4
        smartphone_item = next(i for i in items if i.text == "Smartphones")
        assert smartphone_item.is_distributor() is True
        assert smartphone_item.classification_id == "4066041"

    def test_service_only_activity_is_not_a_distributor(self):
        items = kd._extract_activities(self._REAL_ACTIVITY_HTML)
        consulting = next(i for i in items if i.text == "IT consulting")
        assert consulting.is_distributor() is False

    def test_supplier_role_is_recognized_as_manufacturer(self):
        items = kd._extract_activities(self._REAL_ACTIVITY_HTML)
        solar = next(i for i in items if i.text == "Solar panels")
        assert solar.is_manufacturer() is True
        assert solar.is_distributor() is True  # has both roles
        assert solar.is_pure_distributor() is False  # but not PURE distributor

    def test_pure_distributor_activity_has_no_manufacturer_role(self):
        items = kd._extract_activities(self._REAL_ACTIVITY_HTML)
        smartphones = next(i for i in items if i.text == "Smartphones")
        assert smartphones.is_manufacturer() is False
        assert smartphones.is_pure_distributor() is True

    def test_manufacturer_distributor_combo_activity_excluded_from_market_match(self):
        """Solar panels activity has BOTH supplier(manufacturer) and
        distributor roles — per the user's explicit instruction, this must
        NOT count as a valid market-specific distributor match, even though
        it technically carries a distributor tag."""
        activities = kd._extract_activities(self._REAL_ACTIVITY_HTML)
        found, reason = kd._has_market_specific_distribution_evidence(
            self._REAL_ACTIVITY_HTML, "Solar Panels", activities=activities,
        )
        assert found is False

    def test_market_with_matching_distributor_activity_is_accepted(self):
        found, reason = kd._has_market_specific_distribution_evidence(
            self._REAL_ACTIVITY_HTML, "Smartphones",
            activities=kd._extract_activities(self._REAL_ACTIVITY_HTML),
        )
        assert found is True
        assert "Smartphones" in reason

    def test_market_with_only_service_activity_is_rejected_even_if_present(self):
        """The company has an 'IT consulting' activity (service, not
        distributor) — a market name matching that text must NOT be
        accepted, since the role isn't distributor."""
        found, reason = kd._has_market_specific_distribution_evidence(
            self._REAL_ACTIVITY_HTML, "Consulting",
            activities=kd._extract_activities(self._REAL_ACTIVITY_HTML),
        )
        assert found is False

    def test_market_with_no_matching_activity_at_all_is_rejected(self):
        found, reason = kd._has_market_specific_distribution_evidence(
            self._REAL_ACTIVITY_HTML, "Solar Panels",
            activities=kd._extract_activities(self._REAL_ACTIVITY_HTML),
        )
        assert found is False

    def test_activity_tree_present_does_not_fall_back_to_prose_matching(self):
        """Once an activity tree is extracted, it is authoritative — even if
        prose elsewhere on the page mentions the market term near a
        distribution word, that must not override an activity-tree miss."""
        html_with_misleading_prose = (
            self._REAL_ACTIVITY_HTML
            + "<p>We are a leading solar panels distributor in the region.</p>"
        )
        found, reason = kd._has_market_specific_distribution_evidence(
            html_with_misleading_prose, "Solar Panels",
            activities=kd._extract_activities(html_with_misleading_prose),
        )
        assert found is False


class TestOrganizationJsonLd:
    def test_extracts_matching_named_organization_among_multiple_blocks(self):
        """A page can carry more than one Organization JSON-LD block (e.g. a
        sidebar/related company) — the one matching the profile we actually
        requested must be selected, never a random/first one."""
        html = """
        <script type="application/ld+json">
        {"@type": "Organization", "name": "Some Other Company", "foundingDate": "2010"}
        </script>
        <script type="application/ld+json">
        {"@type": "Organization", "name": "Accent Electronic SA", "foundingDate": "1996",
         "numberOfEmployees": {"value": 83}}
        </script>
        """
        org = kd._extract_organization_json_ld(html, expected_name="Accent Electronic SA")
        assert org is not None
        assert org["foundingDate"] == "1996"

    def test_founded_and_employees_extracted_from_matched_org(self):
        html = """
        <script type="application/ld+json">
        {"@type": "Organization", "name": "Accent Electronic SA", "foundingDate": "1996",
         "numberOfEmployees": {"value": 83}}
        </script>
        """
        org = kd._extract_organization_json_ld(html, expected_name="Accent Electronic SA")
        assert kd._extract_founded_year(html, org) == "1996"
        assert kd._extract_employees(html, org) == "83"

    def test_same_as_website_preferred_over_body_urls(self):
        html = """
        <script type="application/ld+json">
        {"@type": "Organization", "name": "Accent Electronic SA",
         "sameAs": ["https://www.accentgroup.md", "https://ecoflow.distributor.md"]}
        </script>
        <p>See also https://unrelated-partner-site.example.com</p>
        """
        org = kd._extract_organization_json_ld(html, expected_name="Accent Electronic SA")
        website = kd._extract_official_website(html, "accent-electronic-sa", org)
        assert website == "https://www.accentgroup.md"


class TestSlimRowExport:
    def test_row_has_all_slim_columns_and_kompass_metadata_in_summary(self):
        d = kd.VerifiedDistributor(
            company_name="Accent Electronic SA",
            market="Smartphones",
            kompass_company_id="MD001143",
            kompass_profile_url="https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/",
            kompass_product_id="4066041",
            kompass_company_type="Manufacturer Distributor",
            is_actual_market_distributor=True,
            verification_confidence=94,
            founded_year="1996",
            employees="83",
            website="https://www.accentgroup.md",
        )
        row = kd.to_slim_row(d)
        assert set(row.keys()) == set(kd.SLIM_COLUMNS)
        assert row["Company"] == "Accent Electronic SA"
        assert row["Founded"] == "1996"
        assert row["Employees"] == "83"
        assert "MD001143" in row["Summary"]
        assert "4066041" in row["Summary"]
        assert "94/100" in row["Summary"]


class _FakeCheckpoint:
    """Minimal in-memory stand-in for ExpandCheckpoint (bump/data/save) so
    checkpoint/resume logic can be tested without touching disk."""

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


class TestKompassSearchCheckpointResume:
    """run_kompass_distributor_search(market, ckpt=...) must be able to
    resume a killed run from the last-saved candidate rather than
    restarting the whole market from candidate #1."""

    @pytest.fixture(autouse=True)
    def _use_kompass_profile_verification(self, monkeypatch):
        """These tests exercise the Kompass profile-verification loop, which
        is what the checkpoint tracks. The alternative screening backends
        (ChatGPT / Google AI) skip that loop entirely, so force them off —
        otherwise a developer's real .env enabling either one makes these
        tests exercise the wrong code path and fail."""
        monkeypatch.setenv("KOMPASS_CHATGPT_SCREEN", "false")
        monkeypatch.setenv("KOMPASS_GOOGLE_AI_SCREEN", "false")

    def _candidate(self, i: int) -> kd.KompassCompanyCandidate:
        return kd.KompassCompanyCandidate(
            name=f"Company {i}",
            profile_url=f"https://www.kompass.com/z/ww/c/company-{i}/md{i:06d}/",
            country_hint="",
        )

    def _verified(self, i: int, *, accepted: bool) -> kd.VerifiedDistributor:
        return kd.VerifiedDistributor(
            company_name=f"Company {i}",
            market="Smartphones",
            kompass_company_id=f"MD{i:06d}",
            kompass_profile_url=f"https://www.kompass.com/z/ww/c/company-{i}/md{i:06d}/",
            kompass_product_id="4066041",
            kompass_company_type="Distributor",
            is_actual_market_distributor=accepted,
            verification_confidence=90 if accepted else 10,
            rejection_reason="" if accepted else "no market evidence",
        )

    def test_fresh_run_saves_progress_every_n_candidates(self, monkeypatch):
        candidates = [self._candidate(i) for i in range(12)]
        monkeypatch.setattr(kd, "find_market_product_page", lambda market: ("https://kompass.com/prod/1", "1"))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)
        monkeypatch.setattr(kd, "verify_company", lambda cand, **k: self._verified(int(cand.name.split()[-1]), accepted=True))

        ckpt = _FakeCheckpoint()
        result = kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)

        assert result.market_status == "ok"
        assert len(result.distributors) == 12
        assert len(ckpt.data("kompass_processed_urls")) == 12
        assert ckpt.data("kompass_step_done") is None  # only the backend sets this, not the search fn itself

    def test_resume_skips_already_processed_candidates(self, monkeypatch):
        candidates = [self._candidate(i) for i in range(5)]
        monkeypatch.setattr(kd, "find_market_product_page", lambda market: (_ for _ in ()).throw(
            AssertionError("must not re-run page discovery when checkpoint already has it")
        ))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: (_ for _ in ()).throw(
            AssertionError("must not re-run candidate discovery when checkpoint already has candidates")
        ))

        verify_calls: list[str] = []

        def fake_verify(cand, **k):
            verify_calls.append(cand.name)
            return self._verified(int(cand.name.split()[-1]), accepted=True)

        monkeypatch.setattr(kd, "verify_company", fake_verify)

        prior_state = {
            "kompass_product_url": "https://kompass.com/prod/1",
            "kompass_product_id": "1",
            "kompass_candidates": [
                {"name": c.name, "profile_url": c.profile_url, "country_hint": c.country_hint} for c in candidates
            ],
            "kompass_processed_urls": [candidates[0].profile_url.rstrip("/").lower(), candidates[1].profile_url.rstrip("/").lower()],
            "kompass_accepted": [kd._verified_distributor_to_dict(self._verified(0, accepted=True))],
            "kompass_rejected_count": 1,
            "kompass_unreachable_count": 0,
        }
        ckpt = _FakeCheckpoint(prior_state)

        result = kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)

        # only candidates 2, 3, 4 should have been (re-)verified
        assert verify_calls == ["Company 2", "Company 3", "Company 4"]
        # accepted count: 1 carried over + 3 newly accepted
        assert len(result.distributors) == 4
        assert result.rejected_count == 1

    def test_resume_after_market_page_already_found_only_skips_discovery(self, monkeypatch):
        candidates = [self._candidate(i) for i in range(3)]
        monkeypatch.setattr(kd, "find_market_product_page", lambda market: (_ for _ in ()).throw(
            AssertionError("must not re-run page discovery")
        ))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)
        monkeypatch.setattr(kd, "verify_company", lambda cand, **k: self._verified(int(cand.name.split()[-1]), accepted=True))

        ckpt = _FakeCheckpoint({"kompass_product_url": "https://kompass.com/prod/1", "kompass_product_id": "1"})
        result = kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)

        assert result.market_status == "ok"
        assert len(result.distributors) == 3


class TestDerivedCodes:
    """Country Code and Region Code are computed from Headquarters, not
    researched. to_slim_row() never set them, so both sat at 0/431 in a live
    export even though Headquarters was 100% filled."""

    def _row(self, **kw):
        return kd.to_slim_row(
            kd.VerifiedDistributor(
                company_name="Test Co",
                market="Wearable Medical Devices",
                kompass_company_id="X1",
                kompass_profile_url="https://x/",
                kompass_product_id="1",
                kompass_company_type="Distributor",
                is_actual_market_distributor=True,
                verification_confidence=60,
                **kw,
            )
        )

    def test_codes_derived_from_headquarters(self):
        row = self._row(headquarters="Zaragoza, Spain")
        assert row["Country Code"] == "ES"
        assert row["Region Code"] == "EU"

    def test_codes_derived_from_city_and_country(self):
        row = self._row(city="Trieste", country="IT")
        assert row["Country Code"] == "IT"
        assert row["Region Code"] == "EU"

    def test_no_headquarters_leaves_codes_empty(self):
        """Never invent a code when the location is unknown."""
        row = self._row()
        assert row["Country Code"] == ""
        assert row["Region Code"] == ""

    def test_placeholder_headquarters_yields_no_codes(self):
        row = self._row(headquarters=kd.NOT_DISCLOSED)
        assert row["Country Code"] == ""


class TestCaptchaBlockedNotMarkedProcessed:
    """A company Google refused to answer about must stay queued.

    The checkpoint skips anything in kompass_processed_urls on resume. If a
    CAPTCHA-blocked company were recorded there, it would be dropped from
    the market permanently — on the strength of a bot-block page rather
    than any fact about the company.
    """

    @pytest.fixture(autouse=True)
    def _use_screening_path(self, monkeypatch):
        # Without this the run falls through to the real Kompass
        # profile-fetch path and tries to launch a browser.
        monkeypatch.setenv("KOMPASS_GOOGLE_AI_SCREEN", "true")

    def _candidates(self, n):
        return [
            kd.KompassCompanyCandidate(
                name=f"Company {i}", profile_url=f"https://kompass.com/c/co-{i}/md{i:06d}/"
            )
            for i in range(n)
        ]

    def _run(self, monkeypatch, blocked_idx):
        candidates = self._candidates(10)
        monkeypatch.setattr(kd, "find_market_product_page", lambda m: ("https://k/p/1", "1"))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)

        blocked = {candidates[i].profile_url for i in blocked_idx}

        def fake_screen(chunk, *, market, kompass_product_id):
            hit = {c.profile_url for c in chunk} & blocked
            return [], len(chunk) - len(hit), 0, hit

        monkeypatch.setattr(kd, "_screen_candidates_via_chatgpt", fake_screen)

        ckpt = _FakeCheckpoint()
        kd.run_kompass_distributor_search("Smartphones", ckpt=ckpt)
        return candidates, ckpt

    def test_blocked_company_is_not_recorded_as_processed(self, monkeypatch):
        candidates, ckpt = self._run(monkeypatch, blocked_idx=[3, 7])
        processed = set(ckpt.data("kompass_processed_urls"))
        for i in (3, 7):
            key = candidates[i].profile_url.rstrip("/").lower()
            assert key not in processed, f"blocked company {i} was marked done"

    def test_unblocked_companies_are_still_recorded(self, monkeypatch):
        candidates, ckpt = self._run(monkeypatch, blocked_idx=[3, 7])
        processed = set(ckpt.data("kompass_processed_urls"))
        assert len(processed) == 8
        for i in (0, 1, 2, 4, 5, 6, 8, 9):
            assert candidates[i].profile_url.rstrip("/").lower() in processed

    def test_a_later_run_retries_exactly_the_blocked_ones(self, monkeypatch):
        """The point of not marking them: the next run must pick them up."""
        candidates, ckpt = self._run(monkeypatch, blocked_idx=[3, 7])
        processed = set(ckpt.data("kompass_processed_urls"))
        remaining = [
            c for c in candidates if c.profile_url.rstrip("/").lower() not in processed
        ]
        assert {c.name for c in remaining} == {"Company 3", "Company 7"}

    def test_nothing_blocked_marks_everything_processed(self, monkeypatch):
        candidates, ckpt = self._run(monkeypatch, blocked_idx=[])
        assert len(ckpt.data("kompass_processed_urls")) == 10
