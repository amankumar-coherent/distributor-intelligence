"""Unit tests for website reachability + identity verification, and the
ChatGPT-scraper-based domain correction flow. All network/scraper calls are
mocked — this tests the verification/decision logic itself."""
from __future__ import annotations

from unittest.mock import patch

from vendor_intel.pipeline import website_verification as wv
from vendor_intel.scraping.page_result import PageFetchResult


def _fake_result(**overrides):
    defaults = dict(
        url="https://example.com", final_url="https://example.com", alive=True,
        text="", html="", title="", source="httpx", error="",
    )
    defaults.update(overrides)
    return PageFetchResult(**defaults)


class TestNameAndDomainTokens:
    def test_name_tokens_drops_corporate_suffixes(self):
        tokens = wv._name_tokens("Acme Electronics Inc")
        assert tokens == {"acme", "electronics"}

    def test_name_tokens_handles_multiple_suffix_words(self):
        tokens = wv._name_tokens("Global Plastics Manufacturing Group Ltd")
        assert "group" not in tokens
        assert "ltd" not in tokens
        assert "global" in tokens
        assert "plastics" in tokens

    def test_domain_tokens_from_simple_domain(self):
        assert wv._domain_tokens("https://www.acmeelectronics.com") == {"acmeelectronics"}

    def test_domain_tokens_strips_www(self):
        assert wv._domain_tokens("https://www.acme.com") == {"acme"}


class TestIdentityMatch:
    def test_matches_glued_domain_via_substring(self):
        """Regression: a domain like acmeelectronics.com (no separator)
        must still match company name "Acme Electronics" even though
        _domain_tokens can't split it into separate words."""
        assert wv._looks_like_identity_match(
            "Acme Electronics", "https://www.acmeelectronics.com", "", ""
        ) is True

    def test_matches_via_page_title(self):
        assert wv._looks_like_identity_match(
            "Acme Electronics", "https://unrelated-domain.com",
            "Welcome to Acme Electronics — Official Site", "",
        ) is True

    def test_matches_via_page_text_majority_tokens(self):
        assert wv._looks_like_identity_match(
            "Global Plastics Manufacturing", "https://unrelated.com", "",
            "We are Global Plastics, a leading manufacturing company.",
        ) is True

    def test_rejects_unrelated_company_same_generic_word(self):
        """Two unrelated companies both named '... Inc' or sharing one
        generic word must NOT be treated as a match."""
        assert wv._looks_like_identity_match(
            "Acme Software Solutions Inc",
            "https://totally-different-corp.com",
            "TotallyDifferent Corp — Home",
            "TotallyDifferent Corp provides cloud services worldwide.",
        ) is False

    def test_no_name_tokens_defaults_to_match(self):
        """A company name with only generic/stopword tokens shouldn't block
        on identity — nothing meaningful to check against."""
        assert wv._looks_like_identity_match("Inc Ltd Co", "https://example.com", "", "") is True


class TestCheckWebsiteIdentity:
    def test_empty_website_is_not_reachable(self):
        reachable, identity_ok, reason = wv.check_website_identity("Acme Inc", "")
        assert reachable is False
        assert identity_ok is False

    def test_not_publicly_disclosed_treated_as_no_website(self):
        reachable, identity_ok, reason = wv.check_website_identity("Acme Inc", "Not publicly disclosed")
        assert reachable is False

    def test_unreachable_site_reported(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx",
            return_value=_fake_result(alive=False, error="HTTP 404"),
        ):
            reachable, identity_ok, reason = wv.check_website_identity("Acme Inc", "https://acme.com")
        assert reachable is False
        assert identity_ok is False
        assert "404" in reason

    def test_reachable_but_wrong_company(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx",
            return_value=_fake_result(
                alive=True, title="Unrelated Corp", text="We sell unrelated widgets and gadgets online.",
                final_url="https://totallydifferentco.com",
            ),
        ):
            reachable, identity_ok, reason = wv.check_website_identity(
                "Acme Electronics", "https://totallydifferentco.com"
            )
        assert reachable is True
        assert identity_ok is False

    def test_reachable_and_matching(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx",
            return_value=_fake_result(
                alive=True, title="Acme Electronics — Home", text="Acme Electronics builds things.",
                final_url="https://acmeelectronics.com",
            ),
        ):
            reachable, identity_ok, reason = wv.check_website_identity(
                "Acme Electronics", "https://acmeelectronics.com"
            )
        assert reachable is True
        assert identity_ok is True

    def test_fetch_exception_treated_as_unreachable_never_raises(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx", side_effect=RuntimeError("dns failure")
        ):
            reachable, identity_ok, reason = wv.check_website_identity("Acme Inc", "https://acme.com")
        assert reachable is False
        assert "dns failure" in reason


class TestSslCertificateFallback:
    """Regression: a real, live site with a TLS cert misconfiguration must
    not be treated the same as a genuinely dead domain — this was observed
    live against a real company's real website during integration testing."""

    def test_looks_like_ssl_error_detects_cert_failures(self):
        assert wv._looks_like_ssl_error(
            "ConnectError: [SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: Hostname mismatch"
        ) is True

    def test_looks_like_ssl_error_ignores_unrelated_errors(self):
        assert wv._looks_like_ssl_error("ConnectError: [Errno 11001] getaddrinfo failed") is False
        assert wv._looks_like_ssl_error("HTTP 404") is False

    def test_ssl_error_retried_over_unverified_connection_and_succeeds(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx",
            return_value=_fake_result(alive=False, error="SSL: CERTIFICATE_VERIFY_FAILED: Hostname mismatch"),
        ), patch.object(
            wv, "_fetch_ignoring_ssl",
            return_value=_fake_result(
                alive=True, title="Acme Electronics", text="Acme Electronics official site.",
                final_url="https://acmeelectronics.com",
            ),
        ):
            reachable, identity_ok, reason = wv.check_website_identity(
                "Acme Electronics", "https://acmeelectronics.com"
            )
        assert reachable is True
        assert identity_ok is True

    def test_ssl_error_that_also_fails_unverified_is_genuinely_dead(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx",
            return_value=_fake_result(alive=False, error="SSL: CERTIFICATE_VERIFY_FAILED"),
        ), patch.object(wv, "_fetch_ignoring_ssl", return_value=None):
            reachable, identity_ok, reason = wv.check_website_identity("Acme Inc", "https://acme.com")
        assert reachable is False

    def test_non_ssl_failure_never_triggers_unverified_retry(self):
        with patch(
            "vendor_intel.scraping.httpx_fetch.fetch_page_httpx",
            return_value=_fake_result(alive=False, error="HTTP 404"),
        ), patch.object(wv, "_fetch_ignoring_ssl") as mock_retry:
            reachable, identity_ok, reason = wv.check_website_identity("Acme Inc", "https://acme.com")
        mock_retry.assert_not_called()
        assert reachable is False


class TestVerifyAndFixWebsite:
    def test_disabled_returns_unchanged(self):
        with patch.dict("os.environ", {"WEBSITE_VERIFICATION_ENABLED": "false"}):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://acme.com")
        assert result.source == "unchanged"
        assert result.website == "https://acme.com"

    def test_already_good_website_kept_unchanged(self):
        with patch.object(wv, "check_website_identity", return_value=(True, True, "")):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://acme.com")
        assert result.was_corrected is False
        assert result.website == "https://acme.com"
        assert result.is_reachable is True
        assert result.is_identity_match is True

    def test_dead_website_gets_corrected_via_chatgpt_scraper(self):
        call_count = {"n": 0}

        def fake_check(name, url):
            call_count["n"] += 1
            if url == "https://dead-old-domain.com":
                return False, False, "unreachable"
            return True, True, ""

        with patch.object(wv, "check_website_identity", side_effect=fake_check), patch.object(
            wv, "_find_website_via_chatgpt_scraper", return_value="https://real-acme.com"
        ):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://dead-old-domain.com")

        assert result.was_corrected is True
        assert result.website == "https://real-acme.com"
        assert result.source == "chatgpt_scraper"

    def test_wrong_company_website_gets_corrected(self):
        """Reachable but belongs to a different same-named company — must
        still trigger correction, not be accepted just because it's alive."""
        def fake_check(name, url):
            if url == "https://wrong-acme.com":
                return True, False, "reachable but does not appear to belong to this company"
            return True, True, ""

        with patch.object(wv, "check_website_identity", side_effect=fake_check), patch.object(
            wv, "_find_website_via_chatgpt_scraper", return_value="https://correct-acme.com"
        ):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://wrong-acme.com")

        assert result.was_corrected is True
        assert result.website == "https://correct-acme.com"

    def test_no_website_at_all_searches_from_scratch(self):
        with patch.object(wv, "check_website_identity", return_value=(True, True, "")), patch.object(
            wv, "_find_website_via_chatgpt_scraper", return_value="https://found-acme.com"
        ):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "")
        assert result.was_corrected is True
        assert result.website == "https://found-acme.com"

    def test_no_replacement_found_and_original_dead_returns_empty(self):
        with patch.object(wv, "check_website_identity", return_value=(False, False, "unreachable")), patch.object(
            wv, "_find_website_via_chatgpt_scraper", return_value=""
        ):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://dead.com")
        assert result.website == ""
        assert result.source == "none_found"

    def test_never_reoffers_the_same_dead_candidate_twice(self):
        """If ChatGPT scraper keeps proposing the same URL that already
        failed, must not loop forever re-checking it."""
        attempts = {"n": 0}

        def fake_check(name, url):
            attempts["n"] += 1
            return False, False, "unreachable"

        with patch.object(wv, "check_website_identity", side_effect=fake_check), patch.object(
            wv, "_find_website_via_chatgpt_scraper", return_value="https://dead.com"
        ):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://dead.com", max_attempts=3)

        assert result.website == ""
        # Only checked the original once + first (duplicate) proposal broke the loop early.
        assert attempts["n"] <= 2

    def test_chatgpt_scraper_disabled_falls_back_gracefully(self):
        from vendor_intel.scraping import chatgpt_scraper as cs

        with patch.object(cs, "CHATGPT_SCRAPER_ENABLED", False), patch.object(
            wv, "check_website_identity", return_value=(False, False, "unreachable")
        ):
            result = wv.verify_and_fix_website("Acme Inc", "Ohio, USA", "https://dead.com")
        assert result.website == ""
        assert result.source == "none_found"


class TestFindWebsiteViaChatgptScraper:
    def test_uses_headquarters_to_disambiguate(self):
        captured = {}

        def fake_ask(prompt, timeout=90):
            captured["prompt"] = prompt
            return '{"website": "https://real-acme.com", "confidence": 0.9}'

        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", side_effect=fake_ask
        ):
            result = wv._find_website_via_chatgpt_scraper("Acme Inc", "Columbus, Ohio, USA")

        assert result == "https://real-acme.com"
        assert "Columbus, Ohio, USA" in captured["prompt"]
        assert "Acme Inc" in captured["prompt"]

    def test_returns_empty_on_scraper_error(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", side_effect=RuntimeError("scraper down")
        ):
            result = wv._find_website_via_chatgpt_scraper("Acme Inc", "Ohio")
        assert result == ""

    def test_returns_empty_on_unparseable_response(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", return_value="I don't know."
        ):
            result = wv._find_website_via_chatgpt_scraper("Acme Inc", "Ohio")
        assert result == ""
