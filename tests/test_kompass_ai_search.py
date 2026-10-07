"""Tests for kompass_ai_search — verifying that AI-proposed Kompass pages/ids
are never trusted without a real, successful fetch (spec: 'if the AI returns
an ID that cannot be verified against a Kompass URL, reject it')."""
from __future__ import annotations

from unittest.mock import patch

from vendor_intel.clients import kompass as kompass_client
from vendor_intel.discovery import kompass_ai_search as kai


def _fake_disk_cache_miss(*_a, **_k):
    return None


class TestUnderstandMarket:
    """New step: ChatGPT scraper defines the market -> DeepSeek distills
    into search terms -> fed into the existing page-discovery prompt."""

    def test_disabled_via_env_returns_empty(self):
        with patch.dict("os.environ", {"KOMPASS_MARKET_UNDERSTANDING_ENABLED": "false"}):
            assert kai.understand_market("Flexible Packaging") == ""

    def test_falls_back_to_empty_when_chatgpt_scraper_disabled(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=False), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ):
            assert kai.understand_market("Flexible Packaging") == ""

    def test_falls_back_to_empty_when_chatgpt_scraper_errors(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", side_effect=RuntimeError("scraper down")
        ), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ):
            assert kai.understand_market("Flexible Packaging") == ""

    def test_full_chain_definition_then_distillation(self):
        definition = (
            "The flexible packaging market covers non-rigid packaging materials "
            "such as plastic films, foils, and pouches that change shape when filled."
        )
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", return_value=definition
        ), patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), patch(
            "vendor_intel.pipeline.chatgpt_expand._client"
        ), patch("vendor_intel.pipeline.chatgpt_expand._model", return_value="deepseek-chat"), patch(
            "vendor_intel.pipeline.chatgpt_expand._chat_json",
            return_value={"search_terms": "flexible plastic films and pouches"},
        ), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ), patch("vendor_intel.clients.search_cache_disk.put_disk_cached_search") as mock_put:
            terms = kai.understand_market("Flexible Packaging")

        assert terms == "flexible plastic films and pouches"
        mock_put.assert_called_once()

    def test_no_deepseek_key_falls_back_to_empty_even_with_definition(self):
        with patch("vendor_intel.scraping.google_ai_mode.generative_backend_enabled", return_value=True), patch(
            "vendor_intel.scraping.google_ai_mode.ask_generative", return_value="Some definition."
        ), patch.dict("os.environ", {}, clear=False), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ):
            import os

            os.environ.pop("OPENAI_API_KEY", None)
            terms = kai.understand_market("Flexible Packaging")
        assert terms == ""

    def test_uses_cache_on_second_call(self):
        with patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            return_value=[{"definition": "cached def", "search_terms": "cached terms"}],
        ), patch("vendor_intel.scraping.google_ai_mode.ask_generative") as mock_ask:
            terms = kai.understand_market("Flexible Packaging")
        assert terms == "cached terms"
        mock_ask.assert_not_called()


class TestDiscoverProductPageUsesUnderstanding:
    def test_search_terms_passed_through_to_ask_provider(self):
        fake_page = kompass_client.KompassPage(
            url="https://www.kompass.com/z/ww/a/flexible-packaging/12345/",
            final_url="https://www.kompass.com/z/ww/a/flexible-packaging/12345/",
            alive=True,
            html="<html>ok</html>",
        )
        captured = {}

        def fake_ask(provider, market, search_terms=""):
            captured["search_terms"] = search_terms
            return {
                "kompass_product_url": "https://www.kompass.com/z/ww/a/flexible-packaging/12345/",
                "confidence": 0.9,
            }

        with patch.object(kai, "understand_market", return_value="flexible plastic films"), patch.object(
            kai, "_provider_chain", return_value=["chatgpt_scraper"]
        ), patch.object(kai, "_ask_provider", side_effect=fake_ask), patch(
            "vendor_intel.clients.kompass.fetch_kompass_page", return_value=fake_page
        ), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ), patch("vendor_intel.clients.search_cache_disk.put_disk_cached_search"):
            result = kai.discover_product_page("Flexible Packaging")

        assert result is not None
        assert captured["search_terms"] == "flexible plastic films"

    def test_ask_provider_includes_search_terms_in_prompt(self):
        from vendor_intel.placeholders import llm as ph

        captured = {}

        def fake_llm_complete(system, user, **kwargs):
            captured["user"] = user
            return '{"kompass_product_url": "", "confidence": 0.0}'

        with patch.object(ph, "is_configured", return_value=True), patch.object(
            ph, "llm_complete", side_effect=fake_llm_complete
        ):
            kai._ask_provider("anthropic", "Flexible Packaging", search_terms="flexible plastic films")

        assert "flexible plastic films" in captured["user"]
        assert "Flexible Packaging" in captured["user"]


class TestDiscoverProductPage:
    """NOTE: every test here patches kai.understand_market directly — it's
    the actual seam discover_product_page() calls before the provider chain,
    and patching it (rather than just _ask_provider) keeps these tests fully
    isolated from the real CHATGPT_SCRAPER_ENABLED/.env state. Without this,
    a real .env with the scraper enabled makes discover_product_page() launch
    an actual Selenium/Chrome session against chatgpt.com during test runs."""

    def test_rejects_ai_url_that_fails_to_fetch(self):
        """AI proposes a plausible-looking URL, but it 404s / isn't fetchable
        — must be rejected, never trusted."""
        fake_page = kompass_client.KompassPage(
            url="https://www.kompass.com/fake/9999999/",
            final_url="https://www.kompass.com/fake/9999999/",
            alive=False,
            error="HTTP 404",
        )
        with patch.object(kai, "understand_market", return_value=""), patch.object(
            kai, "_provider_chain", return_value=["gemini"]
        ), patch.object(
            kai, "_ask_provider",
            return_value={"kompass_product_url": "https://www.kompass.com/fake/9999999/", "confidence": 0.9},
        ), patch("vendor_intel.clients.kompass.fetch_kompass_page", return_value=fake_page), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ), patch("vendor_intel.clients.search_cache_disk.put_disk_cached_search"):
            result = kai.discover_product_page("Smartphones")
        assert result is None

    def test_rejects_non_kompass_url(self):
        """AI hallucinating a non-kompass.com domain must be rejected before
        any fetch is even attempted."""
        with patch.object(kai, "understand_market", return_value=""), patch.object(
            kai, "_provider_chain", return_value=["gemini"]
        ), patch.object(
            kai, "_ask_provider",
            return_value={"kompass_product_url": "https://example.com/smartphones/123/", "confidence": 0.9},
        ), patch("vendor_intel.clients.search_cache_disk.get_disk_cached_search",
                 side_effect=_fake_disk_cache_miss), patch(
            "vendor_intel.clients.search_cache_disk.put_disk_cached_search"
        ):
            result = kai.discover_product_page("Smartphones")
        assert result is None

    def test_accepts_verified_url_and_extracts_id_from_fetched_page_not_ai(self):
        """The product id in the returned candidate must come from
        extract_product_id() on the VERIFIED fetched URL, not be echoed
        blindly from whatever the AI said."""
        fake_page = kompass_client.KompassPage(
            url="https://www.kompass.com/en/b/telecoms/smartphones/4066041/",
            final_url="https://www.kompass.com/en/b/telecoms/smartphones/4066041/",
            alive=True,
            html="<html>ok</html>",
        )
        with patch.object(kai, "understand_market", return_value=""), patch.object(
            kai, "_provider_chain", return_value=["gemini"]
        ), patch.object(
            kai, "_ask_provider",
            return_value={
                "kompass_product_url": "https://www.kompass.com/en/b/telecoms/smartphones/4066041/",
                "confidence": 0.95,
            },
        ), patch("vendor_intel.clients.kompass.fetch_kompass_page", return_value=fake_page), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ), patch("vendor_intel.clients.search_cache_disk.put_disk_cached_search"):
            result = kai.discover_product_page("Smartphones")
        assert result is not None
        assert result.kompass_product_id == "4066041"
        assert result.confidence == 0.95

    def test_falls_through_to_second_provider_when_first_fails(self):
        fake_page = kompass_client.KompassPage(
            url="https://www.kompass.com/en/b/telecoms/smartphones/4066041/",
            final_url="https://www.kompass.com/en/b/telecoms/smartphones/4066041/",
            alive=True,
            html="<html>ok</html>",
        )

        def fake_ask(provider, market, search_terms=""):
            if provider == "gemini":
                return None  # first provider fails outright
            return {
                "kompass_product_url": "https://www.kompass.com/en/b/telecoms/smartphones/4066041/",
                "confidence": 0.8,
            }

        with patch.object(kai, "understand_market", return_value=""), patch.object(
            kai, "_provider_chain", return_value=["gemini", "openai"]
        ), patch.object(
            kai, "_ask_provider", side_effect=fake_ask
        ), patch("vendor_intel.clients.kompass.fetch_kompass_page", return_value=fake_page), patch(
            "vendor_intel.clients.search_cache_disk.get_disk_cached_search",
            side_effect=_fake_disk_cache_miss,
        ), patch("vendor_intel.clients.search_cache_disk.put_disk_cached_search"):
            result = kai.discover_product_page("Smartphones")
        assert result is not None
        assert result.provider_used == "openai"

    def test_deepseek_alias_normalizes_to_openai_dispatch(self):
        assert kai.ai_search_provider  # sanity import
        with patch.dict("os.environ", {"AI_SEARCH_PROVIDER": "deepseek", "AI_FALLBACK_PROVIDER": ""}):
            chain = kai._provider_chain()
        assert chain == ["openai"]
