# -*- coding: utf-8 -*-
"""Chinese and Japanese Google answers must be usable.

Kompass's country sweep contributes 326 Chinese and 228 Japanese companies
to this market. Google answers those in the company's own language, and the
English-only matchers found nothing — so every one scored "unknown" and was
dropped, with empty geography columns.

The Chinese text below is the real live AI Overview answer for
中技惠州进出口公司 (a China National Technical Import/Export company).
"""
from __future__ import annotations

from vendor_intel.discovery.chatgpt_company_screen import (
    _country_names_in,
    _verdict_from_overview,
)

LIVE_CN_ANSWER = (
    "中技惠州进出口公司通常是指国家级央企中国技术进出口集团有限公司，"
    "其服务的国家和地区遍布全球 105 个国家和地区。"
    "主要供应与服务国家：东南亚与南亚（核心市场）：包括印度尼西亚、马来西亚、"
    "越南、菲律宾、孟加拉国、巴基斯坦等。中亚地区：包括土库曼斯坦、哈萨克斯坦、"
    "乌兹别克斯坦等国。中东与非洲：覆盖伊拉克等中东国家，以及科特迪瓦、赞比亚等。"
)


class TestCountriesFromChinese:
    def test_live_answer_yields_the_named_countries(self):
        got = _country_names_in(LIVE_CN_ANSWER, ignore="中技惠州进出口公司")
        for expected in ("Indonesia", "Malaysia", "Vietnam", "Philippines",
                         "Bangladesh", "Pakistan", "Kazakhstan", "Iraq"):
            assert expected in got, f"{expected} missing from {got}"

    def test_values_are_canonical_english(self):
        """The column must read the same whatever language Google used."""
        got = _country_names_in("公司出口到德国和日本。")
        assert got == ["Germany", "Japan"]

    def test_longer_name_wins_over_its_prefix(self):
        """印度尼西亚 (Indonesia) contains 印度 (India) — matching the prefix
        would record a country the company was never linked to."""
        assert _country_names_in("印度尼西亚和马来西亚") == ["Indonesia", "Malaysia"]

    def test_india_still_matches_on_its_own(self):
        assert _country_names_in("印度和中国") == ["India", "China"]

    def test_japanese_names(self):
        got = _country_names_in("ドイツとフランスとイタリアに輸出しています。")
        assert got == ["Germany", "France", "Italy"]

    def test_english_answers_are_unaffected(self):
        assert _country_names_in("It supplies Italy and Spain.") == ["Italy", "Spain"]

    def test_no_countries_gives_empty(self):
        assert _country_names_in("この会社は医療機器を販売しています。") == []


class TestVerdictFromCJK:
    def test_live_chinese_answer_is_a_distributor(self):
        assert _verdict_from_overview(LIVE_CN_ANSWER, market="Wearable Medical Devices")[0] == (
            "distributor"
        )

    def test_chinese_manufacturer(self):
        assert _verdict_from_overview("该公司是一家医疗器械制造商。", market="M")[0] == "manufacturer"

    def test_japanese_distributor(self):
        assert _verdict_from_overview("この会社は医療機器の代理店です。", market="M")[0] == "distributor"

    def test_japanese_manufacturer(self):
        assert _verdict_from_overview("この会社は医療機器メーカーです。", market="M")[0] == "manufacturer"

    def test_manufacturer_wins_when_both_roles_appear(self):
        """Matches the English rule: a company that makes AND resells is a
        manufacturer for our purposes."""
        assert _verdict_from_overview("该公司是制造商和经销商。", market="M")[0] == "manufacturer"

    def test_traditional_characters(self):
        assert _verdict_from_overview("該公司是經銷商。", market="M")[0] == "distributor"

    def test_korean_distributor(self):
        assert _verdict_from_overview("이 회사는 의료기기 유통업체입니다.", market="M")[0] == (
            "distributor"
        )

    def test_cjk_text_with_no_role_word_stays_unknown(self):
        """Never guess just because the text is CJK."""
        assert _verdict_from_overview("この会社は東京にあります。", market="M")[0] == "unknown"


class TestNoRegression:
    def test_english_distributor(self):
        assert _verdict_from_overview("Acme is a distributor of devices.", market="M")[0] == (
            "distributor"
        )

    def test_english_manufacturer(self):
        assert _verdict_from_overview("Acme manufactures devices.", market="M")[0] == (
            "manufacturer"
        )

    def test_captcha_page_still_blocked(self):
        assert _verdict_from_overview(
            "Our systems have detected unusual traffic from your computer network.",
            market="M",
        )[0] == "unknown"
