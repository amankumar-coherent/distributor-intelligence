"""classify_company bumps pipeline/llm_meter.py's call counter on every real
classification call. All HTTP is mocked via placeholders.llm._post_json; no
real network/API calls.
"""
from __future__ import annotations

import asyncio

from vendor_intel.clients.claude import ClaudeClient
from vendor_intel.config import Settings
from vendor_intel.intelligence.classifier import classify_company
from vendor_intel.pipeline.llm_meter import get_meter, reset_meter
from vendor_intel.placeholders import llm as llm_module


def test_classify_company_bumps_meter_on_real_usage(monkeypatch):
    monkeypatch.setattr(llm_module, "LLM_PROVIDER", "openai")
    monkeypatch.setenv("OPENAI_API_KEY", "sk-real")

    def fake_post_json(url, headers, body):
        return {
            "choices": [{"message": {"content": (
                '{"company": "Acme", "domain": "acme.com", "is_relevant": true, '
                '"role": "Manufacturer", "country_match": true, "confidence": 0.8}'
            )}}],
            "usage": {"prompt_tokens": 700, "completion_tokens": 120},
        }

    monkeypatch.setattr(llm_module, "_post_json", fake_post_json)

    settings = Settings.load().model_copy(
        update={
            "pipeline_profile": "quality",
            "use_mock_data": False,
            "mock_mode": False,
            "llm_provider": "openai",
        }
    )

    reset_meter()

    company = {"name": "Acme Rupture Discs", "domain": "acme.com"}
    smart_data = {
        "data": {},
        "pages": [{"url": "https://acme.com", "text": "Acme manufactures rupture discs. " * 30}],
    }
    query_context = {"industry": "rupture disc", "country": "Brazil", "sections": []}

    client = ClaudeClient(settings)
    assert client.available, "test setup must actually reach the LLM code path"

    result = asyncio.run(
        classify_company(company, smart_data, query_context, settings=settings, client=client)
    )
    assert result["company"]

    summary = get_meter().summary(provider="openai", model="gpt-4o-mini")
    assert summary["classify_calls"] >= 1
