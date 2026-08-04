"""Post-classification audit (src/vendor_intel/pipeline/post_classification_audit.py)
— keep/review/remove verdicts, generalized from final_audit_v2.py's
per-market MARKET_DESCRIPTIONS dict. All LLM/search calls are mocked.
"""
from __future__ import annotations

import asyncio

import vendor_intel.clients.claude as claude_module
from vendor_intel.clients.search_router import SearchResult
from vendor_intel.pipeline import post_classification_audit as pca


class _FakeRouter:
    async def search(self, query, **kw):
        return [SearchResult(title="Evidence", link="https://x.com", snippet="proof", backend="ddgs")]


class _FailingRouter:
    async def search(self, query, **kw):
        raise RuntimeError("search backend down")


class _FakeClient:
    available = True

    def __init__(self, verdict_by_name: dict[str, dict], *, raise_for: set[str] | None = None):
        self._verdicts = verdict_by_name
        self._raise_for = raise_for or set()
        self.calls: list[str] = []

    def complete_json(self, system, user, max_tokens=200):
        self.calls.append(user)
        for name in self._raise_for:
            if name in user:
                raise RuntimeError("LLM exploded")
        for name, verdict in self._verdicts.items():
            if name in user:
                return verdict
        return {"action": "keep", "reason": "default"}


class _UnavailableClient:
    available = False

    def complete_json(self, *a, **kw):
        raise AssertionError("must not be called when unavailable")


def _patch_client(monkeypatch, client):
    monkeypatch.setattr(claude_module, "ClaudeClient", lambda settings: client)


def test_empty_final_returns_empty():
    kept, removed, reviewed = asyncio.run(
        pca.audit_final_companies([], {}, None, settings=object())
    )
    assert (kept, removed, reviewed) == ([], [], [])


def test_client_unavailable_returns_final_unchanged(monkeypatch):
    _patch_client(monkeypatch, _UnavailableClient())
    final = [{"company": "Acme", "domain": "acme.com"}]
    kept, removed, reviewed = asyncio.run(
        pca.audit_final_companies(final, {}, None, settings=object(), router=_FakeRouter())
    )
    assert kept == final
    assert removed == []
    assert reviewed == []


def test_mixed_verdicts_split_correctly(monkeypatch):
    client = _FakeClient(
        {
            "GoodCo": {"action": "keep", "reason": "confirmed"},
            "BadCo": {"action": "remove", "reason": "wrong industry"},
            "MaybeCo": {"action": "review", "reason": "insufficient evidence"},
        }
    )
    _patch_client(monkeypatch, client)
    final = [
        {"company": "GoodCo", "domain": "good.com", "role": "Manufacturer"},
        {"company": "BadCo", "domain": "bad.com", "role": "Manufacturer"},
        {"company": "MaybeCo", "domain": "maybe.com", "role": "Manufacturer"},
    ]
    kept, removed, reviewed = asyncio.run(
        pca.audit_final_companies(
            final, {"country": "Brazil"}, {"market": "Rupture Discs"}, settings=object(), router=_FakeRouter()
        )
    )
    assert [r["company"] for r in kept] == ["GoodCo"]
    assert len(removed) == 1
    assert removed[0]["company"] == "BadCo"
    assert removed[0]["export_reject"] == "post_audit:wrong industry"
    assert len(reviewed) == 1
    assert reviewed[0] == {
        "company": "MaybeCo",
        "domain": "maybe.com",
        "role": "Manufacturer",
        "reason": "post-audit flagged for review: insufficient evidence",
    }


def test_llm_failure_for_one_company_falls_back_to_keep(monkeypatch):
    client = _FakeClient({}, raise_for={"FlakyCo"})
    _patch_client(monkeypatch, client)
    final = [
        {"company": "FlakyCo", "domain": "flaky.com"},
        {"company": "GoodCo", "domain": "good.com"},
    ]
    kept, removed, reviewed = asyncio.run(
        pca.audit_final_companies(final, {}, None, settings=object(), router=_FakeRouter())
    )
    names = sorted(r["company"] for r in kept)
    assert names == ["FlakyCo", "GoodCo"]
    assert removed == []
    assert reviewed == []


def test_search_failure_degrades_gracefully_not_crash(monkeypatch):
    client = _FakeClient({"Acme": {"action": "keep", "reason": "ok"}})
    _patch_client(monkeypatch, client)
    final = [{"company": "Acme", "domain": "acme.com"}]
    kept, removed, reviewed = asyncio.run(
        pca.audit_final_companies(final, {}, None, settings=object(), router=_FailingRouter())
    )
    assert len(kept) == 1
    assert removed == []
    assert reviewed == []


def test_market_context_passed_into_prompt(monkeypatch):
    client = _FakeClient({"Acme": {"action": "keep", "reason": "ok"}})
    _patch_client(monkeypatch, client)
    final = [{"company": "Acme", "domain": "acme.com"}]
    asyncio.run(
        pca.audit_final_companies(
            final,
            {"industry": "Rupture Disc Market"},
            {"market": "Rupture Disc Market", "market_definition": "pressure relief devices"},
            settings=object(),
            router=_FakeRouter(),
        )
    )
    assert any("Rupture Disc Market" in call for call in client.calls)
    assert any("pressure relief devices" in call for call in client.calls)
