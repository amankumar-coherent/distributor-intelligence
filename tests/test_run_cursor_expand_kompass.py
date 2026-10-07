"""Tests for run_cursor_expand.py's --backend kompass dispatch and its
automatic fallback to the ChatGPT/DeepSeek expand pipeline when Kompass
alone doesn't reach MIN_COMPANIES_PER_MARKET."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import run_cursor_expand as rce  # noqa: E402


def _make_args(**overrides) -> argparse.Namespace:
    base = dict(
        query="Smartphones",
        country="global",
        batch="all",
        target=500,
        final=None,
        out_dir=None,
        no_merge=False,
        enrich=False,
        max_fill=500,
        fill_existing=False,
        cap="maximum",
        max_queries=120,
        seeds_only=False,
        seeds=None,
        skip_recall=False,
        no_web_fill=False,
        no_openai_discover=False,
        ddgs_harvest=False,
        fresh=False,
        resume=False,
        no_linkedin_enrich=False,
        no_apollo=False,
        fill_backend="search",
        model="composer-2.5",
        no_stream=False,
    )
    base.update(overrides)
    return argparse.Namespace(**base)


def _fake_result(status="ok", n_distributors=0, rejected=0, company_type="Distributor"):
    from vendor_intel.discovery.kompass_distributors import MarketDistributorResult, VerifiedDistributor

    distributors = [
        VerifiedDistributor(
            company_name=f"Company {i}",
            market="Smartphones",
            kompass_company_id=f"MD{i:06d}",
            kompass_profile_url=f"https://www.kompass.com/z/ww/c/company-{i}/md{i:06d}/",
            kompass_product_id="4066041",
            kompass_company_type=company_type,
            is_actual_market_distributor=True,
            verification_confidence=90,
        )
        for i in range(n_distributors)
    ]
    return MarketDistributorResult(
        market="Smartphones",
        market_status=status,
        kompass_product_id="4066041" if status == "ok" else "",
        kompass_product_url="https://www.kompass.com/z/ww/a/smartphones/4066041/" if status == "ok" else "",
        distributors=distributors,
        rejected_count=rejected,
    )


class TestKompassBackendFallback:
    def test_no_fallback_when_kompass_alone_meets_minimum(self, tmp_path, monkeypatch):
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "5")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=10),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ) as mock_write, patch.object(
            rce, "run_chatgpt_backend"
        ) as mock_fallback:
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        mock_fallback.assert_not_called()
        # Initial Kompass write + the final gate's write (which now always
        # runs, since it also applies website verification, not just the
        # manufacturer/OEM filter).
        assert mock_write.call_count == 2

    def test_fallback_triggered_when_kompass_below_minimum(self, tmp_path, monkeypatch):
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "500")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=3),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(
            rce, "run_chatgpt_backend", return_value=0
        ) as mock_fallback:
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        mock_fallback.assert_called_once()
        fallback_args = mock_fallback.call_args[0][0]
        assert fallback_args.target == 500
        assert fallback_args.no_merge is False  # must merge on top of Kompass rows

    def test_fallback_triggered_when_kompass_market_not_found(self, tmp_path, monkeypatch):
        """Even a total Kompass miss (market not found on Kompass at all)
        must still fall through to the expand pipeline, not just exit."""
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "500")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="kompass_market_not_found"),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(
            rce, "run_chatgpt_backend", return_value=0
        ) as mock_fallback:
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        mock_fallback.assert_called_once()

    def test_final_gate_rejects_manufacturer_rows_even_without_fallback(self, tmp_path, monkeypatch):
        """A manufacturer/OEM row that somehow slipped through Kompass's own
        verify_and_score() must still be caught by the final gate before the
        no-fallback-needed early return."""
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "1")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=2, company_type="Manufacturer"),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ) as mock_write, patch.object(rce, "run_chatgpt_backend") as mock_fallback:
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        # Initial Kompass write + a second write after the gate strips the
        # manufacturer rows out.
        assert mock_write.call_count == 2
        final_call_rows = mock_write.call_args_list[-1].args[1]
        assert final_call_rows == []  # both fake rows were Manufacturer -> rejected

    def test_final_gate_applied_after_fallback_top_up(self, tmp_path, monkeypatch):
        """The gate must also run on the COMBINED row set after the ChatGPT/
        DeepSeek fallback tops up — not just on the Kompass-only rows."""
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "500")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"
        post_fallback_rows = [
            {"Company": "Good Co", "Summary": "Company Type: Distributor"},
            {"Company": "Bad Manufacturer Co", "Summary": "Company Type: Manufacturer"},
        ]

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=1),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=post_fallback_rows
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ) as mock_write, patch.object(rce, "run_chatgpt_backend", return_value=0) as mock_fallback:
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        mock_fallback.assert_called_once()
        final_call_rows = mock_write.call_args_list[-1].args[1]
        final_names = {r["Company"] for r in final_call_rows}
        assert final_names == {"Good Co"}
        assert "Bad Manufacturer Co" not in final_names

    def test_website_verification_fixes_dead_domain(self, tmp_path, monkeypatch):
        from vendor_intel.pipeline.website_verification import WebsiteVerificationResult

        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "1")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        result = _fake_result(status="ok", n_distributors=1)
        result.distributors[0].website = "https://dead-domain.com"

        fixed = WebsiteVerificationResult(
            website="https://real-company-0.com", is_reachable=True, is_identity_match=True,
            was_corrected=True, source="chatgpt_scraper",
        )

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=result,
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ) as mock_write, patch(
            "vendor_intel.pipeline.website_verification.verify_and_fix_website", return_value=fixed
        ) as mock_verify, patch.object(rce, "run_chatgpt_backend"):
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        mock_verify.assert_called()
        final_call_rows = mock_write.call_args_list[-1].args[1]
        assert final_call_rows[0]["Website"] == "https://real-company-0.com"

    def test_existing_final_rows_counted_toward_minimum(self, tmp_path, monkeypatch):
        """Rows already in the FINAL Excel (from a prior run) count toward
        the minimum — Kompass doesn't need to rediscover them."""
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "5")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"
        existing_rows = [{"Company": f"Existing {i}"} for i in range(5)]

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=0),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=existing_rows
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(
            rce, "run_chatgpt_backend"
        ) as mock_fallback:
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        mock_fallback.assert_not_called()


class TestKompassBackendCheckpointResume:
    """run_kompass_backend must persist a checkpoint under the kompass
    batch id and, on a resumed invocation where the search already
    completed, skip re-calling run_kompass_distributor_search entirely."""

    def test_completed_run_writes_done_checkpoint(self, tmp_path, monkeypatch):
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "5")
        args = _make_args()
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=10),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(rce, "run_chatgpt_backend"):
            rc = rce.run_kompass_backend(args)

        assert rc == 0
        ckpt_path = final_path.parent / "chatgpt_checkpoint_batch_kompass.json"
        assert ckpt_path.exists()
        import json as _json

        state = _json.loads(ckpt_path.read_text(encoding="utf-8"))
        assert state["status"] == "done"
        assert state["data"]["kompass_step_done"] is True

    def test_resumed_run_skips_kompass_search_when_already_done(self, tmp_path, monkeypatch):
        """On resume, if the checkpoint already recorded a completed Kompass
        search, run_kompass_backend must NOT call
        run_kompass_distributor_search again — it reuses the saved rows."""
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "5")
        args = _make_args(fresh=False)
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=10),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(rce, "run_chatgpt_backend"):
            rc = rce.run_kompass_backend(args)
        assert rc == 0

        # second invocation: run_kompass_distributor_search must not be called at all
        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            side_effect=AssertionError("must not re-run a completed kompass search on resume"),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(rce, "run_chatgpt_backend") as mock_fallback2:
            rc2 = rce.run_kompass_backend(args)

        assert rc2 == 0
        mock_fallback2.assert_not_called()

    def test_fresh_flag_ignores_prior_checkpoint(self, tmp_path, monkeypatch):
        monkeypatch.setenv("MIN_COMPANIES_PER_MARKET", "5")
        args = _make_args(fresh=False)
        final_path = tmp_path / "smartphones_FINAL.xlsx"

        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=10),
        ), patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(rce, "run_chatgpt_backend"):
            rce.run_kompass_backend(args)

        fresh_args = _make_args(fresh=True)
        with patch(
            "vendor_intel.discovery.kompass_distributors.run_kompass_distributor_search",
            return_value=_fake_result(status="ok", n_distributors=10),
        ) as mock_search, patch(
            "vendor_intel.pipeline.web_expand.resolve_final_path", return_value=final_path
        ), patch(
            "vendor_intel.pipeline.web_expand.read_final_rows", return_value=[]
        ), patch(
            "vendor_intel.pipeline.web_expand.write_final_xlsx"
        ), patch.object(rce, "run_chatgpt_backend"):
            rc = rce.run_kompass_backend(fresh_args)

        assert rc == 0
        mock_search.assert_called_once()
