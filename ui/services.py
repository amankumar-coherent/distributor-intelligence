"""Pipeline and data services for the Streamlit UI."""
from __future__ import annotations

import io
import json
import re
import sys
import threading
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from ui.bootstrap import ROOT, init_env, output_dir, phase1_debug_dir


def slug(query: str, country: str) -> str:
    text = f"{query}_{country}".lower()
    return "".join(c if c.isalnum() else "_" for c in text).strip("_")[:80]


def parse_query_line(raw: str) -> tuple[str, str]:
    if "|" in raw:
        parts = raw.split("|", 1)
        return parts[0].strip(), (parts[1].strip() or "global")
    return raw.strip(), "global"


def load_queries_file(path: Path) -> list[tuple[str, str]]:
    if not path.exists():
        return []
    rows: list[tuple[str, str]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        rows.append(parse_query_line(line))
    return rows


_SKIP_RESULT_CSV = frozenset({"session_log.csv"})


def _is_pipeline_export_csv(path: Path, out: Path) -> bool:
    """True for run_query exports: demo/<slug>.csv or demo/<slug>/<slug>.csv."""
    if path.name.startswith("_") or path.name in _SKIP_RESULT_CSV:
        return False
    if path.parent == out:
        return True
    try:
        rel = path.relative_to(out)
    except ValueError:
        return False
    parts = rel.parts
    if len(parts) == 2 and parts[1].endswith(".csv") and Path(parts[1]).stem == parts[0]:
        return True
    json_path = path.with_suffix(".json")
    if not json_path.is_file():
        return False
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception:
        return False
    return isinstance(data.get("relevant_companies"), list)


def list_result_csvs() -> list[Path]:
    """Pipeline CSV exports under MARKET_QUERY_OUTPUT_DIR (flat + market subfolders)."""
    out = output_dir()
    if not out.is_dir():
        return []
    seen: set[str] = set()
    found: list[Path] = []
    for path in out.rglob("*.csv"):
        if not _is_pipeline_export_csv(path, out):
            continue
        key = str(path.resolve())
        if key in seen:
            continue
        seen.add(key)
        found.append(path)
    found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return found


def result_csv_label(path: Path) -> str:
    """Display key for Streamlit selectbox (relative to output dir)."""
    out = output_dir()
    try:
        return path.relative_to(out).as_posix()
    except ValueError:
        return path.name


def find_market_artifacts(csv_path: Path) -> dict[str, Path | None]:
    """Sibling JSON / Excel files for a pipeline CSV export."""
    out = output_dir()
    market_dir = csv_path.parent
    stem = csv_path.stem
    json_path = csv_path.with_suffix(".json")
    artifacts: dict[str, Path | None] = {
        "csv": csv_path,
        "json": json_path if json_path.is_file() else None,
        "fmcg_xlsx": None,
        "founded_hq_xlsx": None,
    }
    for base in (market_dir, out):
        if not base.is_dir():
            continue
        for p in sorted(base.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
            name = p.name.lower()
            if "founded_and_hq" in name and stem.replace("_", "") in name.replace("_", ""):
                artifacts["founded_hq_xlsx"] = p
            elif "founded_and_hq" not in name and stem in p.stem:
                artifacts["fmcg_xlsx"] = p
    if artifacts["founded_hq_xlsx"] is None:
        for p in market_dir.glob("*founded_and_hq*.xlsx"):
            artifacts["founded_hq_xlsx"] = p
            break
    return artifacts


def _dataframe_from_pipeline_json(json_path: Path, *, market_type: str = "fmcg") -> "Any":
    import pandas as pd

    from vendor_intel.export.fmcg_landscape import _export_columns, row_from_verdict

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows_raw = data.get("relevant_companies") or []
    rows = [row_from_verdict(r) for r in rows_raw if r.get("is_relevant", True)]
    cols = _export_columns(market_type)
    return pd.DataFrame([{c: row.get(c, "") for c in cols} for row in rows])


def load_results_dataframe(
    csv_path: Path,
    view: str = "landscape",
) -> tuple[Any, str, dict[str, Path | None]]:
    """Load export table: pipeline CSV, full FMCG landscape, or Founded+HQ Excel."""
    import pandas as pd

    artifacts = find_market_artifacts(csv_path)
    if view == "csv":
        return pd.read_csv(csv_path), "Pipeline CSV", artifacts
    if view == "founded_hq" and artifacts.get("founded_hq_xlsx"):
        xlsx = artifacts["founded_hq_xlsx"]
        assert xlsx is not None
        return pd.read_excel(xlsx, header=2), "Founded + HQ Excel", artifacts
    if view == "landscape" and artifacts.get("fmcg_xlsx"):
        xlsx = artifacts["fmcg_xlsx"]
        assert xlsx is not None
        return pd.read_excel(xlsx, header=2), "FMCG landscape Excel", artifacts
    if artifacts.get("json"):
        return (
            _dataframe_from_pipeline_json(artifacts["json"]),
            "Full landscape (from JSON — 30+ columns)",
            artifacts,
        )
    return pd.read_csv(csv_path), "Pipeline CSV (fallback)", artifacts


def get_integration_status() -> list[dict[str, str]]:
    """Agent Reach, LinkedIn MCP, search stack — for Streamlit health panels."""
    init_env()
    rows: list[dict[str, str]] = []
    try:
        from vendor_intel.integrations.agent_reach_client import (
            agent_reach_enabled,
            exa_search_enabled,
        )

        rows.append(
            {
                "Integration": "Agent Reach (Jina)",
                "Enabled": "yes" if agent_reach_enabled() else "no",
                "Detail": "Web page read for enrichment",
            }
        )
        rows.append(
            {
                "Integration": "Exa search",
                "Enabled": "yes" if exa_search_enabled() else "no",
                "Detail": "Semantic search fallback",
            }
        )
    except Exception as exc:
        rows.append({"Integration": "Agent Reach", "Enabled": "error", "Detail": str(exc)[:120]})

    try:
        from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_available, linkedin_mcp_status

        li = linkedin_mcp_status()
        rows.append(
            {
                "Integration": "LinkedIn MCP",
                "Enabled": "yes" if linkedin_mcp_available() else "no",
                "Detail": str(li.get("detail") or li.get("url") or li)[:120],
            }
        )
    except Exception as exc:
        rows.append({"Integration": "LinkedIn MCP", "Enabled": "error", "Detail": str(exc)[:120]})

    import os

    rows.append(
        {
            "Integration": "Wikidata",
            "Enabled": "yes" if os.getenv("WIKIDATA_ENABLED", "true").lower() in ("1", "true", "yes") else "no",
            "Detail": "Free SPARQL firmographics in gap-fill",
        }
    )
    rows.append(
        {
            "Integration": "SearXNG",
            "Enabled": "yes" if os.getenv("SEARXNG_BASE_URL", "").strip() else "no",
            "Detail": os.getenv("SEARXNG_BASE_URL", "not configured")[:120],
        }
    )
    rows.append(
        {
            "Integration": "Apollo contacts",
            "Enabled": "yes" if os.getenv("APOLLO_API_KEY", "").strip() else "no",
            "Detail": "Post-enrich step 3 — contact names",
        }
    )
    rows.append(
        {
            "Integration": "Neon DB",
            "Enabled": "yes" if os.getenv("DATABASE_URL", "").strip() else "no",
            "Detail": "Post-enrich step 8 — push export",
        }
    )
    return rows


def load_session_log() -> list[dict]:
    log_path = output_dir() / "session_log.json"
    if not log_path.exists():
        return []
    try:
        return json.loads(log_path.read_text(encoding="utf-8"))
    except Exception:
        return []


def append_session_log(entry: dict) -> None:
    log = load_session_log()
    log.append(entry)
    out = output_dir()
    out.mkdir(parents=True, exist_ok=True)
    (out / "session_log.json").write_text(
        json.dumps(log, indent=2, default=str), encoding="utf-8"
    )


def _safe_stdout_write(stream: Any, data: str) -> None:
    """Windows consoles often use cp1252 — avoid crashing on ≤, →, etc."""
    try:
        stream.write(data)
    except UnicodeEncodeError:
        enc = getattr(stream, "encoding", None) or "utf-8"
        stream.write(data.encode(enc, errors="replace").decode(enc, errors="replace"))


@contextmanager
def capture_stdout(callback: Callable[[str], None]):
    buffer = io.StringIO()
    old_stdout = sys.stdout

    class Tee:
        def write(self, data: str) -> int:
            _safe_stdout_write(old_stdout, data)
            buffer.write(data)
            callback(data)
            return len(data)

        def flush(self) -> None:
            old_stdout.flush()

    sys.stdout = Tee()
    try:
        yield buffer
    finally:
        sys.stdout = old_stdout


def run_full_pipeline(
    query: str,
    country: str,
    profile: str,
    *,
    cap: str | None = None,
    pipeline_mode: str = "quality",
    post_enrich: bool = True,
    fmcg_xlsx: bool = True,
    push_neon: bool = False,
    no_push_neon: bool = False,
    skip_name_resolve: bool = False,
    log_callback: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    from run_query import _parse_input, _run_post_enrich, _should_push_to_neon, run_one_query

    from ui.bootstrap import load_settings, pipeline_caps

    init_env()
    settings = load_settings(profile)
    if cap:
        from vendor_intel.pipeline.cap_profiles import apply_cap

        settings = apply_cap(settings, cap)

    from vendor_intel.pipeline.pipeline_modes import apply_mode_settings, resolve_batch_options

    batch_opts = resolve_batch_options(
        pipeline_mode if pipeline_mode in ("bulk", "standard", "quality", "full") else "quality",
        max_concurrent=None,
        cap=cap,
        profile=profile,
        fmcg_xlsx=fmcg_xlsx,
        full_enrich=False,
        with_audit=False,
        gap_fill=False,
        push_neon=push_neon,
        search_backfill=False,
    )
    if pipeline_mode in ("quality", "full"):
        settings = apply_mode_settings(settings, pipeline_mode, cap_override=batch_opts["cap"])
    do_fmcg = fmcg_xlsx or batch_opts.get("fmcg_xlsx", False)
    do_post = post_enrich and batch_opts.get("post_market") is not None
    post_market_kwargs = dict(batch_opts.get("post_market") or {})
    if skip_name_resolve:
        post_market_kwargs["resolve_partials"] = False

    # Let users type "Market | Country | Section A; Section B" in the query box.
    market, q_country, sections = _parse_input(query)
    use_country = (
        country.strip()
        if country and country.strip() and country.strip().lower() != "global"
        else q_country
    )
    classify_cap, enrich_cap = pipeline_caps(settings, country=use_country)

    def _noop(_: str) -> None:
        pass

    cb = log_callback or _noop
    with capture_stdout(cb):
        try:
            from vendor_intel.integrations.agent_reach_client import log_agent_reach_status

            log_agent_reach_status()
        except Exception:
            pass
        if do_post:
            print(
                "  Post-enrich: ON (gap-fill → LinkedIn MCP → Apollo → brand → audit → Excel → Neon)",
                flush=True,
            )
        result = run_one_query(
            market,
            use_country,
            settings,
            enrich_limit=enrich_cap,
            classify_limit=classify_cap,
            sections=sections,
            fmcg_xlsx=do_fmcg,
        )
        if do_post:
            post_stats = _run_post_enrich(
                result,
                push_neon=push_neon,
                no_push_neon=no_push_neon,
                post_market_kwargs=post_market_kwargs,
            )
            if post_stats:
                result["post_market_pipeline"] = post_stats
        elif _should_push_to_neon(push_neon=push_neon, no_push_neon=no_push_neon):
            from run_query import _maybe_push_to_neon

            _maybe_push_to_neon(result, push_neon=push_neon, no_push_neon=no_push_neon)
    return result


def run_phase1_preview(
    query: str,
    country: str,
    *,
    with_search: bool = False,
    log_callback: Callable[[str], None] | None = None,
) -> tuple[dict, str]:
    from test_phase1 import _full_query, build_report

    from vendor_intel.phase1.runner import run_phase1_sync, print_phase1_summary

    from ui.bootstrap import load_settings

    init_env()
    settings = load_settings("quality")
    full_query = _full_query(query, country)

    if not with_search:
        import vendor_intel.pipeline.geo_limits as geo_limits

        original = geo_limits.pipeline_limits

        def plan_only_limits(s, *, recall, country):
            lim = original(s, recall=recall, country=country)
            return {**lim, "smoke_prompts": 0}

        geo_limits.pipeline_limits = plan_only_limits

    def _noop(_: str) -> None:
        pass

    cb = log_callback or _noop
    with capture_stdout(cb):
        manifest = run_phase1_sync(full_query, settings)
        try:
            print_phase1_summary(manifest)
        except UnicodeEncodeError:
            pass

    report_md = build_report(query, country, manifest)
    debug = phase1_debug_dir()
    s = slug(query, country)
    (debug / f"{s}.md").write_text(report_md, encoding="utf-8")
    (debug / f"{s}.json").write_text(
        json.dumps(manifest, indent=2, default=str), encoding="utf-8"
    )
    return manifest, report_md


def shutdown_search_pool() -> None:
    try:
        from vendor_intel.clients.search_stats import log_summary, persist_summary

        log_summary(prefix="search_stats")
        persist_summary()
    except Exception:
        pass
    try:
        from vendor_intel.clients.ddg_worker_pool import shutdown_ddg_pool

        shutdown_ddg_pool(wait=True)
    except Exception:
        pass


# One pipeline at a time — avoids duplicate DDG worker pools and search timeouts.
_pipeline_lock = threading.Lock()


def pipeline_is_busy() -> bool:
    return _pipeline_lock.locked()


class JobRunner:
    """Background thread runner with session-state friendly status dict."""

    def __init__(self, state_key: str = "active_job"):
        self.state_key = state_key

    def start(
        self,
        session_state,
        *,
        job_type: str,
        query: str,
        country: str,
        profile: str = "quality",
        cap: str | None = None,
        pipeline_mode: str = "quality",
        post_enrich: bool = True,
        fmcg_xlsx: bool = True,
        push_neon: bool = False,
        no_push_neon: bool = False,
        skip_name_resolve: bool = False,
        with_search: bool = False,
    ) -> bool:
        if session_state.get(self.state_key, {}).get("running") or pipeline_is_busy():
            return False

        # Mutable dict shared with worker — never read st.session_state from the thread.
        job: dict[str, Any] = {
            "running": True,
            "job_type": job_type,
            "query": query,
            "country": country,
            "profile": profile,
            "cap": cap,
            "pipeline_mode": pipeline_mode,
            "post_enrich": post_enrich,
            "fmcg_xlsx": fmcg_xlsx,
            "push_neon": push_neon,
            "skip_name_resolve": skip_name_resolve,
            "log": "",
            "result": None,
            "error": None,
            "started_at": time.time(),
        }
        session_state[self.state_key] = job

        def worker() -> None:
            if not _pipeline_lock.acquire(blocking=False):
                job["running"] = False
                job["error"] = "Another pipeline is already running on this machine."
                return

            def on_log(chunk: str) -> None:
                job["log"] = (job.get("log") or "") + chunk
                if len(job["log"]) > 120_000:
                    job["log"] = job["log"][-100_000:]

            try:
                if job_type == "pipeline":
                    result = run_full_pipeline(
                        query,
                        country,
                        profile,
                        cap=cap,
                        pipeline_mode=job.get("pipeline_mode") or "quality",
                        post_enrich=bool(job.get("post_enrich", True)),
                        fmcg_xlsx=bool(job.get("fmcg_xlsx", True)),
                        push_neon=bool(job.get("push_neon", False)),
                        no_push_neon=False,
                        skip_name_resolve=bool(job.get("skip_name_resolve", False)),
                        log_callback=on_log,
                    )
                    job["result"] = result
                    llm = result.get("llm_usage") or {}
                    post = result.get("post_market_pipeline") or {}
                    elapsed = time.time() - job["started_at"]
                    append_session_log(
                        {
                            "query": query,
                            "country": country,
                            "status": "ok",
                            "companies_exported": len(result.get("relevant_companies") or []),
                            "founded_hq_rows": post.get("filtered_companies"),
                            "elapsed_seconds": round(elapsed, 1),
                            "elapsed_minutes": round(elapsed / 60, 2),
                            "llm_calls": llm.get("llm_calls_total"),
                            "estimated_cost_usd": llm.get("estimated_cost_usd"),
                            "csv_file": result.get("_csv_path", ""),
                            "xlsx_file": result.get("_xlsx_path", ""),
                            "post_enrich": bool(job.get("post_enrich", True)),
                            "ran_at": datetime.now(timezone.utc).isoformat(),
                        }
                    )
                else:
                    manifest, report_md = run_phase1_preview(
                        query,
                        country,
                        with_search=with_search,
                        log_callback=on_log,
                    )
                    job["result"] = {"manifest": manifest, "report_md": report_md}
            except Exception as exc:
                job["error"] = str(exc)
                if job_type == "pipeline":
                    append_session_log(
                        {
                            "query": query,
                            "country": country,
                            "status": "error",
                            "error": str(exc),
                            "elapsed_seconds": round(time.time() - job["started_at"], 1),
                            "ran_at": datetime.now(timezone.utc).isoformat(),
                        }
                    )
            finally:
                job["running"] = False
                shutdown_search_pool()
                if _pipeline_lock.locked():
                    _pipeline_lock.release()

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        return True

    @staticmethod
    def verdict_badge(manifest: dict) -> tuple[str, list[str]]:
        from test_phase1 import _verdict

        scope = manifest.get("scope") or {}
        return _verdict(scope, manifest)


def markdown_to_plain_preview(md: str, max_lines: int = 40) -> str:
    lines = md.splitlines()[:max_lines]
    text = "\n".join(lines)
    return re.sub(r"\*\*", "", text)
