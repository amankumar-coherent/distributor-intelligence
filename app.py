#!/usr/bin/env python3
"""
Vendor Intelligence — Streamlit UI

Run:
  .venv\\Scripts\\python.exe -m streamlit run app.py

Or:
  streamlit run app.py
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import pandas as pd
import streamlit as st

from ui.bootstrap import (
    ROOT,
    env_warnings,
    init_env,
    live_validation_warnings,
    load_settings,
    output_dir,
    phase1_debug_dir,
    pipeline_caps,
)
from ui.services import (
    JobRunner,
    find_market_artifacts,
    get_integration_status,
    list_result_csvs,
    load_queries_file,
    load_results_dataframe,
    load_session_log,
    pipeline_is_busy,
    result_csv_label,
    slug,
)
from ui.styles import CUSTOM_CSS

st.set_page_config(
    page_title="Vendor Intelligence",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_env()

# ---------------------------------------------------------------------------
# Session defaults
# ---------------------------------------------------------------------------
if "active_job" not in st.session_state:
    st.session_state.active_job = {"running": False}
if "selected_csv" not in st.session_state:
    st.session_state.selected_csv = None

JOB_RUNNER = JobRunner("active_job")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
PROFILE_HELP = {
    "quality": "Strict export gate — CEO-ready CSV, fewer rows.",
    "balanced": "Middle ground between quality and recall.",
    "recall": "More companies exported; noisier results.",
}

PHASES = [
    ("1", "Market plan", "LLM maps the market, keywords, seeds"),
    ("2", "Discovery", "Web search finds candidate companies"),
    ("3", "Enrichment", "Site crawl + signals extraction"),
    ("4", "Classification", "LLM assigns role & relevance"),
    ("5", "Export", "Quality filter → CSV"),
]


GEO_EXAMPLES = "global · United States · Europe · North America · Latin America · APAC · MENA · India · Germany"


def geography_input(label: str = "Geography", *, key: str | None = None) -> str:
    """Free-text geography — same as CLI --country."""
    if key and key not in st.session_state:
        st.session_state[key] = "global"
    kwargs: dict = {
        "label": label,
        "placeholder": "e.g. United States, Europe, North America",
        "help": 'Type any country or region. Use "global" for worldwide.',
    }
    if key:
        kwargs["key"] = key
    raw = st.text_input(**kwargs)
    if raw is None:
        raw = ""
    return raw.strip() or "global"


def render_hero(title: str, subtitle: str) -> None:
    st.markdown(
        f'<div class="hero"><h1>{title}</h1><p>{subtitle}</p></div>',
        unsafe_allow_html=True,
    )


def metric_card(label: str, value: str) -> None:
    st.markdown(
        f'<div class="metric-card"><div class="label">{label}</div>'
        f'<div class="value">{value}</div></div>',
        unsafe_allow_html=True,
    )


def render_phase_pills() -> None:
    pills = "".join(f'<span class="phase-pill">Phase {n}: {name}</span>' for n, name, _ in PHASES)
    st.markdown(pills, unsafe_allow_html=True)


def status_class(verdict: str) -> str:
    v = verdict.upper()
    if v == "GOOD":
        return "status-ok"
    if v in ("WARN", "PARTIAL"):
        return "status-warn"
    return "status-fail"


@st.fragment(run_every=3)
def _live_job_status() -> None:
    """Auto-refresh only this block while a job runs (avoids full-page rerun spam)."""
    job = st.session_state.get("active_job") or {}
    if not (job.get("running") or job.get("result") or job.get("error")):
        return

    st.markdown("### Run status")
    if job.get("running"):
        elapsed = int(time.time() - job.get("started_at", time.time()))
        st.progress(
            min(elapsed / 1500, 0.95),
            text=f"Running… {elapsed // 60}m {elapsed % 60}s elapsed",
        )
        st.info(
            "Keep this tab open. Post-enrich uses LinkedIn MCP, gap-fill, Apollo — "
            "can take 1–4 hours on large markets."
        )
    elif job.get("error"):
        st.error(f"Pipeline failed: {job['error']}")
    elif job.get("result"):
        result = job["result"]
        if job.get("job_type") == "pipeline":
            companies = len(result.get("relevant_companies") or [])
            llm = result.get("llm_usage") or {}
            post = result.get("post_market_pipeline") or {}
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Exported", companies)
            m2.metric("Founded+HQ", post.get("filtered_companies", "—"))
            m3.metric("LLM calls", llm.get("llm_calls_total", "—"))
            m4.metric("Est. cost", f"${llm.get('estimated_cost_usd', 0):.4f}")
            if post:
                audit = post.get("market_audit") or {}
                if audit.get("queued") is not None:
                    st.caption(
                        f"Post-enrich audit: {audit.get('kept', 0)} keep · "
                        f"{audit.get('review', 0)} review · {audit.get('removed', 0)} removed"
                    )
                neon = post.get("neon_push") or {}
                if neon.get("vendor_intel_companies"):
                    st.caption(f"Neon pushed: {neon['vendor_intel_companies']} companies")
            if result.get("_csv_path"):
                st.success(f"CSV saved: `{result['_csv_path']}`")
            if result.get("_xlsx_path"):
                st.success(f"Excel saved: `{result['_xlsx_path']}`")
            excel = (post.get("excel") or {}).get("file")
            if excel:
                st.success(f"Founded+HQ Excel: `{excel}`")
            if result.get("_csv_path"):
                try:
                    df, label, _ = load_results_dataframe(Path(result["_csv_path"]), view="landscape")
                    st.caption(f"Preview — {label} ({len(df.columns)} columns)")
                    st.dataframe(df, width="stretch", hide_index=True)
                    st.download_button(
                        "Download full landscape CSV",
                        df.to_csv(index=False).encode("utf-8"),
                        file_name=Path(result["_csv_path"]).stem + "_landscape.csv",
                        mime="text/csv",
                        key="dl_pipeline_landscape",
                    )
                except Exception as exc:
                    st.warning(str(exc))
        else:
            manifest = result.get("manifest") or {}
            report_md = result.get("report_md") or ""
            verdict, notes = JOB_RUNNER.verdict_badge(manifest)
            st.markdown(
                f'Verdict: <span class="{status_class(verdict)}">{verdict}</span>',
                unsafe_allow_html=True,
            )
            for note in notes:
                st.markdown(f"- {note}")
            scope = manifest.get("scope") or {}
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Discovery prompts", len(manifest.get("discovery_prompts") or []))
            c2.metric("Seed companies", len(scope.get("seed_companies") or []))
            c3.metric("Include keywords", len(scope.get("include_keywords") or []))
            c4.metric("Value-chain layers", len(scope.get("value_chain_layers") or []))
            defn = str(scope.get("market_definition") or "").strip()
            if defn:
                st.markdown("**Market definition**")
                st.info(defn)
            with st.expander("Full Phase 1 report", expanded=False):
                st.markdown(report_md)
            if job.get("query") and job.get("country"):
                s = slug(job["query"], job["country"])
                st.caption(f"Saved to `output/phase1_debug/{s}.json`")

    if job.get("log"):
        st.markdown("**Live log**")
        st.markdown(f'<div class="log-box">{job["log"][-8000:]}</div>', unsafe_allow_html=True)

    if not job.get("running") and (job.get("result") or job.get("error")):
        if st.button("Clear run status", key="clear_job_status"):
            st.session_state.active_job = {"running": False}
            st.rerun()


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
def page_home() -> None:
    render_hero(
        "Vendor Intelligence",
        "Discover market participants, classify their roles, and export CEO-quality landscape CSVs.",
    )

    csvs = list_result_csvs()
    log = load_session_log()
    ok_runs = [e for e in log if e.get("status") == "ok"]
    total_exported = sum(int(e.get("companies_exported") or 0) for e in ok_runs)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Demo CSVs", str(len(csvs)))
    with c2:
        metric_card("Pipeline runs", str(len(log)))
    with c3:
        metric_card("Companies exported", str(total_exported))
    with c4:
        out = output_dir()
        metric_card("Output folder", out.name)

    st.markdown("### Pipeline at a glance")
    render_phase_pills()
    for n, name, desc in PHASES:
        st.markdown(f"**Phase {n} — {name}:** {desc}")

    st.markdown("### Recent runs")
    if not log:
        st.info("No runs yet. Go to **Run Pipeline** to start your first market query.")
    else:
        for entry in reversed(log[-8:]):
            status = entry.get("status", "?")
            q = entry.get("query", "?")
            c = entry.get("country", "global")
            n_co = entry.get("companies_exported", "—")
            t = entry.get("elapsed_minutes", "—")
            icon = "✅" if status == "ok" else "❌"
            st.markdown(
                f'<div class="result-card">{icon} <strong>{q}</strong> ({c}) '
                f"— {n_co} companies · {t} min</div>",
                unsafe_allow_html=True,
            )

    if csvs:
        st.markdown("### Latest results")
        latest = csvs[0]
        try:
            df = pd.read_csv(latest)
            st.caption(result_csv_label(latest))
            st.dataframe(df.head(8), width="stretch", hide_index=True)
        except Exception as exc:
            st.warning(f"Could not preview {latest.name}: {exc}")


def page_run_pipeline() -> None:
    render_hero(
        "Run Pipeline",
        "Phases 1–5 discovery + optional 8-step post-enrich (LinkedIn MCP, gap-fill, Apollo, Neon).",
    )

    warnings = env_warnings()
    for w in warnings:
        st.warning(w)

    with st.expander("Integrations status (LinkedIn MCP, Jina, Exa, …)", expanded=False):
        st.dataframe(
            pd.DataFrame(get_integration_status()),
            hide_index=True,
            width="stretch",
        )

    with st.form("pipeline_form", border=True):
        col1, col2 = st.columns([2, 1])
        with col1:
            query = st.text_input(
                "Market query",
                placeholder='e.g. "Global Smartwatch Market" or "ERP vendors"',
            )
        with col2:
            country = geography_input(key="run_country")
            st.caption(f"Examples: {GEO_EXAMPLES}")
        from vendor_intel.pipeline.cap_profiles import CAP_TIERS, DEFAULT_CAP, cap_keys

        colp, colc = st.columns(2)
        with colp:
            profile = st.selectbox(
                "Pipeline profile",
                ["quality", "balanced", "recall"],
                index=2,
                help=PROFILE_HELP["quality"],
            )
        with colc:
            _caps = cap_keys()
            cap = st.selectbox(
                "Result cap (how many companies)",
                _caps,
                index=_caps.index("maximum") if "maximum" in _caps else _caps.index(DEFAULT_CAP),
                format_func=lambda k: f"{CAP_TIERS[k]['label']} — {CAP_TIERS[k]['approx']}",
                help="Bigger cap = more companies + more discovery, slower run.",
            )
        st.caption(PROFILE_HELP.get(profile, ""))
        _t = CAP_TIERS[cap]
        st.caption(
            f"**{_t['label']}** cap → {_t['approx']}, {_t['time']} "
            f"(discover≤{_t['discover']}, export≤{_t['export']})."
        )

        st.markdown("**Post-export (same as terminal `run_query.py --mode quality`)**")
        c1, c2, c3 = st.columns(3)
        with c1:
            post_enrich = st.checkbox(
                "8-step post-enrich",
                value=True,
                help="Gap-fill → LinkedIn MCP verify → Apollo → name resolve → brand fix → audit → Excel → Neon",
            )
        with c2:
            fmcg_xlsx = st.checkbox("30-column FMCG Excel", value=True)
        with c3:
            push_neon = st.checkbox("Push to Neon", value=True)
        skip_name_resolve = st.checkbox(
            "Skip slow name resolve (step 4)",
            value=False,
            help="Faster post-enrich; uses LinkedIn MCP + leadership only for contacts.",
        )

        submitted = st.form_submit_button("Start pipeline run", type="primary", width="stretch")

    job = st.session_state.active_job
    if submitted:
        if not query.strip():
            st.error("Enter a market query.")
        elif job.get("running") or pipeline_is_busy():
            st.warning("A pipeline is already running — wait for it to finish.")
        elif not JOB_RUNNER.start(
            st.session_state,
            job_type="pipeline",
            query=query.strip(),
            country=country,
            profile=profile,
            cap=cap,
            pipeline_mode="quality" if post_enrich else "standard",
            post_enrich=post_enrich,
            fmcg_xlsx=fmcg_xlsx,
            push_neon=push_neon,
            skip_name_resolve=skip_name_resolve,
        ):
            st.warning("Could not start — another pipeline may be running.")
        else:
            st.rerun()

    _live_job_status()


def page_phase1() -> None:
    render_hero(
        "Phase 1 Preview",
        "Fast market-understanding check (~1–3 min) before a full pipeline run.",
    )

    with st.form("phase1_form", border=True):
        col1, col2 = st.columns([2, 1])
        with col1:
            query = st.text_input("Market query", key="p1_query")
        with col2:
            country = geography_input(key="p1_country")
            st.caption(f"Examples: {GEO_EXAMPLES}")
        with_search = st.checkbox("Include search smoke test (slower)", value=False)
        go = st.form_submit_button("Run Phase 1 preview", type="primary")

    job = st.session_state.active_job
    if go:
        if not query.strip():
            st.error("Enter a market query.")
        elif job.get("running") or pipeline_is_busy():
            st.warning("A job is already running — wait for it to finish.")
        elif not JOB_RUNNER.start(
            st.session_state,
            job_type="phase1",
            query=query.strip(),
            country=country,
            with_search=with_search,
        ):
            st.warning("Could not start — another pipeline may be running.")
        else:
            st.rerun()

    _live_job_status()


def page_results() -> None:
    render_hero("Results Explorer", "Browse all columns — pipeline CSV, 30-col landscape, or Founded+HQ Excel.")

    csvs = list_result_csvs()
    if not csvs:
        st.info(f"No CSV files in `{output_dir()}`. Run a pipeline first.")
        return

    labels = [result_csv_label(p) for p in csvs]
    default_idx = 0
    if st.session_state.selected_csv in labels:
        default_idx = labels.index(st.session_state.selected_csv)

    selected_label = st.selectbox("Select result file", labels, index=default_idx)
    selected = csvs[labels.index(selected_label)]
    st.session_state.selected_csv = selected_label

    artifacts = find_market_artifacts(selected)
    view_options = ["landscape", "csv"]
    if artifacts.get("founded_hq_xlsx"):
        view_options.append("founded_hq")
    view = st.radio(
        "Column set",
        view_options,
        format_func=lambda v: {
            "landscape": "Full landscape (30+ columns)",
            "csv": "Pipeline CSV (fewer columns)",
            "founded_hq": "Founded + HQ only (post-enrich Excel)",
        }.get(v, v),
        horizontal=True,
    )

    try:
        df, view_label, _ = load_results_dataframe(selected, view=view)
    except Exception as exc:
        st.error(f"Could not load export: {exc}")
        return

    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button(
            "Download view as CSV",
            df.to_csv(index=False).encode("utf-8"),
            file_name=f"{selected.stem}_{view}.csv",
            mime="text/csv",
        )
    with dl_cols[1]:
        if artifacts.get("fmcg_xlsx"):
            st.download_button(
                "Download FMCG Excel",
                artifacts["fmcg_xlsx"].read_bytes(),
                file_name=artifacts["fmcg_xlsx"].name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    with dl_cols[2]:
        if artifacts.get("founded_hq_xlsx"):
            st.download_button(
                "Download Founded+HQ Excel",
                artifacts["founded_hq_xlsx"].read_bytes(),
                file_name=artifacts["founded_hq_xlsx"].name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows", len(df))
    c2.metric("Columns", len(df.columns))
    c3.metric("View", view_label[:24] + ("…" if len(view_label) > 24 else ""))
    if "Role" in df.columns:
        c4.metric("Unique roles", df["Role"].nunique())
    elif "Confidence" in df.columns:
        c4.metric("Avg confidence", f"{df['Confidence'].mean():.2f}")

    col_a, col_b = st.columns(2)
    with col_a:
        role_col = "Role" if "Role" in df.columns else None
        role_filter = st.multiselect(
            "Filter by role",
            sorted(df[role_col].dropna().unique()) if role_col else [],
        )
    with col_b:
        search = st.text_input("Search company / brand", "")

    filtered = df.copy()
    if role_filter and role_col:
        filtered = filtered[filtered[role_col].isin(role_filter)]
    if search:
        mask = pd.Series(False, index=filtered.index)
        for col in ("Company", "Brand", "Key_Products", "Functionality", "Role_Description", "Contact Person"):
            if col in filtered.columns:
                mask |= filtered[col].astype(str).str.contains(search, case=False, na=False)
        filtered = filtered[mask]

    st.dataframe(
        filtered,
        width="stretch",
        hide_index=True,
        column_config={
            "Website": st.column_config.LinkColumn("Website"),
            "LinkedIn": st.column_config.LinkColumn("LinkedIn"),
            "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0, max_value=1),
            "Quality_Score": st.column_config.ProgressColumn("Quality", min_value=0, max_value=1),
            "Data Confidence": st.column_config.TextColumn("Data Confidence"),
        },
    )

    role_col_chart = "Role" if "Role" in df.columns else None
    if role_col_chart and not df[role_col_chart].value_counts().empty:
        st.markdown("### Role distribution")
        st.bar_chart(df[role_col_chart].value_counts())

    json_path = artifacts.get("json") or selected.with_suffix(".json")
    if json_path and Path(json_path).exists():
        with st.expander("Pipeline JSON summary"):
            try:
                summary = json.loads(Path(json_path).read_text(encoding="utf-8"))
                show_keys = (
                    "query",
                    "scope",
                    "pipeline_profile",
                    "elapsed_minutes",
                    "llm_usage",
                    "post_market_pipeline",
                )
                st.json({k: summary.get(k) for k in show_keys if k in summary})
            except Exception:
                st.caption("Could not parse JSON.")


def page_batch() -> None:
    render_hero("Batch Runner", "Run multiple markets from `queries/markets.txt` or a custom list.")

    markets_path = ROOT / "queries" / "markets.txt"
    preset = load_queries_file(markets_path)

    source = st.radio("Query source", ["Preset markets.txt", "Custom list"], horizontal=True)

    if source == "Preset markets.txt":
        if not preset:
            st.warning(f"No queries in `{markets_path}`.")
            return
        st.markdown("**Queries to run:**")
        for i, (q, c) in enumerate(preset, 1):
            st.markdown(f"{i}. {q} · *{c}*")
        queries = preset
    else:
        raw = st.text_area(
            "One query per line",
            placeholder="Atomic Clock Market | global\nERP vendors | Europe",
            height=160,
        )
        queries = []
        for line in raw.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            from ui.services import parse_query_line

            queries.append(parse_query_line(line))
        if not queries:
            st.caption("Enter at least one query.")

    profile = st.selectbox("Profile", ["quality", "balanced", "recall"], key="batch_profile")

    if st.button("Run batch (sequential)", type="primary", disabled=not queries):
        progress = st.progress(0.0)
        log_box = st.empty()
        results: list[dict] = []

        for i, (q, c) in enumerate(queries):
            progress.progress(i / len(queries), text=f"Running {i + 1}/{len(queries)}: {q}")
            log_lines = [f"=== {q} ({c}) ===\n"]

            def on_log(chunk: str, acc=log_lines) -> None:
                acc[0] += chunk
                log_box.markdown(
                    f'<div class="log-box">{"".join(acc)[-6000:]}</div>',
                    unsafe_allow_html=True,
                )

            try:
                from ui.services import run_full_pipeline, shutdown_search_pool

                result = run_full_pipeline(q, c, profile, log_callback=on_log)
                results.append(
                    {
                        "query": q,
                        "country": c,
                        "status": "ok",
                        "exported": len(result.get("relevant_companies") or []),
                        "csv": result.get("_csv_path", ""),
                    }
                )
            except Exception as exc:
                results.append({"query": q, "country": c, "status": "error", "error": str(exc)})
            finally:
                shutdown_search_pool()

        progress.progress(1.0, text="Batch complete")
        st.markdown("### Batch summary")
        for r in results:
            if r["status"] == "ok":
                st.success(f"{r['query']} ({r['country']}) — {r['exported']} rows → `{r['csv']}`")
            else:
                st.error(f"{r['query']} ({r['country']}) — {r.get('error', 'failed')}")

        summary_path = output_dir() / "run_summary.json"
        summary_path.write_text(json.dumps(results, indent=2, default=str), encoding="utf-8")
        st.caption(f"Saved `{summary_path}`")


def page_health() -> None:
    render_hero("System Health", "Configuration and API readiness before live runs.")

    settings = load_settings()
    warnings = env_warnings() + live_validation_warnings(settings)

    if warnings:
        st.markdown("### Warnings")
        for w in warnings:
            st.warning(w)
    else:
        st.success("Configuration looks good for live runs.")

    st.markdown("### Integrations")
    st.dataframe(pd.DataFrame(get_integration_status()), hide_index=True, width="stretch")

    try:
        from vendor_intel.clients.search_stats import snapshot

        st.markdown("### Search stats (this session)")
        st.json(snapshot())
    except Exception:
        pass

    st.markdown("### Environment")
    import os

    rows = [
        ("USE_MOCK_DATA", os.getenv("USE_MOCK_DATA", "true")),
        ("LLM_PROVIDER", os.getenv("LLM_PROVIDER", "anthropic")),
        ("PIPELINE_PROFILE", os.getenv("PIPELINE_PROFILE", "quality")),
        ("MARKET_QUERY_OUTPUT_DIR", os.getenv("MARKET_QUERY_OUTPUT_DIR", "output/demo")),
        ("DDG_WORKER_COUNT", os.getenv("DDG_WORKER_COUNT", "0")),
        ("WEB_FETCH_ENABLED", os.getenv("WEB_FETCH_ENABLED", "true")),
    ]
    st.dataframe(pd.DataFrame(rows, columns=["Variable", "Value"]), hide_index=True, width="stretch")

    classify_cap, enrich_cap = pipeline_caps(settings)
    st.markdown("### Pipeline caps (global)")
    c1, c2 = st.columns(2)
    c1.metric("Discover cap", classify_cap)
    c2.metric("Enrich cap", enrich_cap)

    st.markdown("### Output paths")
    st.code(
        f"Demo CSV/JSON : {output_dir()}\n"
        f"Phase 1 debug : {phase1_debug_dir()}\n"
        f"Project root   : {ROOT}",
        language="text",
    )


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

with st.sidebar:
    st.markdown('<div class="sidebar-brand">🔍 Vendor Intelligence</div>', unsafe_allow_html=True)
    page = st.radio(
        "Navigate",
        [
            "Home",
            "Run Pipeline",
            "Phase 1 Preview",
            "Results",
            "Batch Run",
            "System Health",
        ],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption(f"Output: `{output_dir().name}/`")
    st.caption("Pipeline ~25–45 min · +post-enrich ~1–4 hr")

PAGES = {
    "Home": page_home,
    "Run Pipeline": page_run_pipeline,
    "Phase 1 Preview": page_phase1,
    "Results": page_results,
    "Batch Run": page_batch,
    "System Health": page_health,
}

PAGES[page]()
