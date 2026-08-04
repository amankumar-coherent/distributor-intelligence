#!/usr/bin/env python3
"""OpenAI cost dashboard for ChatGPT expand runs — auto-refreshes while a run is live.

Run once (leave open):
  .venv\\Scripts\\python.exe -m streamlit run cost_dashboard.py

Then in another terminal start expand. This page updates automatically.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.pipeline import openai_cost  # noqa: E402

st.set_page_config(page_title="OpenAI Cost Dashboard", page_icon="💲", layout="wide")
st.title("OpenAI Cost Dashboard")
st.caption(
    "Auto-refreshes every few seconds. Leave this tab open while "
    "`run_cursor_expand.py --backend chatgpt` runs — no need to stop or re-run."
)

auto = st.sidebar.toggle("Auto-refresh", value=True)
interval = st.sidebar.slider("Refresh every (seconds)", 2, 30, 5)
st.sidebar.markdown(f"Logs folder:\n`{openai_cost.cost_dir()}`")


def _render_body() -> None:
    totals = openai_cost.totals()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total estimated cost", f"${totals['estimated_cost_usd']:.4f}")
    c2.metric("Total tokens", f"{totals['total_tokens']:,}")
    c3.metric("API calls", f"{totals['calls']:,}")
    c4.metric("Runs", f"{totals['runs']:,}")

    live_path = openai_cost.cost_dir() / "live.json"
    st.subheader("Live / last run")
    if live_path.exists():
        live = json.loads(live_path.read_text(encoding="utf-8"))
        ended = live.get("ended_at")
        status = "finished" if ended else "RUNNING"
        st.markdown(f"**Status:** `{status}` · updated `{live.get('updated_at', '—')}`")
        l1, l2, l3, l4 = st.columns(4)
        l1.metric("Market", str(live.get("market") or "—")[:40])
        l2.metric("Run cost", f"${float(live.get('estimated_cost_usd') or 0):.4f}")
        l3.metric("Tokens", f"{int(live.get('total_tokens') or 0):,}")
        l4.metric("Calls", f"{int(live.get('calls') or 0):,}")
        if live.get("last_call"):
            lc = live["last_call"]
            st.caption(
                f"Last call: `{lc.get('label')}` · "
                f"in={lc.get('prompt_tokens')} out={lc.get('completion_tokens')} · "
                f"${float(lc.get('estimated_cost_usd') or 0):.4f}"
            )
    else:
        st.warning(
            "**No cost log yet.** Common reasons:\n\n"
            "1. Expand was started **before** cost tracking was added — **stop it (Ctrl+C) and start again**.\n"
            "2. Expand is on **Step 2 web search** — that step does **not** call OpenAI, so $ may not move until Step 3.\n"
            "3. Wrong folder — logs must be under `D:\\vendor-intel\\output\\cost_dashboard\\`.\n\n"
            "After a fresh expand start you should see `cost run_id=...` in the terminal and `live.json` appear here."
        )

    st.subheader("Runs")
    runs = openai_cost.load_runs()
    if runs:
        df_runs = pd.DataFrame(runs)
        cols = [
            c
            for c in [
                "started_at",
                "ended_at",
                "market",
                "model",
                "calls",
                "prompt_tokens",
                "completion_tokens",
                "total_tokens",
                "estimated_cost_usd",
                "run_id",
            ]
            if c in df_runs.columns
        ]
        st.dataframe(df_runs[cols].iloc[::-1], width="stretch", hide_index=True)
    else:
        st.info("No completed runs logged yet.")

    st.subheader("Recent API calls")
    calls = openai_cost.load_calls(500)
    if calls:
        df_calls = pd.DataFrame(calls)
        cols = [
            c
            for c in [
                "ts",
                "market",
                "label",
                "model",
                "prompt_tokens",
                "completion_tokens",
                "estimated_cost_usd",
                "run_id",
            ]
            if c in df_calls.columns
        ]
        st.dataframe(df_calls[cols].iloc[::-1], width="stretch", hide_index=True)

        if "estimated_cost_usd" in df_calls.columns and "label" in df_calls.columns:
            st.subheader("Cost by step")
            by_label = (
                df_calls.groupby(df_calls["label"].str.split("-").str[0], dropna=False)[
                    "estimated_cost_usd"
                ]
                .sum()
                .sort_values(ascending=False)
            )
            st.bar_chart(by_label)
    else:
        st.info("No call logs yet.")

    st.markdown("---")
    st.caption(
        "Costs are estimates from token counts × public OpenAI list prices. "
        "Exact billing is on your OpenAI invoice."
    )


if auto:
    @st.fragment(run_every=interval)
    def _auto() -> None:
        _render_body()

    _auto()
else:
    if st.button("Refresh now"):
        st.rerun()
    _render_body()
