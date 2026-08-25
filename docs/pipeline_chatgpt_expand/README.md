# `run_chatgpt_expand_batch.py` — Multi-Market ChatGPT Expand Pipeline

**Generic pipeline — works for 1 to 10,000+ markets from one input file.**
No market name is hardcoded; each row in the input file is just data fed
through the same 8-step per-market flow.

See [`chatgpt_expand_flowchart.pdf`](chatgpt_expand_flowchart.pdf) /
[`chatgpt_expand_flowchart.png`](chatgpt_expand_flowchart.png) for the full
diagram.

## What it does

`run_chatgpt_expand_batch.py` is the bulk / large-scale version of the
ChatGPT-only expand pipeline (same underlying steps as
`run_cursor_expand.py --backend chatgpt`, but driven from a market list +
SQLite registry so it can process thousands of markets with resume and
concurrency support). For **every market in the input file** it:

1. Recalls known channel companies for that market via ChatGPT.
2. Discovers new candidates via OpenAI's hosted `web_search`.
3. Optionally harvests more candidates via DDGS/SearXNG.
4. Verifies keep/reject relevance for every candidate.
5. Fills in missing columns via a waterfall of free/cheap sources before
   falling back to OpenAI web_search.
6. Runs a field-level QA pass (categories, brands, channel fit, contacts,
   geography).
7. Runs a final ChatGPT verify pass and writes the FINAL Excel + audit JSON.

A local SQLite registry (keyed by `--batch-id`) tracks each market's status
(`pending` / `running` / `done` / `failed`), which is what makes the whole
thing resumable and safe to run unattended over very large lists.

## How to run it

```bash
# Load ~10k markets, then expand — 1 market at a time (safest for API limits)
python run_chatgpt_expand_batch.py \
    --input queries/my_markets.csv --batch-id expand-10k --yes

# Resume after Ctrl+C — done markets are skipped automatically
python run_chatgpt_expand_batch.py --batch-id expand-10k --yes

# Smoke test: 2 markets, 2 at once
python run_chatgpt_expand_batch.py \
    --input queries/my_markets.csv --batch-id expand-smoke \
    --limit 2 --max-concurrent 1 --yes
```

Input file format: a `.csv` with `query|name` per row, or a `.txt` with one
market per line — the market name is free text, same as `run_query.py`.

## Pipeline stages

| Stage | What happens |
|---|---|
| Load `.env` | OpenAI/DeepSeek, Apollo, LinkedIn keys, etc. |
| Populate registry | Loads `--input` into a SQLite table keyed by `--batch-id`; new markets inserted, duplicates skipped |
| Reap / requeue | Stale `running` rows (e.g. from a crash) go back to `pending`; `failed` rows are requeued automatically; `--fresh` also resets `done` rows |
| Pending check | If nothing is pending, the run exits immediately (nothing to do) |
| Confirm | Prints the pending count + concurrency and asks to proceed (skip with `--yes`) |
| **Per-market expand** (steps 0–7) | Runs for every pending market, up to `--max-concurrent` in parallel — see below |
| Mark done/failed | Registry updated per market; failures are logged and auto-requeued on the next run |
| Loop | Pulls the next pending market until none remain |
| Final summary | Prints done/failed/pending counts |

### Per-market expand (steps 0–7)

0. **Seed candidates** — load curated seed companies + create a blank Excel scaffold
1. **ChatGPT recall** — known channel companies for this market
2. **OpenAI web_search discovery** — new candidate companies
3. *(optional)* **DDGS/SearXNG harvest** + ChatGPT extract & score relevance (`--ddgs-harvest`)
4. **ChatGPT verify** — keep/reject each candidate against the market definition
5. **Fill columns** — waterfall: Google AI scraper → DDGS/SearXNG/Wikipedia/Owler → OpenAI `web_search` residual for anything still blank
6. **Field verify & correct (QA)** — Core Categories, Specialty Focus, Key Brands Represented + channel-fit keep/remove, Online/Offline/Both classification, Contact Person/Role/LinkedIn via Apollo free search + web-search fallback, multi-location geography
7. **Final verify + Excel** — ChatGPT final pass on filled rows, writes the wrap-formatted FINAL Excel + an audit JSON

## Key flags

| Flag | Purpose |
|---|---|
| `--input <path>` | Market list to load (CSV or TXT) — omit to just resume `--batch-id` |
| `--batch-id <id>` | Registry key; reuse the same id to resume |
| `--limit N` | Process at most N markets this run |
| `--max-concurrent N` | Markets processed in parallel (default 1 — safest for API rate limits) |
| `--target` / `--max-fill` / `--max-queries` | Row-count and query-budget tuning per market |
| `--fill-backend search\|openai` | `search` = free-source waterfall first (default); `openai` = OpenAI residual only |
| `--ddgs-harvest` / `--no-ddgs-harvest` | Enable/disable optional Step 3 |
| `--no-openai-discover` / `--no-web-fill` / `--no-apollo` / `--no-linkedin-enrich` | Skip individual steps |
| `--seeds-only` | Only export curated seed companies, skip discovery |
| `--skip-recall` | Skip Step 1 (ChatGPT recall) |
| `--fresh` | Ignore done/failed checkpoints and redo every market in the batch (use with care on large batches) |
| `--yes` | Skip the confirmation prompt |

## Resuming

The pipeline is **Ctrl+C-safe**: re-running the exact same `--batch-id`
command resumes automatically —

- Markets stuck in `running` (from an interrupted process) are reaped back
  to `pending`.
- Markets that previously `failed` are requeued to `pending`.
- Markets already `done` are skipped.

To force a full redo of a batch, pass `--fresh` (or start a new
`--batch-id`).

## Output

Each market gets its own FINAL Excel + audit JSON (path depends on the
configured output directory — see `web_expand.default_output_dir` /
`resolve_final_path`). Progress/status is tracked in the SQLite registry,
not in per-market log files, so `python run_chatgpt_expand_batch.py
--batch-id <id>` (no `--input`) at any time will show/resume where the
batch left off.
