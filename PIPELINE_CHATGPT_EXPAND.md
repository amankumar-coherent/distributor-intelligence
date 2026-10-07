# Pipeline: `chatgpt_expand`

Reference for the OpenAI-SDK market expand pipeline. Written so Claude or
an engineer can work on it from a cold start on another machine.

Source of truth is the code — this summarises it, it does not replace it:

| Thing | Path |
|---|---|
| Pipeline | `src/vendor_intel/pipeline/chatgpt_expand.py` (~2,200 lines) |
| Entry point | `run_chatgpt_expand(...)` — async, L1521 |
| Single-market CLI | `run_cursor_expand.py --backend chatgpt` |
| Multi-market CLI | `run_chatgpt_expand_batch.py` (SQLite registry, resume) |
| Checkpoint | `src/vendor_intel/pipeline/openai_expand_checkpoint.py` |

---

## 1. What it does

Given a market name, produce a verified landscape of channel companies
(distributors / resellers / channel partners) with ~35 populated columns,
as a formatted FINAL Excel plus an audit JSON.

It is **discovery + enrichment + verification in one run**, unlike the
Google-AI-Mode market-research pipeline which separates discovery from a
later verification pass.

---

## 2. The seven steps

From the module docstring, which is accurate and worth reading in full:

| Step | What happens |
|---|---|
| 1 | ChatGPT **recalls** known channel companies for the market |
| 2 | OpenAI SDK **`web_search` discovery** (+ optional ddgs harvest) |
| 3 | ChatGPT **extracts + scores** market relevance from ddgs hits (only if harvest used) |
| 4 | ChatGPT **verifies** keep/reject against the market |
| 5 | **Fill columns** — Google AI scraper → DDGS / SearXNG / Wikipedia / Owler first, then OpenAI `web_search` for residual gaps |
| 6 | **Per-field verify/correct** (`field_verify_fill.py`): Core Categories, Specialty Focus, Key Brands, channel-fit keep/remove, Online/Offline/Both, Contact Person/Role/LinkedIn (Apollo free tier + web fallback), multi-location geography |
| 7 | **Final verify** on filled rows → wrap-formatted FINAL Excel + audit JSON |

Key ordering decision in step 5: the **cheap search stack runs before the
paid model**. OpenAI `web_search` is only used for leftover
"Not publicly disclosed" gaps. Reversing this works but costs far more.

---

## 3. Running it

### Single market

```bash
python run_cursor_expand.py --backend chatgpt --query "Global Drone Sensors Market" --country india
```

### Many markets (preferred for anything beyond a one-off)

```bash
# Load a market list and start
python run_chatgpt_expand_batch.py --input queries/neon_markets_all.csv --batch-id expand-10k --yes

# Resume after Ctrl+C — completed markets skipped automatically
python run_chatgpt_expand_batch.py --batch-id expand-10k --yes

# Smoke test: 2 markets, one at a time
python run_chatgpt_expand_batch.py --input queries/neon_markets_all.csv \
    --batch-id expand-smoke --limit 2 --max-concurrent 1 --yes
```

`--batch-id` is the resume key. Reuse the same id to continue; use a new
one to start fresh.

---

## 4. Entry point arguments

`run_chatgpt_expand()` is **async** — call it with `asyncio.run(...)` if
invoking directly rather than through a CLI.

```python
async def run_chatgpt_expand(
    query: str, *,
    country: str = "global",
    target: int = 320,            # companies wanted
    batch: str = "all",
    max_queries: int = 60,        # search-query budget
    merge_existing: bool = True,
    final_path: str | None = None,
    out_dir: str | None = None,
    fill_existing: bool = False,  # fill an existing sheet, skip discovery
    max_fill: int = 500,
    seeds_only: bool = False,     # use only queries/seeds/<market>.txt
    seeds_path: str | None = None,
    skip_recall: bool = False,    # skip step 1
    web_fill: bool = True,
    openai_discover: bool = True, # skip to disable step 2
    ddgs_harvest: bool = False,   # enables step 3
    resume: bool = True,
    use_linkedin_enrich: bool = True,
    use_apollo: bool = True,
    fill_backend: str = "search", # "search" (cheap first) vs model-only
    force: bool = False,
    family: str | None = None,
) -> dict[str, Any]
```

Useful CLI flags on the batch runner: `--target`, `--max-fill`,
`--max-queries`, `--ddgs-harvest`, `--seeds-only`, `--skip-recall`,
`--no-openai-discover`, `--no-web-fill`, `--limit`, `--max-concurrent`.

---

## 5. Provider configuration

| Env var | Meaning |
|---|---|
| `OPENAI_BASE_URL` | Points the OpenAI SDK at OpenAI or a compatible endpoint |
| `OPENAI_MODEL` | Model id |
| `EXPAND_SKIP_DDGS_HARVEST` | Skip the optional ddgs harvest |
| `EXPAND_RESIDUAL_CONCURRENT` | Concurrency for residual web-search fill |

**DeepSeek is supported but degraded, by design.** `_is_deepseek()` and
`_supports_hosted_web_search()` detect it, because hosted `web_search` is
OpenAI-only — on DeepSeek the pipeline falls back to `chat.completions`
and loses step 2's web discovery. If discovery suddenly returns nothing,
**check which provider `OPENAI_BASE_URL` points at before debugging
anything else.**

---

## 6. Resume and outputs

- Checkpoint via `ExpandCheckpoint`; each step logs `SKIP (checkpoint)`
  when already done. Re-running a finished market is near-free and exits
  with `status=done`.
- Output directory: `default_output_dir(query, country)` unless `out_dir`
  is passed. Internal state lives under `_pipeline/` there.
- Final artifacts: wrap-formatted FINAL `.xlsx` + audit JSON.

---

## 7. Things that will bite you

**JSON truncation is expected, not exceptional.** `_repair_json()` exists
because models truncate mid-string on long responses. `gpt_recall_companies()`
deliberately pages in *small* batches for the same reason. If you increase
page sizes, you will reintroduce truncation.

**Never invent values to fill a column.** `_ensure_complete_row()` keeps
only evidenced values and explicitly refuses to synthesise
`MARKET_DEFAULTS` / `known_*` text. This is a hard project rule, learned
from a real incident: a prompt that demanded every field be populated
produced 94% formula-generated `info@<domain>` emails and `/in/<name-slug>`
LinkedIn URLs across 493 companies — indistinguishable from real data
downstream. A blank cell is correct; a plausible invention is not.
`_looks_generic_default()` is the gate for this.

**A column being empty is a signal, not a defect to patch over.**
`_row_needs_web_refill()` decides when firmographics are too weak and
another pass is warranted — use it rather than lowering the evidence bar.

**Cost control lives in step ordering.** `fill_backend="search"` runs the
free stack first. Switching to model-first is the single fastest way to
make a run expensive.

---

## 8. Relationship to the other pipeline

Two distinct systems in this repo; don't confuse them.

| | `chatgpt_expand` | Google AI Mode market research |
|---|---|---|
| Backend | OpenAI SDK (`web_search`) | Google AI Mode via Patchright |
| Shape | Discovery + fill + verify in one run | Discovery, then a separate verify pass |
| Scale limit | API spend / rate limits | ~200-250 queries/day **per IP** |
| Entry | `run_chatgpt_expand_batch.py` | `run_fire_market.py`, `supervisor_all.py`, etc. |
| Verification | Steps 4, 6, 7 inline | `verify_distributors_ai_mode.py` |

The Google AI Mode side has a hard per-IP daily ceiling that no amount of
parallelism raises — see
`memory/ai-mode-has-a-cumulative-per-ip-daily-limit.md`. `chatgpt_expand`
is bounded by API cost instead, which is why it is the better choice for
very large market lists.
