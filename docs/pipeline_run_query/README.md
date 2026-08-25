# `run_query.py` — Market Discovery + Post-Enrich Pipeline

**Generic pipeline — works for ANY market query.** The market name and
country are just input parameters; nothing in this flow is hardcoded to a
specific vendor, industry, or geography.

See [`run_query_flowchart.pdf`](run_query_flowchart.pdf) /
[`run_query_flowchart.png`](run_query_flowchart.png) for the full diagram.

## What it does

`run_query.py` is the interactive / single-command entry point for building
a channel-partner landscape for one market at a time (or a batch file of
markets, one per line). Given a free-text market name — e.g. `"Atomic Clock
Market"`, `"Industrial Gearbox Market"`, `"Renewable Energy EPC Market"` —
it:

1. Discovers candidate companies for that market via an LLM-driven search
   plan.
2. Enriches each company with firmographic + contact data.
3. Classifies/scores companies into requested sections.
4. Exports CSV / DOCX / JSON (+ optional FMCG Excel).
5. Optionally runs a full **post-enrich pipeline** that fills gaps, verifies
   Founded/HQ, resolves contacts via Apollo, audits market relevance, and
   pushes the result to Neon (Postgres).

## How to run it

```bash
# Single market, interactive follow-up prompt after it finishes
python run_query.py --query "<Market Name>" --country "<Country>"

# Non-interactive, one-shot
python run_query.py --query "<Market Name>"

# Batch: one market per line in a text file
python run_query.py --file queries/my_markets.txt --mode quality

# Resume post-enrich only, on a folder that already has a CSV/JSON
python run_query.py --resume --query "<Market Name>" --push-neon

# Resume discovery: reuse the saved Phase 1/2 output, redo Phase 3/4
python run_query.py --resume-checkpoint --query "<Market Name>"
```

Interactive mode (no `--query`/`--file`) prompts for a query, runs it,
prints a summary, and loops for the next one (or `quit`/`history`).

## Pipeline stages

| Stage | What happens |
|---|---|
| Load & validate `.env` | Checks `LLM_PROVIDER`/API key and `USE_MOCK_DATA=false` before starting |
| Resolve mode/profile/cap | `--mode bulk\|standard\|quality\|full` and `--profile quality\|balanced\|recall` set discovery/enrichment caps |
| Auto-load curated seeds (optional) | If `queries/seeds/<slug>.txt` exists for this market's slug, those must-have companies are merged in — purely a naming convention, not a per-market code branch |
| **Phase 1 — Query Plan** | LLM turns the free-text query + country + sections into a structured search/discovery plan |
| **Phase 2 — Discovery + Entity Gate** | Searches the web/LLM for candidate companies matching the plan; filters out non-entities |
| **Phase 3 — Enrichment** | Pulls website, HQ, founded year, employee count, contacts, etc. for each candidate (capped) |
| **Phase 4 — Classification** | Scores and buckets companies into the requested sections; applies a confidence filter |
| Export | Writes CSV + DOCX + JSON (+ optional 30-column FMCG XLSX) to `output/<slug>/` |
| Channel-Purity Filter | Drops OEM/utility/non-channel rows using an auto-detected channel family for the market |
| **Post-Enrich Pipeline** (unless `--no-post-enrich`) | 8 steps — see below |
| Save session log | Appends the run's stats to `output/<...>/session_log.json` |

### Post-enrich pipeline (steps 1–8)

1. **Gap-fill** — Founded year, HQ, leadership, contact pages
2. **LinkedIn verify** — confirm Founded + HQ
   - 2b. Wikidata fallback for anything LinkedIn missed
   - 2c. Wikipedia infobox fallback for anything still missing
3. **Apollo free search** — `contact_person` only, no paid credits used
4. **Name upgrade** — resolve partial/obfuscated contact names via LinkedIn/Google/RocketReach
5. **Brand/domain correction** — fix company/brand/domain + `Key Brands Represented`
6. **Market-relevance audit** — web search + LLM keep/review/remove pass
7. **Excel refresh** — landscape Excel + a Founded+HQ-filtered Excel
8. **Neon push** — upserts companies/markets/metadata (only if `--push-neon` or `DATABASE_URL` is set)

Every step is **resumable** — re-running with `--resume` on the same market
folder skips finished steps and already-processed companies.

## Key flags

| Flag | Purpose |
|---|---|
| `--query "<name>"` / `--country "<c>"` | The market to run (country defaults to `global`) |
| `--file <path>` | Batch mode — one query per line (`Query \| Country \| Sections`) |
| `--mode bulk\|standard\|quality\|full` | Preset combining cap + whether post-enrich runs |
| `--profile quality\|balanced\|recall` | Discovery precision/recall trade-off |
| `--cap bulk\|focused\|standard\|broad\|maximum` | Result size cap |
| `--sections "A; B; C"` | Force exact output sections/categories |
| `--seeds "Name1; Name2"` / `--seeds-file <path>` | Must-have companies to guarantee in output |
| `--fmcg-xlsx` | Also export the 30-column FMCG landscape Excel |
| `--push-neon` / `--no-push-neon` | Force/skip pushing results to Neon |
| `--no-post-enrich` | Skip the entire 8-step post-enrich pipeline |
| `--no-name-resolve` / `--no-linkedin-verify` / `--no-wikidata-firmographic` / `--no-wikipedia-firmographic` | Skip individual post-enrich sub-steps |
| `--resume` | Resume post-enrich only, on an existing market export folder |
| `--resume-checkpoint` | Resume discovery — reuse Phase 1/2, redo Phase 3/4 |
| `--no-geo-filter` | Keep global players active in the target geography (for country markets with international suppliers) |
| `--shuffle` | Randomize company order within sections |

## Output

```
output/demo/<slug>/
  <slug>.csv
  <slug>.json
  <slug>.docx
  <slug>.xlsx                          (only with --fmcg-xlsx)
  <slug>_channel_purity_dropped.json   (rows removed by the purity filter)
  <slug>_founded_and_hq_v2.xlsx        (post-enrich)
  _neon_filtered_export.json           (post-enrich)
output/demo/session_log.json           (all runs in this session)
```

`<slug>` is derived from the market name + country, so every market gets its
own isolated output folder — the pipeline itself never branches on which
market it is.
