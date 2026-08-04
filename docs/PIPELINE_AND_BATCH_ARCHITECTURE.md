# Vendor Intelligence Scraper — How It Works

This document explains two things:

1. **How the agent finds and fills in company data for one market** — every step, and exactly where each piece of data comes from.
2. **How the same pipeline scales from 1 market to 10,000 markets at a time** — the batch architecture, and the exact commands to run it.

---

## Part 1 — The single-market pipeline (what happens when you run one market)

Entry point: `run_pipeline(query_context, settings)` in `src/vendor_intel/pipeline/orchestrator.py`.
Input: a market name + country (e.g. `"Global Avocado Oil Market"`, `"global"`).
Output: a list of verified companies with role, products, confidence, and firmographic data, written to CSV/DOCX/XLSX/JSON.

```
Query ("Avocado Oil Market", "global")
   │
   ▼
Phase 1 — Market understanding & query plan      (LLM)
   │
   ▼
Phase 1b — Seed resolution & known-player recall (LLM + search)
   │
   ▼
Phase 2 — Discovery                              (search engines)
   │
   ▼
Phase 2b — Multi-source discovery                (B2B directories, LinkedIn, Owler)
   │
   ▼
Seed merge + entity gate + dedup                 (rule-based filtering)
   │
   ▼
Phase 3 — Enrichment                              (crawl each company's own website)
   │
   ▼
Phase 4 — Classification                          (LLM, per company)
   │
   ▼
Quality export filter, geo filter, dedup          (rule-based)
   │
   ▼
Confirmation pass (verified vs. "not verified")
   │
   ▼
(optional) Post-classification audit             (LLM keep/review/remove)
   │
   ▼
Export: CSV + DOCX + XLSX (schema by market_type) + JSON
```

### Step-by-step, with data sources

**1. Market understanding (`phase1/runner.py` → `stages/a_compiler.py`)**
The market name is sent to the LLM (OpenAI/opencode, `gpt-4o-mini` by default) to produce a **query plan**: market layers, segments, search queries, and an initial list of LLM-recalled seed company names. This is the LLM's own trained knowledge — no web lookup yet. Also generates a **funnel** of 3 themed levels (L0–L2) used later to steer search queries.

**2. Seed / known-player resolution (`pipeline/plan_seeds.py`, `funnel/market_understanding.py`)**
- If the caller supplied curated "must-have" company names (from `queries/seeds/<market_slug>.txt`, auto-matched by keyword overlap with the market name), each name is resolved to an official domain via search and pinned into the run — these are never dropped by later filters just because a crawl failed.
- Separately, `enumerate_market_players()` asks the LLM directly "what real companies do you know in this market?" — again, LLM training knowledge, not a live web search. This purely **widens the candidate pool**; every name still goes through the same enrichment/classification gates as anything else.

**3. Discovery (`phase2/discovery_fast.py`, `clients/search_router.py`)**
This is the first live web data collection. Each search query from the plan is run through the **free search stack**, in order, until enough hits are found:
1. `ddgs` (aggregator library) across backends: **DuckDuckGo, Bing, Brave, Mojeek, Google**
2. Bing HTML (direct scrape) if `ddgs` comes up short
3. SearXNG (self-hosted meta-search, if running locally)
4. Wikipedia (thin fallback, last resort)

Every engine call goes through a shared rate limiter / circuit breaker (see Part 2) so a block on one engine doesn't cascade. Search results are titles + snippets + links; company names and domains are extracted from these.

**4. Multi-source discovery (`pipeline/multi_source_discovery.py`)**
Runs 8 additional targeted searches per market against B2B trade directories and professional/company-profile sites: **Alibaba, Global Sources, Kompass, ThomasNet, TradeIndia, ImportYeti, LinkedIn company pages, Owler company pages** (`site:owler.com/company {market}` etc.). This only *discovers candidates* (name + domain) — it does not pull firmographic details from these sites. Candidates flow into the exact same filtering pipeline as search-engine discoveries; no special trust is given to a directory hit.

**5. Merge, entity gate, dedup (`pipeline/plan_seeds.py`, `pipeline/entity_gate.py`, `utils/domain_corrections.py`)**
All candidates from steps 3–4 are merged with the pinned seeds, deduplicated by domain, and passed through a rule-based "is this actually a company, not a directory/blog/marketplace" filter. Domain corrections/aliasing fixes common issues (e.g. `braskem.com` → `www.braskem.com.br`). An off-market keyword pre-filter drops obvious mismatches before spending a crawl on them.

**6. Enrichment (`enrichment/smart_enrichment.py` → `crawler/smart_crawl.py`)**
This is the second (and main) live web data collection stage: **each surviving candidate's own website is crawled** (headless browser, concurrency-capped to protect host memory — see Part 2). This is where firmographic detail actually comes from today: headquarters, founded year, employee count, products, email, revenue if stated on-site, etc. — all extracted from the company's own pages, not from a third-party database. There is no Owler/LinkedIn *data* pull here — only their own site.

**7. Classification (`intelligence/classifier.py`)**
For every enriched candidate, an LLM call decides: is this company actually relevant to the market, what's its role (manufacturer / supplier / distributor / etc.), what are its key products, and a confidence + quality score. This runs concurrently (bounded by `pipeline_classify_concurrent`) and is wrapped in `asyncio.to_thread` so it doesn't block other concurrent markets.

**8. Quality export filter, geo filter, dedup (`pipeline/quality_export.py`, `pipeline/geo_filter.py`)**
Rule-based: drops low-quality/low-confidence rows (no padding to hit a target count), drops companies clearly headquartered outside the target geography (seeds exempt), dedupes by domain.

**9. Confirmation pass**
Companies that made it this far but whose crawl failed or role couldn't be substantiated are moved to a separate "Not Verified — Could Not Confirm" bucket with a stated reason, instead of silently being dropped or silently kept.

**10. Post-classification audit — *optional, off by default*** (`pipeline/post_classification_audit.py`)
An additional LLM keep/review/remove pass, folded in from a former manual audit script. `remove` verdicts move to rejected, `review` verdicts move to the unverified bucket. Controlled by `settings.pipeline_post_audit_enabled` (default `False` until validated at Pilot Gate 1).

**11. Export (`pipeline/orchestrator.py`, `export/fmcg_landscape.py`)**
Writes CSV (sectioned), DOCX (CEO-readable Word doc), XLSX (column schema chosen by `market_type` — see Part 2), and JSON (full result, all fields, used for re-processing).

### Summary: where data actually comes from

| Data | Source |
|---|---|
| Market structure, segments, initial seed names | LLM training knowledge (no web call) |
| Candidate company names/domains | Live search engines (DuckDuckGo/Bing/Brave/Mojeek/Google/SearXNG/Wikipedia) + B2B directory/LinkedIn/Owler *searches* |
| Firmographics (HQ, founded year, employees, products, email) | Crawling each company's **own website** |
| Relevance / role / confidence | LLM judgment, given the crawled site content |
| Curated must-haves | Hand-maintained files in `queries/seeds/*.txt` |

---

## Part 2 — Scaling to 10,000 markets

Running the same `run_pipeline()` above once per market, one after another, is what the old `run_<market>_market.py` scripts and `run_query.py --file` did. That doesn't survive 10,000 markets: no resumability, no shared rate limiting (free search engines get blocked), no cost tracking, no cross-market caching, no concurrency.

The batch system wraps the exact same `run_pipeline()` in an orchestration layer:

```
run_batch.py --input markets.csv --batch-id X --max-concurrent N --budget-cap $Y
        │
        ▼
data/batch.db (SQLite, WAL mode) ── the shared queue + state + cost ledger
        │
        ▼
batch/orchestrator.py — one process, an asyncio pool of "market slots"
        │
        ├── slot 0 ─┐
        ├── slot 1 ─┼─→ batch/worker.py → run_pipeline() (same as Part 1) → export
        ├── slot 2 ─┤
        └── slot N ─┘
```

### Why one process, not multiple processes, by default

Two existing safety mechanisms only work correctly if all concurrent markets share one Python process:
- `crawler/smart_crawl.py`'s browser-concurrency cap is a module-level `asyncio.Semaphore` (it already caused a real host OOM once when this wasn't respected).
- The search rate limiter / circuit breaker (below) is shared in-process state.

So the default scaling knob is `--max-concurrent` (asyncio slots within one process), not more OS processes. Multiple processes are only introduced if a pilot run proves one process is CPU/RAM-bound — and even then, each additional process is launched with an explicit slice of the concurrency budget so the combined total stays within the proven-safe ceiling.

### What SQLite (`data/batch.db`) tracks

| Table | Purpose |
|---|---|
| `markets` | The queue itself — status (pending → running → done/failed), which worker claimed it, retry count, cost, output file paths |
| `companies` + `company_market_verdicts` | Cross-market company cache, keyed by domain — a company already crawled for market #12 isn't re-crawled for market #400 |
| `search_cache` | Query → results cache with a TTL, so overlapping/regional-variant search queries across markets reuse results instead of re-hitting search engines |
| `search_rate_state` | Token-bucket + circuit-breaker cooldown, per search engine, shared across every concurrent market — a block discovered by market #340 protects markets #341–10,000 |
| `llm_cost_ledger` | Every real LLM call, priced per provider/model, for accurate cost tracking |
| `batches` | One row per batch run: budget cap, paused flag, totals |

**Task claiming** is a single atomic SQL statement — safe for any number of concurrent workers/slots to pull from the same queue without double-claiming:
```sql
UPDATE markets SET status='running', worker_id=?, started_at=?
WHERE id = (SELECT id FROM markets WHERE batch_id=? AND status='pending' ORDER BY priority, id LIMIT 1)
RETURNING *
```

### Resumability

If the process is killed or crashes mid-run: markets already `done` are untouched, and any left `running` by the dead process are automatically reaped back to `pending` the next time `run_batch.py` starts with the same `--batch-id`. No manual cleanup needed for a normal crash/restart.

### Cost control

`run_batch.py --budget-cap X` stops the orchestrator from **claiming new** markets once total real spend (from the LLM cost ledger, not an estimate) hits `X`. In-flight markets finish normally. Raising the cap and re-running the same `--batch-id` resumes a paused batch. A pre-flight estimate is printed before the run starts (calibrated from real run history once enough markets have completed; a static per-cap estimate until then).

### Export schema by market type

`config/export_schemas.yaml` defines column sets per `market_type` (`fmcg` = 30 columns including retail/foodservice/e-commerce channel fields; `general` = 17 universal firmographic columns; unknown types fall back to `general`). `market_type` is set per row in the input CSV.

### How to actually run 10,000 markets

**1. Build the market list as a CSV** (not `.txt` — CSV is what supports `market_type`/`priority`/per-row seeds at this scale):
```csv
query,country,market_type,sections,seeds_file,priority
Global Avocado Oil Market,global,fmcg,,,100
Bio Based Ethylene Market,global,general,,,100
... (10,000 rows)
```

**2. Load + launch (recommended for 10k on 1 PC — bulk mode):**
```
python run_batch.py --input queries/markets_10k.csv --batch-id full-10k \
    --mode bulk --yes
```
Bulk mode (~5–12 min/market): small cap, skips multi-source discovery and player
enumeration, exports CSV/JSON only. Default concurrency is 3 markets at once.
Run `gap_fill_batch.py` or re-process shortlists with `--mode quality` later.

**Alternative tiers:**
```
python run_batch.py --input queries/markets_10k.csv --batch-id full-10k \
    --mode standard --max-concurrent 2 --yes          # pipeline + all exports, no post-enrich
python run_batch.py --input queries/markets_10k.csv --batch-id full-10k \
    --mode quality --max-concurrent 1 --push-neon --yes   # full post_market per market
```
- `--mode` — `bulk` | `standard` | `quality` | `full` (presets cap, core tuning, post steps).
- `--max-concurrent` — how many markets run at once (bulk default: 3).
- `--cap` — override cap tier: `bulk` / `focused` / `standard` / `broad` / `maximum`.
- `--yes` — skip the interactive confirmation (needed for a walk-away run).

**3. Monitor from a second terminal:**
```
python status_batch.py --batch-id full-10k --watch --interval 30
```
Shows pending/running/done/failed counts, companies found, cost so far vs. cap, ETA, and recent failures.

**4. If it stops (crash, Ctrl+C, budget cap hit):** re-run the same command (drop `--input` once the batch is already loaded). It resumes automatically from SQLite state.

**5. Scaling beyond one process** (only if proven necessary by a pilot): launch additional `run_batch.py` processes against the same `--batch-id`, each with an explicit slice of `--max-concurrent`, so the sum stays within the pilot-proven-safe ceiling. The atomic claim query is what makes this safe — no extra coordination needed.

### Staged rollout (why this isn't run at 10,000 on the first try)

Concurrency ceiling, search-engine block rate, and cost-per-market are empirical, not something code review can answer. The rollout is gated:

1. **Smoke test** — 1–2 markets, sequential/low concurrency, to confirm every phase executes end-to-end.
2. **Pilot Gate 1** — ~50 markets (mixed types, including known/hand-audited markets for a quality baseline). Checks: search circuit-breaker trip rate per engine, cost-per-market vs. estimate, classifier quality vs. the manual baseline, wall-clock time and host memory at the chosen concurrency.
3. **Pilot Gate 2** — ~500 markets. Re-checks the same metrics at 10x volume, plus search-cache hit-rate and SQLite contention under load.
4. **Full rollout** — 10,000 markets, using the concurrency/cap/budget numbers validated by the two gates above.
