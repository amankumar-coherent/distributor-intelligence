# Brands_V2

A pipeline for building a global database of nutraceutical/supplement brands (name, official website, official address) for every keyword in a canonical ingredient/product list, using Google's AI Overview / AI Mode as the data source.

## What this project does

1. Start from a list of ~1,246 nutraceutical keywords ([nutraceutical_canonical_names.csv](nutraceutical_canonical_names.csv)) — ingredient and product names like "Vitamin A", "Retinol", etc.
2. For each keyword, ask Google ("list of 100 {keyword} brands globally...") via automated browser scraping.
3. Parse the AI-generated answer text into structured rows: `keyword, brand_name, website, address`.
4. Append results to a checkpointed CSV so long runs can be stopped/resumed without re-querying or duplicating work.

There are two independent scraping strategies, and they can be used separately or chained as primary/fallback.

## Architecture

```
nutraceutical_canonical_names.csv  (1,246 keywords)
              │
              ▼
   scraper/scrape_ai_mode.py  ──(primary)──▶  Playwright → Chrome (persistent profile)
              │                                  → google.com/search?...&udm=50 (AI Mode)
              │                                  → extracts answer text + link hrefs directly from the DOM
              │ falls back to
              ▼
   scraper/scrape_brands.py   ──(fallback / standalone)──▶  google-ai-scraper HTTP API (localhost:15551)
              │                                               → Chrome/Firefox extension relay
              │                                               → classic Google AI Overview (no udm=50)
              ▼
   shared text parser (scrape_brands.parse_brands)
              │
              ▼
   scraper/out/*.csv  (keyword, brand_name, website, address)
   + *.done.txt (checkpoint)  + *.weak.txt (low-yield, retry queue)  + *.errors.log / *.sources.log
```

### Component 1 — `scraper/scrape_ai_mode.py` (primary scraper)

Drives a real Chrome browser directly via Playwright to Google's **AI Mode** (`udm=50`), which returns richer results (real URLs, full addresses, up to ~100 brands per query) than the classic AI Overview. Chrome's omnibox normally intercepts `udm=50` into an unreadable internal page, so this script navigates with Playwright to bypass that.

- Uses a **persistent, cloned Chrome profile** (`scraper/.pw-chrome-profile`) with real cookies/history, because a brand-new empty automated profile triggers an immediate CAPTCHA.
- Extracts the rendered answer as text, rewriting `<a href>` links inline as `text (https://url)` so the shared parser can recover URLs even when the visible link label is just "Official Website".
- Detects Google's CAPTCHA/"unusual traffic" interstitial and pauses for manual solving.
- If AI Mode returns too few brands (or fails/CAPTCHAs), automatically **falls back** to `scrape_brands.py`'s HTTP pipeline (classic AI Overview) and keeps whichever result has more brands.

### Component 2 — `scraper/scrape_brands.py` (HTTP fallback / standalone scraper)

Talks to a separately-running local service — the **google-ai-scraper** backend — over plain HTTP (`GET /ask?q=...`), rather than driving a browser itself.

- Contains the shared `parse_brands()` text parser used by both scripts (handles multiple observed AI answer formats: labeled `Website:`/`Address:` blocks, pipe-delimited tables, unlabeled trailing addresses).
- Can be run standalone if you don't need AI Mode / Playwright — only requires the backend service (below) and its browser extension to be running.

### Component 3 — `chrome-scraper-extensions/google-ai-scraper` (the HTTP backend used by Component 2)

A separate, self-contained sub-project (its own git repo) that Component 2 depends on. It relays queries to a real browser tab so Google renders the AI Overview naturally, without stealing window focus:

```
scrape_brands.py (HTTP client)
        ↕ GET /ask?q=...
   FastAPI backend (:15551)
        ↕ polling (1.5s)
 Chrome/Firefox extension
        → opens a background tab → scrapes AI Overview → posts result back
```

See [chrome-scraper-extensions/google-ai-scraper/README.md](chrome-scraper-extensions/google-ai-scraper/README.md) and its [CLAUDE.md](chrome-scraper-extensions/google-ai-scraper/CLAUDE.md) for full internals (server endpoints, extension message flow, DOM extraction heuristics, MCP server, troubleshooting).

### Output & checkpointing (both scrapers, same conventions)

For an `--out out/X.csv` run, four files are maintained side by side:

| File | Purpose |
|---|---|
| `X.csv` | Accumulated results: `keyword, brand_name, website, address` |
| `X.done.txt` | One keyword per line — completed keywords, skipped on re-run |
| `X.weak.txt` | Keywords that returned fewer than `MIN_ACCEPTABLE_BRANDS` (10) brands — not written to CSV, not marked done, retryable via `--retry-weak` |
| `X.errors.log` (scrape_brands.py) / `X.sources.log` (scrape_ai_mode.py) | Per-keyword error messages, or per-keyword source used (`ai_mode` / `ai_overview_fallback` / `ai_mode_weak`) + brand count |

Debug artifacts (raw markdown/HTML/screenshots for low-yield or empty responses) are written to `scraper/out/raw_debug/` and `scraper/out/raw_debug_aimode/`.

## Prerequisites

- **Python 3.13+**
- **Google Chrome** installed (the `channel="chrome"` Playwright launch requires the real Chrome, not just Chromium)
- For `scrape_ai_mode.py`: [Playwright for Python](https://playwright.dev/python/)
- For `scrape_brands.py` (used standalone or as fallback): the `google-ai-scraper` backend + browser extension running (see below)
- [uv](https://docs.astral.sh/uv/getting-started/installation/) — used to run the `google-ai-scraper` backend

## Setup

### 1. Install Python dependencies for the scraper scripts

```bash
pip install playwright
playwright install chromium
```

(`scrape_brands.py` itself uses only the Python standard library — no extra install needed if you only run it standalone.)

### 2. Set up the `google-ai-scraper` backend (required for `scrape_brands.py`, and as the fallback path for `scrape_ai_mode.py`)

```bash
cd chrome-scraper-extensions/google-ai-scraper/server
uv sync
uv run google-ai-scraper --backend --port 15551
```

Then install its browser extension and point it at `http://localhost:15551` — see [chrome-scraper-extensions/google-ai-scraper/README.md](chrome-scraper-extensions/google-ai-scraper/README.md) for the Chrome Web Store link / Firefox local-load steps.

Verify it's up:

```bash
curl -s http://localhost:15551/health
```

### 3. First-run Chrome profile for AI Mode (`scrape_ai_mode.py` only)

The first time you run `scrape_ai_mode.py`, Playwright launches a persistent Chrome profile at `scraper/.pw-chrome-profile`. If Google shows a CAPTCHA on early runs, solve it manually in the visible browser window — the script waits and retries automatically. Do **not** run with `--headless` until the profile is "warmed up" (has passed a CAPTCHA at least once), since headless mode can't be solved interactively.

## Running

All commands below are run from the `scraper/` directory.

### Primary path — AI Mode via Playwright (with automatic fallback)

```bash
cd scraper
python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_full.csv
```

Useful flags:

```bash
# Process only a slice of the keyword list
python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_full.csv --start 0 --limit 10

# Adjust delay between keywords (default 6s)
python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_full.csv --sleep 8

# Run headless (only once the Chrome profile is CAPTCHA-warmed)
python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_full.csv --headless

# Re-run only keywords that previously yielded too few brands
python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_full.csv --retry-weak
```

### Fallback / standalone path — classic AI Overview via HTTP backend

Requires the `google-ai-scraper` backend + extension running (Setup step 2) and connected — check with `curl -s http://localhost:15551/health`.

```bash
cd scraper
python scrape_brands.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_test.csv --start 0 --limit 10
```

Same `--sleep`, `--start`, `--limit`, and `--retry-weak` flags are supported.

### Resuming an interrupted run

Both scripts are idempotent by keyword: re-running the same `--out` path skips any keyword already present in `<out>.done.txt`. Just re-run the same command to continue where it left off.

## Repository layout

```
Brands_V2/
├── nutraceutical_canonical_names.csv     # input: ~1,246 keywords to search
├── scraper/
│   ├── scrape_ai_mode.py                 # primary: Playwright-driven AI Mode scraper (+ fallback to scrape_brands)
│   ├── scrape_brands.py                  # HTTP client for the google-ai-scraper backend + shared answer parser
│   ├── .pw-chrome-profile/               # persistent Chrome profile used by Playwright (gitignore-worthy, large)
│   └── out/                              # scrape outputs, checkpoints, logs, debug dumps (gitignore-worthy)
└── chrome-scraper-extensions/
    └── google-ai-scraper/                # separate sub-project: FastAPI backend + Chrome/Firefox extension + MCP server
        ├── server/                       # FastAPI app + MCP server (Python, uv-managed)
        ├── extension/                    # Chrome MV3 extension
        ├── firefox-extension/            # Firefox MV3 add-on
        └── README.md / CLAUDE.md         # full docs for this sub-project
```

## Notes & gotchas

- `MIN_ACCEPTABLE_BRANDS = 10` — any keyword yielding fewer brands than this is treated as a weak/failed result and queued for retry rather than partially written.
- The parser (`scrape_brands.parse_brands`) validates that a brand's "official website" URL actually contains a token from the brand name, to avoid recording an unrelated citation link as the website.
- CAPTCHAs are detected by matching known Google interstitial text ("unusual traffic", `/sorry/index`, etc.); both scripts pause (60s) and retry rather than failing outright.
- `scraper/.pw-chrome-profile` and `scraper/out/` contain large/generated data and should generally not be committed to version control.

## Moving to another machine

If you're copying this project to a new laptop to run it there, don't copy everything — some folders are machine-specific caches/state that must be regenerated locally instead.

### Copy

```
Brands_V2/
├── README.md
├── nutraceutical_canonical_names.csv
├── scraper/
│   ├── scrape_ai_mode.py
│   ├── scrape_brands.py
│   └── out/                              # optional — copy to resume from existing checkpoints/progress
└── chrome-scraper-extensions/
    └── google-ai-scraper/                # everything EXCEPT server/.venv
        ├── server/
        │   ├── google_ai_scraper/
        │   ├── pyproject.toml
        │   ├── uv.lock
        │   └── start-services.sh
        ├── extension/
        ├── firefox-extension/
        ├── README.md
        └── CLAUDE.md
```

### Do NOT copy — regenerate instead

| Path | Why skip it | Regenerated by |
|---|---|---|
| `scraper/.pw-chrome-profile/` | ~1GB live Chrome profile (cookies, history, caches). Machine-specific; copying it can break Playwright's Chrome launch on another machine and defeats the "warm profile avoids CAPTCHA" trick. | Playwright auto-creates a fresh one on first run — solve one CAPTCHA manually to warm it up |
| `scraper/__pycache__/` | Compiled `.pyc` cache | Python regenerates automatically |
| `chrome-scraper-extensions/google-ai-scraper/server/.venv/` | Virtualenv with absolute paths baked in | `uv sync` |
| `chrome-scraper-extensions/.git/` | Git history — only copy if you want version control there | n/a |
| `.claude/` (project root) | Claude Code's local session/settings state | n/a |
| `.playwright-mcp/` | Leftover debug snapshots, not part of the app | n/a |
| any `.venv/`, `.uv/`, `__pycache__/`, `*.pyc` | build/cache artifacts | regenerated automatically |

### Setup on the new machine

```bash
# 1. Python deps for the scraper scripts
pip install playwright
playwright install chromium

# 2. google-ai-scraper backend
cd chrome-scraper-extensions/google-ai-scraper/server
uv sync
uv run google-ai-scraper --backend --port 15551

# 3. Install the Chrome extension (Chrome Web Store link above) and set its
#    server URL to http://localhost:15551

# 4. Run the scraper (first run needs one manual CAPTCHA solve to warm the new profile)
cd ../../../scraper
python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv --out out/brands_full.csv
```
