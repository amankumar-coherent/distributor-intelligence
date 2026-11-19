# How the Distributor Intelligence Pipeline Works

This is the pipeline used for every market delivered since October 2026 (Brazil RO, Antioxidants, Lactoferrin, Geo, H2, Data Center, Empty Capsules, MENA Magnesium Glycinate, U.S. Entresto, India Spices, Soft Serve, Vacuum Insulation Panels, U.S. 503B Compounding Pharmacies).

You give it a market (definition, inclusion, exclusion, sample report, company type, region, cap and number of browsers). It produces a checked list of companies as **one polished Excel file and one JSON file**.

---

## 1. The big picture

```
 Market brief (definition, exclusions, sample report PDF)
        |
        v
 [0] SET UP      market module + runner + launch scripts, preflight check
        |
        v
 [1] DISCOVER    10 browsers ask Google AI Mode for companies, region by region
        |
        v
 [2] VERIFY      10 browsers re-check every company: KEEP or DROP
        |        (steps 1-2 repeat, round after round, until the cap is
        |         reached or every region has been searched)
        v
 [3] STATUS      10 browsers check whether each company is active, closed
     CHECK       or acquired -> "ABC (acquired by XYZ)"
        |
        v
 [4] BUILD       evidence gate, duplicate merge, flagged list,
                 polished Excel + JSON
        |
        v
 [5] AUDIT       automatic final checks (duplicates, region, JSON = Excel...)
        |
        v
 [6] REVIEW      human review of acquired-by results and flagged companies,
                 fixes go in overrides.json, then rebuild
```

The search engine is **Google AI Mode**. It is driven through real Chromium browsers (patchright with the captcha-raptor extension), each with its own profile. The API key in `.env` (DeepSeek) is used only as a fallback when AI Mode refuses to answer. `.env` is never committed.

---

## 2. Step 0: Setting up a new market

Each market (or group of markets) is a **batch** with a short name, for example `vip`, `ent`, `cmp`. A batch needs these files:

| File | What it holds |
|---|---|
| `<x>_markets.py` | The market itself (see below) |
| `run_<x>_market.py` | Runner: one browser searching one region (copied from an existing one) |
| `market_batches.py` | Registration: `BATCHES` + `_DEFINITION_MODULES` |
| `run_fire_10_browsers.py` | Registration in `_BATCHES` (runner, profile prefix, log folder, regions) |
| `scripts/audit_final.py` | `COUNTRY_GATE` entry if the market is limited to certain countries |
| `logs\run_<x>_batch.ps1` | Starts the run |
| `logs\run_<x>_watchdog.ps1` | Starts the watchdog |

### The market module (`<x>_markets.py`)

| Item | Purpose |
|---|---|
| `MARKETS_X` | `(slug, market name, product scope, kind)`. The **scope** is the short text sent to AI Mode with every search: what the product is, which company type to return, what to exclude. |
| **kind** | The company rule: `distributor` (resells others' products, makers excluded), `hybrid` (distributors + channel partners / integrators), `service` (companies that perform a service, e.g. 503B outsourcing facilities), `software`. |
| `X_REGIONS` | The regions searched one by one: countries (VIP: 34), U.S. states (Entresto, 503B) or Indian states (Spices). |
| Country restriction | `discovery.set_country_restriction(...)` in the runner: only companies headquartered there are accepted. |
| `DEFINITIONS` | The full market definition. It is printed in **row 2** of the final Excel. |
| `VERIFY_RULES` | One `KEEP ... / DROP ...` rule used by the verifier (kept under ~1,000 characters, or AI Mode stops answering). |
| `EVIDENCE` | A pattern (regex) of product words. A verified company whose profile never mentions them goes to the flagged list. **Test it on the verified list before delivery.** |

### Preflight check

AI Mode receives the whole prompt inside one Google URL, and **the URL must stay under 7,800 encoded characters**. A longer prompt gets cut and AI Mode answers in prose, which yields 0 companies (this happened on VIP, 2026-10-25). `scripts/preflight_prompt.py` measures the real URL for every region. The launcher **refuses to start** if any region is over the limit.

---

## 3. Step 1: Discovery (finding companies)

`run_fire_10_browsers.py --batch <x> --rolling --workers-per-market 10 --target <cap> --status-workers 10`

- **10 browsers** run at once. Each takes a different region (country / state) and has its own Chromium profile (`data/<x>_profile_<slug>_wN`).
- Each browser asks AI Mode in rounds: *"give me 10 more companies of this market in <region>, not these ones: …"*. It keeps asking until the region runs dry (empty answers twice in a row).
- Every answer is checked before a company is accepted:
  - the company's verdict must match the market's **kind** (e.g. `distributor`, or `service_provider` for service markets);
  - the HQ must be in an allowed country;
  - it must not be a duplicate (same name, or same website as a company already found).
- Accepted companies are saved in `data/market_research/<market>.json`. Searched regions are saved in `data/region_coverage/<x>_profile_<slug>.json`, so a restart never searches them twice.
- After each round the found companies are exported to `output/market_research/<market>.xlsx`.

Each company row has: Company, Website, Founded, Headquarters, Geography, Operational Presence, Ownership, Employees, Core Categories, Specialty Focus, Key Products, Online/Offline, Contact Person, Role, Email, LinkedIn, Office No., Office Website, Office Email and Country Code.

---

## 4. Step 2: Verification (KEEP or DROP)

Discovery alone is not trusted. After every discovery round, **every new company is re-checked** in a separate pass:

- `scripts/parallel_verify.py` splits the unchecked companies over **10 browsers**. Each asks AI Mode about 8 companies at a time (`VERIFY_CHUNK=8`), using the market's `VERIFY_RULES`.
- For each company AI Mode returns: in market?, right company type?, fits criteria?, closed?, confidence. Only a company that passes all of them, with enough confidence, is **KEPT**.
- Verdicts are merged into `output/market_research/<market>_verify_checkpoint.json`. Kept companies are written to `<market>_verified_distributors.xlsx`.
- If a verify run is cut off, it is retried once and resumes from its checkpoint.

The launcher then logs a line such as
`cmp_503b: Missouri, Idaho, ... +8 found, verified 142/150`
and starts the next round with the next 10 regions. It stops when:
- the verified count reaches the **cap**, or
- every region has been searched, or
- several rounds in a row find nothing new.

---

## 5. Step 3: Status / ownership check

`scripts/status_check_folder.py` (run automatically by the launcher's `finalize`):

- The verified Excel is copied to `output/final/<x>/<x>_verified/`.
- **10 browsers** ask AI Mode about each company (identified by its website): **active / closed / absorbed / unknown**, who acquired it, the year, and one sentence of evidence.
- Answers are saved in `data/status_check_v2/<x>_verified/shard_N.json` (resumable).

What happens to each company:

| Result | Action |
|---|---|
| Active, not acquired | **Kept** as is |
| Acquired, still trading | **Renamed** `ABC (acquired by XYZ)` |
| Closed or absorbed (with evidence) | **Removed** (listed in `status_check_report.xlsx`) |
| Suspicious ownership claim (two different owners, deal only "agreed") | Moved to **flagged_for_review.xlsx** |

---

## 6. Step 4: Building the final files

Still inside `status_check_folder.py --build`:

1. **Market definition** is written into row 2 of the sheet.
2. **Evidence gate**: a company whose profile never matches the market's `EVIDENCE` words goes to the flagged list, not the main Excel.
3. **Duplicate merge**: the same company listed twice is kept once (the fullest row, or the one already labelled "acquired by"). Two rows count as the same company when they have:
   - the same website domain and country;
   - the same name and city;
   - the same name in the same country (legal forms like Ltd, GmbH, S.r.l. ignored);
   - the same name and the same website.

   A flagged row whose twin is already in the main list counts as a duplicate, not a flag.
4. **Polish**: filters, frozen header, column widths, styling.
5. Output in `output/final/<x>/<x> distributors (status checked)/`:

| File | Content |
|---|---|
| `<slug>_verified_distributors.xlsx` | **The deliverable**: checked companies only |
| `<slug>_verified_distributors.json` | Same data as JSON |
| `flagged_for_review.xlsx` | Companies kept out of the main list, each with the reason |
| `status_check_report.xlsx` | Every rename, removal and duplicate merge, with evidence |

No zip file is made. The deliverable is Excel + JSON only.

---

## 7. Step 5: Automatic audit

`scripts/audit_final.py --batch <x>` runs at the end of finalize and prints `AUDIT <x>: OK` or a list of problems. It checks:

- row 2 holds the market definition;
- no blank company names or websites;
- no duplicates left;
- no flagged company also in the main list;
- every headquarters is in an allowed country (`COUNTRY_GATE`);
- the JSON has exactly the same rows as the Excel;
- no zip file.

---

## 8. Step 6: Human review and corrections

Before delivery the results are reviewed:

- **Acquired-by labels**: is the owner real, and was the deal completed? (e.g. "acquired by New leadership team" is not an owner.)
- **Removals**: being bought is **not** "absorbed". A company that still trades under its own name stays, with an "(acquired by …)" label.
- **Flagged list**: genuine companies wrongly flagged usually mean the `EVIDENCE` pattern is too narrow. Fix the pattern, not the data.
- **Remaining duplicates**: other sites or arms of one company with different websites.

Corrections go into `data/status_check_v2/<x>_verified/overrides.json`, keyed like the verdicts (`companyname|domain`):

```json
{
 "prorxpharma|prorxpharma.com": {"status": "active", "acquired_by": "", "note": "management change only", "flag": false},
 "fagronnorthamerica|fagron.us": {"note": "not a 503B facility - bulk API supplier"}
}
```

- `"flag": false` = accepted correction (the company stays in the main list with the corrected data).
- A `note` without `"flag": false` = the company goes to the flagged list with that note.
- `"status": "absorbed"` + evidence = removed (used for duplicates the automatic merge missed).

Write this file with an editor or the Write tool. If Windows PowerShell saves it, it adds a byte-order mark; the loader now handles that, but an editor is safer. Then rebuild (no browsers needed) and re-audit:

```
python scripts\status_check_folder.py --input output\final\<x>\<x>_verified --output "output\final\<x>\<x> distributors (status checked)" --build
python scripts\audit_final.py --batch <x>
```

Overrides survive re-runs because the original AI Mode answers are never edited.

---

## 9. Running it

### Start

Always start through WMI with a **headless console**. A normal PowerShell start opens a Windows Terminal window, and closing that window kills the run.

```powershell
foreach ($f in "run_<x>_batch.ps1", "run_<x>_watchdog.ps1") {
  Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{
    CommandLine = "conhost.exe --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"C:\Users\aman.kumar\distributor-intelligence\logs\$f`""
    CurrentDirectory = "C:\Users\aman.kumar\distributor-intelligence" }
}
```

The batch `.ps1` sets `VERIFY_CHUNK=8`, `PARALLEL_VERIFY=10` and `GOOGLE_AI_MODE_MAX_RESETS=10`, then runs the launcher. All output goes to `logs\run_<x>_batch.log`.

### Watchdog (`scripts/batch_watchdog.py`)

Every 10 minutes it:
- writes found / verified counts to `logs\<x>_progress.log`;
- restarts the launcher if it died before the market finished (e.g. after a PC restart);
- kills and resets browsers that have been silent 25+ minutes (stuck on a CAPTCHA or block page);
- runs the final status check itself if the launcher did not.

### Checking progress

| Question | Where to look |
|---|---|
| How many verified? | `logs\run_<x>_batch.log` — one line per round: `+N found, verified V/cap` |
| Timeline | `logs\<x>_progress.log` |
| What a browser is doing right now | `logs_<x>\<slug>.log` |
| Verification detail | `logs_<x>\<slug>_parallel_verify.log` |
| Finished? | `FINALIZE <x>: done` and `AUDIT <x>: OK` in the batch log |

### Resuming

Everything is checkpointed: found companies, searched regions, verify verdicts and status answers. After a crash or PC restart, start the same `.ps1` again and it continues where it stopped. If the PC restarted, check whether an export xlsx was corrupted mid-write (`BadZipFile`); re-create it with `python market_batches.py --only <slug>`.

---

## 10. Typical numbers

| Market | Found | Verified | Final list |
|---|---|---|---|
| North America & Europe Vacuum Insulation Panels | 615 | 422 | 331 |
| U.S. 503B Compounding Pharmacies | 193 | 143 | 100 |

Usually 25–35% of found companies fail verification, and 10–25% of verified rows are duplicates or other sites of the same company. A run with 10 browsers takes about **2.5–4 hours** from start to final files.

---

## 11. Known pitfalls (already handled in the pipeline)

| Problem | Protection |
|---|---|
| Prompt URL over 7,800 characters -> 0 companies | Preflight check refuses to launch |
| Verifier criteria too long -> empty answers | Warning above ~1,000 characters; `VERIFY_RULES` kept short |
| Distributor rule drops service-type companies (e.g. 503B facilities) | Choose `kind = "service"`; verification then skips the distributor test |
| Evidence pattern too narrow -> genuine companies flagged | Test `EVIDENCE` on the verified list; widen the pattern |
| Browser stuck on CAPTCHA | captcha-raptor extension, profile reset, watchdog kills silent browsers |
| AI Mode refuses a status chunk | `STATUS_CHUNK=1` or a plain-English prompt |
| PC restarts | Checkpoints everywhere; watchdog or a re-launch resumes |
| Duplicates under different spellings | Name normalisation + domain / city / country rules; flagged twins merged |
| Owner label outdated ("acquired by X", later sold to Y) | Reviewer override replaces the label |

---

## 12. What is never pushed to GitHub

- `.env` (API key)
- browser profiles (`data/*_profile_*`)
- `data/_archived*`, `ai_mode_*` folders
- `data/india_leaders`

The **Food batch** (75 ProFood India markets) is paused. It resumes only on request, via `logs\run_food_batch.ps1` + `run_food_watchdog.ps1`; see `RESUME_FOOD.md`.
