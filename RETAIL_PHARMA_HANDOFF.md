# Retail + Pharma leadership lookup — handoff

Context for Claude picking this up on another machine. Written 2026-10-07.

Everything below was produced by the Google AI Mode scraper in this repo. Two
workbooks exist: global/HQ leadership and India-specific leadership.

---

## 1. What was asked for

Find leadership contacts for the companies in two spreadsheets:

- `Retail_companies_for_email_search - 5 Oct 2026.xlsx` — sheet `Retail (442)`
- `Pharma_companies_for_email_search - 5 Oct 2026.xlsx` — sheet `Pharma (174)`

Main sheets only; the `Not included` sheets were explicitly excluded. The two
files overlap by 7 companies, so the working set is **609 unique companies**.

Designations requested, verbatim:

> chro, ceo, founder, owner, md, co founder, co owner and if no designation
> details come those are mention then you find add if possible (ANY HEAD OF
> THESE COMPANIES, SO ATLEAST SOME DATA COMES)

That last clause became the `OTHER HEAD` field: the most senior person at the
company, carrying their real title (Chairman, President, Director, Proprietor,
GM...) in a separate `OTHER HEAD TITLE` column. It exists so no company comes
back completely blank.

A later request added an India-specific pass: **India person only**, not the
global one — for a multinational, the India country head rather than the global
CEO. That is a separate workbook, not extra columns on the first.

Company list, deduped and ready to use:

    output/retail_pharma/companies_609.txt

---

## 2. Current results

| | Global / HQ | India only |
|---|---|---|
| Total names | **2,461** | **2,302** |
| Companies with ≥1 name | 607 / 609 | 605 / 609 |
| Websites | 608 / 609 | 608 / 609 |
| Emails | 47 | 62 |
| Names rejected as unsupported | **0** | **0** |

Per designation:

| Designation | Global | India |
|---|---|---|
| OTHER HEAD | 600 | 586 |
| CEO | 486 | 404 |
| Founder | 466 | 411 |
| MD | 381 | 389 |
| Co-founder | 223 | 196 |
| CHRO | 169 | 174 |
| Owner | 114 | 107 |
| Co-owner | 22 | 35 |

**The India file having fewer names is correct, not a shortfall.** Where a
company has no separate India leadership, the prompt leaves the field blank
rather than substituting the global person — which is what "India only" means.

### Deliverables

    output/retail_pharma/backup/retail_pharma_global_20261005_143025.xlsx   <- global
    output/retail_pharma_india/retail_pharma_india.xlsx                     <- India

Matching `.json` files sit beside each; those are the checkpoints every script
reads and writes. The Excel is always regenerated from the JSON, never edited
directly.

There is also `output/retail_pharma/retail_pharma.xlsx`. It is the global file
from before one website was backfilled — identical apart from the HyperCITY
cell. **Use the `backup/` copy.** (The name "backup" is the user's word for it;
it is the current global deliverable, not a stale copy.)

### Known gaps, already chased

- Never resolved, attempted 3×: **Fresher Than Fresh**, **NBCC** (global);
  plus **Al-Futtaim Group**, **Kin Marché**, **BNB SARL** (India — likely no
  India entity at all, so blank is the honest answer).
- NBCC is the one company with no website. Its real site is `nbccindia.in`, but
  that came from model knowledge rather than a scraper result, so it was left
  blank pending the user's say-so.
- **Two rows need human eyes.** One-word company names resolved to the wrong
  firms: `Sterling` → Sterling Check (US background screening) and `Vertex` →
  Vertex Pharmaceuticals (US biotech). The names are real; the companies are
  probably wrong for an India retail/pharma list. Flagged to the user, not yet
  corrected.

---

## 3. Scripts and how to run them

All take `--auto-reset` and resume from their checkpoint, so re-running after an
interruption only queries what is still missing.

| Script | Purpose |
|---|---|
| `run_retail_pharma.py` | Global pass, one worker. Defines `ROLES` and `write_excel`; the others import from it. |
| `run_retail_pharma_parallel.py` | Launches N workers over the global pass. |
| `merge_retail_pharma_shards.py` | Folds worker shards back into one checkpoint + Excel. |
| `run_retail_pharma_gaps.py` | **Second pass.** Asks each company only for its missing designations. |
| `run_retail_pharma_india.py` | India pass, writes the separate India workbook. |
| `run_retail_pharma_india_gaps.py` | India second pass. |

### Typical sequence

Broad pass, 10 workers:

```powershell
.\.venv\Scripts\python.exe run_retail_pharma_parallel.py `
  --companies-file output\retail_pharma\companies_609.txt `
  --workers 10 --batch-size 5

.\.venv\Scripts\python.exe merge_retail_pharma_shards.py `
  --companies-file output\retail_pharma\companies_609.txt
```

Then the gap pass — this is where most of the data comes from (see §4):

```powershell
# one worker per shard; shard files are written by the parallel script
.\.venv\Scripts\python.exe run_retail_pharma_gaps.py `
  --companies-file output\retail_pharma\gaps\shard_0.txt `
  --source output\retail_pharma\retail_pharma.json `
  --checkpoint output\retail_pharma\gaps\g_0.json `
  --roles "CEO,MD,Founder,CHRO,Owner,Co-founder,Co-owner,OTHER HEAD" `
  --batch-size 6 --auto-reset
```

Every script supports `--dry-run`, which prints the exact prompt and its
character count without opening a browser. **Always dry-run after editing a
prompt** — see the length constraint in §4.

### Launching workers in parallel

Each worker needs its own Chrome profile or they fight over the profile lock:

```bash
for w in 0 1 2 3 4 5 6 7 8 9; do
  GOOGLE_AI_MODE_PROFILE_DIR="<repo>/data/ai_mode_rp$w" \
  GOOGLE_AI_MODE_ENABLED=true PYTHONIOENCODING=utf-8 \
  nohup ./.venv/Scripts/python.exe <script> ... > "logs/w$w.log" 2>&1 &
  sleep 8   # stagger, so 10 browsers do not start at once
done
```

`captcha-raptor` is staged per worker at `<profile>_extensions/`. The parallel
script does this automatically; a hand-launched worker needs the directory to
already exist (copy it from another worker's).

---

## 4. Hard-won details — read before changing anything

**Keep prompts under ~800 characters.** AI Mode echoes the prompt back *twice*
in the page transcript and the reply is capped a few thousand characters. A
1,022-char prompt once left ~250 characters for actual data and silently
truncated the JSON, abandoning ~300 companies as "unparseable" even though
Google had answered correctly. Current prompts: 714 (global), 782 (India).

**The gap pass is worth more than the broad pass.** Asking for all 8
designations at once makes the model spend its output budget on the roles it
finds easily. Asking each company only for what it is missing — grouped so
companies with identical gaps share one query — added **1,000 names to the
global file (+69%)** and **864 to the India file (+60%)**, with zero rejections.
If coverage needs lifting again, run another gap pass rather than re-running
the broad one.

**8–10 parallel workers is the working range.** At 15, three workers hit 5
CAPTCHAs each and cleared none. At 8–10 every CAPTCHA cleared. Across these
runs: 19 CAPTCHAs total, all solved, **zero rate limits**.

**Separate profiles get separate quotas.** An earlier prediction that parallel
workers would exhaust a shared per-IP quota turned out to be wrong — 10 and 15
worker runs saw no rate limiting at all. `run_retail_pharma_parallel.py` still
carries a stale docstring warning about this; ignore it.

**Never run a diagnostic browser while a run is active** — it collides on the
profile lock. Stop the run first.

**Batch size scales inversely with role count.** 8 companies for 3 roles,
6 for 4, 5 for 8 roles. Too large and the reply truncates mid-JSON.

### Anti-fabrication — do not weaken this

Every name is written back only if it appears **verbatim** in the model's own
reply; anything else is dropped and counted as `rejected`. This exists because
an early run invented "Rukam Sitara" as a company's CEO — a name that appeared
nowhere in the source, apparently assembled from a LinkedIn slug and an
unrelated investor's name. The user caught it.

The user's standing rule: **real data or a blank cell, never a plausible
guess.** Prompts say "a wrong name is worse than a blank" for exactly this
reason. All passes here reported 0 rejections, which means the model stayed
grounded, not that the check is inert.

Relatedly: **sparse columns are usually correct, not incomplete.** Co-owner at
22 and Owner at 114 reflect reality — listed companies have shareholders, not
owners, and a 40-person pharma firm has no CHRO. Pushing harder on those
columns is where fabrication risk starts.

### Emails are genuinely scarce

47 global, 62 India. AI Mode almost never surfaces published personal email
addresses; this is expected, not a failure. Names were the stated priority.
If emails are wanted for these 2,461 names, SalesQL or Apollo against the
names already collected is the realistic path — this repo has
`run_salesql_ceo_cto.py`, `run_salesql_reveal.py` and `run_apollo_contacts.py`
from earlier work. Note: SalesQL *search* is free; only email reveal costs
credits.

---

## 5. Environment

Required in `.env`:

```
GOOGLE_AI_MODE_ENABLED=true
GOOGLE_AI_MODE_HEADLESS=false
GOOGLE_AI_MODE_BROWSER=chromium
GOOGLE_AI_MODE_EXTENSIONS=<path>\chrome-scraper-extensions\captcha-raptor\extension
GOOGLE_AI_MODE_QUERY_DELAY=0
GOOGLE_AI_MODE_MAX_RESETS=40
```

`GOOGLE_AI_MODE_BROWSER` must be **chromium**, not chrome — captcha-raptor only
loads under the bundled Chromium build. On a new machine the extension path will
differ; update it before running.

Profile reset (`--auto-reset`) must clear **both** the Chrome and Chromium
profile directories. An early version cleared only Chrome, so the run's 230MB
of Chromium cookies survived and the rate limit persisted. `reset_profile()` in
`run_retail_pharma.py` handles both.

---

## 6. Working style the user expects

- **The user runs long pipelines, not Claude** — unless they say "you run".
  They did say it here, so these were run directly.
- Diagnose before fixing. Root cause by measurement, not inference.
- An empty result and a failed result are different things; distinguish them
  before adding retry logic.
- Report honestly: if something was skipped or came back wrong, say so plainly
  with the numbers.
