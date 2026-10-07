# Handover: deliverable status-check (active / acquired-by)

Context for picking this work up on another machine. Written for Claude or
any engineer resuming it — it assumes no memory of the original session.

**Status: COMPLETE.** All 50 workbooks processed, 13,065 / 13,065 distinct
companies judged. Nothing is pending. The rest of this file explains what
was built, what the traps are, and how to re-run or extend it.

---

## 1. What this task was

Two folders of finished deliverables, 25 workbooks each:

```
C:\Users\aman.kumar\Desktop\Drone Market      (25 files)
C:\Users\aman.kumar\Desktop\Fire Market       (25 files)
```

The ask: for every company in them, determine

1. **Is it still operating?** Drop confirmed-closed companies.
2. **Was it acquired?** Rename the Company cell to `ABC Ltd (acquired by XYZ)`.

Explicit constraint from the user: **do not change the production pipeline**
for this. It was a one-off pass over already-exported files.

### Decisions the user made (do not silently revisit)

| Question | Decision |
|---|---|
| Company whose status can't be confirmed | **Keep it.** Only drop on positive evidence of closure. |
| Where acquisitions are recorded | **Rename the Company column** to `ABC (acquired by XYZ)`. No extra column. |
| Output location | **New sibling files**, `*_status_checked.xlsx`. Originals never modified. |

---

## 2. The script

`status_check_deliverables.py` (repo root). Standalone; the production
pipeline is untouched.

```bash
python status_check_deliverables.py --status          # progress, no queries
python status_check_deliverables.py                   # single worker
python status_check_deliverables.py --limit 50        # cap queries this run
python status_check_deliverables.py --shard 3 --shards 10   # one parallel worker
```

Per input workbook it writes `<name>_status_checked.xlsx` beside it with
closed companies deleted and acquired companies renamed. Everything else —
the four banner rows, the header on row 5, all 21 columns — is copied
through untouched.

### Workbook layout these files use

- rows 1-4: banner text
- row 5: header
- row 6+: data
- column B: `Company`

### Checkpoints

`data/status_check/<workbook stem>.json`, one per workbook, holding
`{"verdicts": {normalised_name: {...}}}`. Resuming is always safe and free —
already-judged companies are skipped. Delete a checkpoint to force a
re-judge of that workbook.

---

## 3. Running it in parallel (what actually works)

**10 separate processes, one shard each.** Not threads.

```bash
cd /c/Users/aman.kumar/Downloads/vendor-intel
for i in 0 1 2 3 4 5 6 7 8 9; do
  GOOGLE_AI_MODE_PROFILE_DIR="$(pwd)/data/status_profile_w$i" \
  nohup python status_check_deliverables.py --shard $i --shards 10 \
    > "logs_status_w$i.log" 2>&1 &
  sleep 3
done
```

Two non-obvious requirements, both learned the hard way:

- **Threads gain nothing.** `ask()` in
  `src/vendor_intel/scraping/google_ai_mode.py` holds a process-wide
  `_browser_lock` ("one query at a time per page"), so threads inside one
  process serialise. Parallelism must come from separate processes.
- **Each worker needs its own `GOOGLE_AI_MODE_PROFILE_DIR`.** A Chromium
  profile takes an exclusive lock; two processes sharing one produce
  `about:blank` tabs and zero results.

Profiles live at `data/status_profile_w{0..9}` with a sibling
`data/status_profile_w{N}_extensions` holding a copy of the captcha-solver
extension (seeded from `data/ai_mode_chrome_profile_extensions`).

Sharding is **by file, not by company**, deliberately: one checkpoint per
workbook means two workers on the same workbook would overwrite each
other's verdicts. Disjoint file sets mean no locking is needed at all.
Assignment is round-robin (`i % shards`) so large and small workbooks
spread evenly.

Stop everything:

```powershell
Get-CimInstance Win32_Process -Filter "Name='python.exe'" |
  Where-Object { $_.CommandLine -match 'status_check_deliverables' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
```

---

## 4. Four bugs that were found and fixed

These are the whole value of this document. Each was a real failure that
produced wrong or misleading output, and each is easy to reintroduce.

### 4.1 Files written as "complete" while unjudged

**Symptom.** A workbook reported `276 kept, 0 closed, 0 acquired` when zero
of its 276 companies had actually been examined. Unjudged companies pass
through `write_copy()` untouched — which is the *safe* behaviour — so the
output looked finished and the all-zero result read as a finding.

**Why it mattered.** Three files (`drone_analytics`, `drone_batteries`,
`drone_charging`) were reported to the user with fabricated-looking
"no closures found" results. Their real numbers were 2/2, 2/1 and 3/1.

**Fix.** Count unjudged before writing and say so in the log line:
`..., 104 UNJUDGED (re-run to finish)`.

**Lesson.** Finding the first instance of a bug is not the same as knowing
its blast radius. Audit every file, not just the one that surfaced.

### 4.2 A network outage misdiagnosed as the daily rate limit

**Symptom.** Google AI Mode has a measured ~200-250 queries/day/IP ceiling
(see `memory/ai-mode-has-a-cumulative-per-ip-daily-limit.md`). The script
treated "3 consecutive empty answers" as that ceiling. A Wi-Fi drop
produced `ERR_NAME_NOT_RESOLVED` and exactly the same signature.

**Why it mattered.** Had the outage lasted ~30s longer, the run would have
stopped and reported "daily AI Mode limit reached" — a confident wrong
diagnosis, and the user would have waited a day for nothing.

**Fix.** `internet_up()` probes DNS directly before concluding anything. No
network → wait 60s and continue. Network up and still empty → genuine
limit. The probe fails *safe*: if the probe itself errors it returns True,
so a flaky probe can't trap the run in a wait loop.

**Lesson.** One symptom, two causes, opposite correct responses. Ask the
system for ground truth instead of inferring from a correlated signal.

### 4.3 Valid-but-near-empty responses stranded whole chunks

**Symptom.** Three Fire workbooks each ended with ~11-12 unjudged
companies. The first hypothesis — long similar Fire company names breaking
name-matching — was **wrong**.

**What the evidence showed.** The unjudged positions were contiguous and
chunk-aligned: 108-119, 24-35, 36-47. Chunk size is 12. So one *query* per
file returned valid JSON naming 0-1 of its 12 companies, and the rest were
silently dropped. The retry only fired on *no JSON at all*, so a
well-formed but near-empty answer counted as success.

**Fix.** Retry on **coverage**, not parseability: if an answer judges fewer
than half the batch, re-ask and keep whichever response covered more.

**Lesson.** Check positions, not just counts. "11 missing" looked random;
"11 missing, contiguous, chunk-aligned" named the cause immediately.

### 4.4 Phantom "113 companies left" in the status display

**Symptom.** After every file was genuinely complete, `--status` still
reported `113 left`. A background run to "finish" them would have queried
forever and closed nothing.

**Cause.** Verdicts are keyed by *normalised* name, so `ABC Ltd` and
`ABC Ltd.` collapse to one verdict entry. `show_status()` compared
`len(verdicts)` against the raw **row** count, inventing a shortfall.
Proof: duplicate-name collisions across the 50 files total **exactly 113**.

**Fix.** Count distinct normalised companies on both sides.

**Lesson.** Verify a reported gap is real before spending work on it.

---

## 5. Final results

| | |
|---|---|
| Distinct companies judged | **13,065 / 13,065 (100%)** |
| Output files | **50 / 50**, none partial |
| Closed companies removed | **169** |
| Acquisitions annotated | **79** rows |
| Original files | **Unmodified** |

Acquisitions found were substantive, not filler — e.g. Hexagon, Avnet,
Savex Technologies, 63 Moons Technologies, ARKANCE.

**Why 102 acquisitions were detected but only 79 annotated:** companies
that were both acquired *and* confirmed closed are dropped entirely rather
than renamed, so they never reach the output. Correct under the
"active companies only" rule.

---

## 6. If you extend this

- **The same checks now exist in the production pipeline** —
  `verify_distributors_ai_mode.py` asks for `status` and `acquired_by` in
  its verification JSON, so markets verified from now on get this for free.
  This standalone script exists only because the 50 deliverables were
  already finished before that change landed.
- **Prompt length is a hard constraint.** `criteria_prompt()` has a
  measured cliff: ~780-880 chars answers reliably, ~2,096 chars returns an
  empty page for every chunk. Adding instructions there is not free. The
  status/acquired rules pushed it to ~1,050-1,150 and were smoke-tested
  live before being trusted.
- **Never fabricate to fill a field.** A blank or an untouched row is
  correct; a plausible invention is worse than useless because it can't be
  distinguished from real data downstream. `_clean_acquirer()` exists for
  exactly this: it rejects `"none"`/`"N/A"`/`"unknown"` placeholders and
  the company's own name echoed back (which would render
  `ABC (acquired by ABC)`).
