# Verify Logic: What Happens After a Company Comes In

This documents the actual verification pipeline in `verify_distributors_ai_mode.py` —
how a discovered candidate company gets judged KEEP (genuine distributor) or DROP
(everything else) for a given market. It reflects the code and behavior as run and
debugged live against 6 market files (~2,400 companies total), not an idealized design.

Reference implementation: `verify_distributors_ai_mode.py` (repo root).

---

## 1. The shape of the whole thing

```
input Excel (Company, Website, Core Categories, Specialty Focus, evidence cols...)
        ↓
load_rows() — read all rows with a non-empty Company
        ↓
resume from checkpoint offset (0 if first run)
        ↓
split into chunks of 10
        ↓
per chunk: criteria block + JSON company payload → one AI Mode query
        ↓
extract_json() — pull the model's {"results":[...]} out of the rendered page
        ↓
apply_results_to_chunk() → verdict per company (keep: true/false)
        ↓
checkpoint the offset + all verdicts so far, then next chunk
        ↓
after all chunks: write <name>_verified_distributors.xlsx (KEEP rows only)
                   write <name>_verify_log.json (every verdict, KEEP and DROP)
```

Each market in `MARKETS` (the list at the top of the script) is a dict of
`key`, `path`, `market_name`, and `evidence_cols` — the columns from the
source sheet that get sent to the model as supporting evidence.

---

## 2. The criteria block (what the model is told)

`criteria_prompt(market_name)` builds one fixed KEEP/DROP prompt per market,
reused unchanged across every chunk in that market:

```
You are a strict market-landscape verifier. Decide KEEP or DROP for each company.
Market: {market_name}
Required type: Distributor / Wholesaler / Dealer / Reseller / VAD / Importer / Channel Partner
KEEP: companies whose PRIMARY business in this market is distributing, wholesaling,
reselling, importing, or acting as an authorized channel partner for products in
this market. Judge by the company's PRIMARY business — what it IS, not something it
also does.
DROP: pure manufacturers/OEMs/brand owners (even if they also sell direct),
consultancies, market-research firms, news/media outlets and industry associations
(reporting on or advising a market is not operating in it), pure holding companies,
governments, country/city/geo labels, retailers to end consumers only (unless they
are also a recognized wholesale/B2B distributor), and any company that does not
genuinely operate as a distributor in THIS market.
A parent that does not itself distribute in THIS market is a DROP — the operating
subsidiary is the company that belongs here.
If you are not sure the company genuinely operates as a distributor in this market, DROP.
This is a verification pass using existing research data (categories, focus, products).
Use it as evidence.
For EACH company return one result. Never skip a name.
When unsure, fits_criteria=false.
Return compact JSON only, no markdown fences, no commentary:
{"results":[{"Company":"...","in_market":true|false,
"is_distributor":true|false,"fits_criteria":true|false,
"confidence":0-100,"reason":"..."}]}
```

Key properties:
- **"Required type"** states the target explicitly instead of leaving the
  model to infer "distributor" from the market name.
- **"PRIMARY business — what it IS"** stops a manufacturer that also sells
  direct from being kept as if it were a distributor.
- **The parent/subsidiary line** closes the loophole where a holding
  company gets kept because a subsidiary somewhere in the group distributes.
- **"When unsure, fits_criteria=false"** — ambiguity fails closed at the
  model level, before the code-side gate even runs.

Every market currently uses the same distributor-only criteria (this was a
deliberate choice — see §7 on the semiconductor equipment/materials sheet).

---

## 3. Building the per-chunk payload

`build_row_payload()` sends only the fields that inform a KEEP/DROP call:

```json
{"Company": "Acme Distribution Co.",
 "Website": "https://acme.example",
 "Core Categories": "...",
 "Specialty Focus": "...",
 "Key Products / Technologies": "..."}
```

- Each evidence field is capped at `SUMMARY_CAP = 180` chars, so one
  verbose company can't push the rest of its chunk's payload past the URL
  budget.
- `CHUNK_SIZE = 10` companies per query. This is a URL-length constraint,
  not a preference: the whole prompt (criteria block + JSON payload) rides
  in Google's `?q=` search parameter, which 400s past roughly 8,000
  characters once percent-encoded. 10 companies with capped fields keeps
  the encoded URL comfortably under that.

### GLP-1-specific redaction

One market (`glp1`) sets `"genericize_glp1": True`. Before the payload is
built, `_genericize_glp1_terms()` regex-replaces a list of drug brand and
molecule names (Ozempic, Mounjaro, Rybelsus, semaglutide, tirzepatide, …)
with the generic phrase `"a GLP-1 medication"`. This was added after a live
chunk containing these terms was refused by AI Mode's content policy
(`"Something went wrong, and an AI response wasn't generated"`) —
genericizing turned out **not** to be the actual fix (see §7), but the code
is kept because it's harmless and may reduce refusal *rate* even if it
doesn't eliminate it.

---

## 4. Sending the query and reading the answer

```python
text = gam.ask(prompt, retries=2)
```

`gam` is `vendor_intel.scraping.google_ai_mode` — a Patchright-driven,
persistent-profile Chromium session that navigates Google's `udm=50`
(AI Mode) endpoint directly. The verify script retries up to 3 times per
chunk if the answer is empty or a refusal, force-relaunching the browser
(`gam.shutdown()`) between attempts as cheap insurance against a dead
context.

### Extracting the answer from the page

`extract_json()` is the part that took the most iteration to get right.
The rendered AI Mode page **echoes the full prompt** (including the JSON
company payload we sent) before the model's actual answer, so:

- The *first* `{...}` block in the page text is almost always our own
  echoed input, not the verdict.
- Naively grabbing the first brace-balanced JSON object silently parses
  the wrong thing.

The fix: scan every top-level `{...}` block in the page, and return the
**last** one that both parses as JSON and contains a `"results"` list.
That's structurally guaranteed to be the model's answer, not the echo.

### Refusal detection

`gam.is_ai_mode_refusal(text)` checks the page for known refusal phrases
("no response available", "something went wrong, and an AI response
wasn't generated", etc.). This check only looks at the **tail** of the
page text — because the page is long (it echoes a 10-company JSON prompt),
a flat length cap on the whole page missed refusals that were genuinely
short but appeared after thousands of characters of echoed input. Checking
the tail catches the model's real answer/failure message regardless of how
long the input was.

---

## 5. Turning a raw answer into per-company verdicts

`apply_results_to_chunk(chunk, results)`:

1. Matches each result back to its source company by a **normalized name
   key** (`_norm()` — lowercased, non-alphanumerics stripped), not a raw
   string match, so minor formatting differences between what was sent and
   what the model echoed back don't cause a false miss.
2. For every company that *does* get a result, records:
   ```json
   {"Company": "...", "keep": true|false, "reason": "...",
    "confidence": 0-100, "in_market": ..., "is_distributor": ...,
    "fits_criteria": ...}
   ```
3. **Never-skip guard**: any company in the chunk that the model's answer
   didn't mention at all gets a synthetic `keep: false` verdict
   ("no verdict returned by model — dropped fail-closed"). A silent
   omission is treated the same as an explicit DROP, never as an implicit
   KEEP.

### The code-side KEEP gate

The model's own `fits_criteria` flag is **not trusted on its own**.
`should_keep()` re-derives the decision in code:

```python
def should_keep(result):
    if not (result["in_market"] and result["is_distributor"] and result["fits_criteria"]):
        return False
    if result["confidence"] < MIN_CONFIDENCE:  # 70
        return False
    return True
```

All three boolean signals must agree, **and** confidence must clear 70.
A company the model marked `fits_criteria: true` but at confidence 55 is
still dropped — on a run covering thousands of candidates, a
plausible-but-wrong KEEP costs more than a missed real distributor.

---

## 6. Failure handling

### A chunk that returns nothing usable

If `extract_json()` finds no `results` list after 3 attempts:

- If `USE_DEEPSEEK_FALLBACK` is `True`: the identical criteria block and
  payload are sent to DeepSeek (`verify_chunk_via_deepseek()`, reusing the
  `_chat_json` helper from `chatgpt_expand.py`) as a second attempt before
  giving up.
- Either way, if there's still no usable result: **the whole chunk of 10
  is fail-closed dropped** — every company in it gets
  `"unparseable answer from both backends — dropped fail-closed"`.

This is deliberate: an unparseable answer is *no evidence*, and treating
no-evidence as an implicit approval is exactly how non-distributors would
end up in a "verified" list.

`USE_DEEPSEEK_FALLBACK` currently defaults to `False` — per-run choice,
not a permanent architectural decision (see §7 for why it exists at all).

### A crashed market

`main()` wraps each market's `verify_market()` call in a try/except. If one
market crashes outright (not just a bad chunk — an unhandled exception),
the error is logged and the loop moves to the next market. Nothing is
lost: the crashed market's checkpoint already has every chunk completed
before the crash, so a re-run resumes exactly where it stopped.

---

## 7. Checkpointing and resume

After **every chunk**, not just at the end of a market:

```python
save_checkpoint(ckpt_path, {"offset": i + CHUNK_SIZE, "verdicts": verdicts})
```

`<name>_verify_checkpoint.json` sits next to the source file. On the next
run, `load_checkpoint()` reads the last offset and verdict list back in,
and the chunk loop starts from `range(offset, total, CHUNK_SIZE)` — chunks
already checkpointed are never re-queried.

This mattered in practice: the live run was interrupted repeatedly by
browser crashes, IP/session soft-blocks, and manual stops for diagnosis.
Every restart resumed from the exact last completed chunk instead of
re-verifying from zero.

---

## 8. Real failure modes hit in production (and what fixed each)

These are documented because they were each initially mis-diagnosed —
worth knowing if the same symptoms show up again.

| Symptom | Looked like | Actually was | Fix |
|---|---|---|---|
| Process dies with no traceback | Random crash | `UnicodeEncodeError` printing an em-dash/⚠ on Windows' cp1252 console, right as the code was correctly handling a closed browser | Force `sys.stdout.reconfigure(encoding="utf-8")` at the top of the script |
| Whole chunks fail-closed dropped, log shows nothing | AI Mode silently returning garbage | `is_ai_mode_refusal()`'s flat page-length cap missed a real refusal buried after a long echoed prompt | Check only the **tail** of the page text for refusal markers |
| One specific GLP-1 chunk refuses every retry | Content-policy block on drug brand names | **Wrong theory** — genericizing brand names did *not* fix it (verified directly) | Turned out to be §below |
| A trivial, unrelated query ("capital of France?") also got refused | — | The automation profile/session itself had been soft-flagged by Google after sustained high-volume querying — **not** content-specific at all | `python manage_ai_mode_extensions.py --reset-profile` wipes browsing state (keeps installed extensions), re-seeds from real Chrome on next launch. Cleared the block immediately, confirmed on the exact chunk that had been stuck for hours. |
| Second collision: "Opening in existing browser session" | Race condition | Two script instances (an orphaned prior run + a new one) both trying to hold the same persistent profile lock | Kill any `chrome.exe` process whose command line references `ai_mode_chrome_profile*` before relaunching |

The single biggest lesson: **a refusal that survives a profile reset is a
different problem than one that doesn't.** The GLP-1 "brand name" theory
looked plausible and even seemed briefly confirmed, but a controlled test
(same chunk, brand names left as-is, after a profile reset) proved it was
never about the words — it was the session being flagged. Resetting the
profile, not rewording the prompt, was the actual fix.

---

## 9. Known scope limitation

The `semiconductor` market entry (`Semiconductor_Market.xlsx`) applies the
same strict distributor-only criteria as every other market, even though
that sheet is a general semiconductor-industry landscape (equipment
makers, materials suppliers, foundries, IDMs — not primarily
distributors). Result: 17 kept out of 345, because almost everything in
that sheet is correctly DROPped as "pure manufacturer/OEM, not a
distributor" under these criteria. This was an explicit choice (asked and
confirmed), not a bug — the criteria block would need a different
`market_name`/KEEP definition ("semiconductor market participant" instead
of "distributor") to get a meaningfully different keep rate on that
specific file.

---

## 10. Output files per market

For a source file at `path/to/market.xlsx`:

| File | Contents |
|---|---|
| `path/to/market_verify_checkpoint.json` | `{"offset": int, "verdicts": [...]}` — resume state |
| `path/to/market_verified_distributors.xlsx` | KEEP rows only, all original columns, wrapped/top-aligned cells |
| `path/to/market_verify_log.json` | Every verdict (KEEP and DROP), for audit |

A separate export step (`export_final_verified.py`) copies each market's
`*_verified_distributors.xlsx` to a destination folder with column widths
sized per field and wrap-text enabled, so results are readable without
manually expanding cells.
