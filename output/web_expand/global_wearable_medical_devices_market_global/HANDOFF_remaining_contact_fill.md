# Handoff: fill remaining "Not publicly disclosed" cells

## Target file
`D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global\global_wearable_medical_devices_market_global_FINAL.xlsx`
Sheet name: `Wearable Medical Companies`. Row 1 = headers, rows 2-227 = data (226 companies, 21 columns). No merged title row.

## What this file is
A real, previously-built company database of distributors/channel-partners in the global wearable medical device market. Every column already has a value for every row — but many cells hold the literal placeholder string `"Not publicly disclosed"` where real data wasn't found yet. Detect these with a case-insensitive check for the substring `not publicly disclosed`.

## What's already been done (across sessions, most recent first)

**Session 2026-08-24:**
- Ran `scripts/_apollo_fill_wearable_contacts.py --resume` (Apollo + DeepSeek, no WebSearch quota used) — filled 19 new Contact Person/Role/LinkedIn rows out of 41 todo (22 genuinely no_match: giant retailers/small obscure companies with no public sales contact findable via Apollo/DeepSeek). Backup: `..._FINAL_BEFORE_APOLLO_CONTACTS_20260823_231901.xlsx`.
- Ran `scripts/_fill_wearable_office_contacts.py --resume --workers 12` (scrape + DeepSeek extraction, no WebSearch quota used) — filled 8 more Office No. + 5 more Office Email out of 108 remaining gap rows (most sites just don't publish a scrapable phone/email, or already had NPD correctly set). Backup: `..._FINAL_BEFORE_OFFICE_20260823_232204.xlsx`.
- Attempted to fan out 7 parallel general-purpose WebSearch agents (2 factual, 5 LinkedIn-only, 2 contact-needed-fallback batches) to mop up remaining gaps. **6 of 7 agents failed immediately with "org's monthly spend limit hit"** (a billing/credit limit, NOT the per-session 200-WebSearch-call cap seen in the prior session) — this is a different, harder blocker than the search-count cap. Only 1 agent (factual batch 1, 14 companies) completed before the limit hit, and it found exactly one verifiable fact: **Panvel (row 90) Key Brands Represented = "Omron, Medicate, Multilaser, Mediclini, Open Medical"** — applied to the sheet, backup `..._FINAL.backup_before_PANVEL_BRANDS_20260824_013030.xlsx`.
- **Action needed next session**: check `/usage-credits` / org spend limit status before spawning WebSearch agents again — if still capped, either wait for the monthly reset or get an admin to raise the limit. The Apollo/office scripts (no WebSearch dependency) can still be re-run any time regardless of this limit — they're already close to exhausting what they can find, though (see counts below), so most remaining value is genuinely WebSearch-shaped work (personal LinkedIn lookups, personal emails, and Key Brands/Founded/Employees facts for smaller companies).

**Prior session (2026-08-23):**
- Filled all blank Email/Region Code cells that were truly empty (not "Not publicly disclosed").
- **Found and fixed fabricated data**: 8 rows had FAKE personal LinkedIn URLs (templated fake-looking suffix pattern) that didn't match real LinkedIn IDs; reset to "Not publicly disclosed". **Lesson: always sanity-check LinkedIn URLs an agent returns — don't trust a `/in/` URL unless the agent actually surfaced it from a real search result, and watch for suspiciously templated suffixes.**
- Ran ~9 parallel WebSearch-based sub-agents; hit the session-wide 200-WebSearch-call cap partway through.
- Applied all genuinely-found values, leaving genuine "not found" results as-is and never overwriting a real value with a guess.

## Critical rules for continuing this work
1. **Never invent or guess** a person's name, title, email address, or LinkedIn URL. If a search doesn't turn up a verifiable value, leave/set the cell to "Not publicly disclosed" — do not produce something "plausible."
2. **LinkedIn URLs**: only use a URL an agent actually found in real search results. Reject anything that looks templated/generated. A real personal `/in/...` profile is preferred; a company `/company/...` page is an acceptable fallback ONLY if no named individual is findable.
3. **"Email" vs "Office Email"**: "Email" means a specific person's direct email (rare, usually genuinely N/A — don't burn search budget chasing these). "Office Email" is a general company contact address (info@, sales@, customerservice@) — much easier to find, and the office-contacts script already covers this well via scraping.
4. **"Contact Person"/"Role"**: prefer someone in sales, business development, channel partnerships, or commercial leadership. A generic "Customer Service" listing is not a valid Contact Person. Note some Apollo-sourced names come pre-obfuscated with asterisks (e.g. "Matt Pi***e") — that's expected, not a bug; it's fine to keep as-is or use it to disambiguate a LinkedIn search.
5. **Never overwrite a cell that already has real data** — only touch cells that currently read "Not publicly disclosed".
6. **Back up before writing.** Copy the file to `..._FINAL.backup_before_<DESCRIPTION>_<YYYYMMDD_HHMMSS>.xlsx` before saving changes, using openpyxl.

## Two different budget walls to watch for
1. **Per-session WebSearch call cap (~200)** — seen 2026-08-23. Keep total planned WebSearch calls across ALL agents well under 200 per session. Cap each agent at ~1 search per company; split work across multiple chat sessions if needed.
2. **Org monthly spend/credit limit** — seen 2026-08-24, hit almost immediately when 7 agents were launched in parallel. This is NOT about search count — it's a billing ceiling on the whole org's Claude usage for the month. If you hit `"You've hit your org's monthly spend limit"` errors, WebSearch-agent fan-out is not viable until the admin raises the limit or the month rolls over. Check this BEFORE dispatching a wave of agents (e.g. try one small agent first as a canary) rather than after burning a whole batch on failures.
3. **Workaround that sidesteps both**: `scripts/_apollo_fill_wearable_contacts.py` and `scripts/_fill_wearable_office_contacts.py` call Apollo/DeepSeek APIs directly via `requests`/`openai` SDK, NOT the Claude WebSearch tool or any Claude Agent — so they're immune to both walls. Already run twice; diminishing returns going forward since most easy wins are exhausted, but re-run with `--resume` is cheap and safe to try again first each session.

## Current column gap counts (as of 2026-08-24, 223 of 226 rows have at least one gap)
```
LinkedIn: 215
Email: 175
Role: 146
Contact Person: 145
Office Email: 100
Office No.: 44
Key Brands Represented: 19
Founded: 6
Online / Offline / Both: 5
Employees: 4
Website: 2
```
Reality check: LinkedIn/Email gaps are large because for most small-to-mid distributors and even many large retailers, a *named individual's* personal email or LinkedIn profile is genuinely not public — Apollo, DeepSeek, and even manual WebSearch have already failed on most of the "big name" rows (Amazon, Walmart, CVS, Cardinal Health, etc.) across two sessions. Don't expect these counts to drop to zero; the realistic ceiling is probably 30-40% fill on LinkedIn/Email even with unlimited search budget, since it's public-record data that mostly doesn't exist for procurement/sales contacts outside of Apollo-covered records already tried.

## Suggested next-session workflow
1. Re-run both scripts with `--resume` first (free, no WebSearch/spend-limit dependency) in case Apollo's underlying data changed.
2. Test the WebSearch/agent budget situation with ONE small canary agent (3-5 companies) before committing to a full parallel fan-out.
3. If the canary succeeds, prioritize in this order (best ROI first):
   - Factual gaps only 19 rows: Key Brands Represented, Founded (6), Employees (4), Online/Offline/Both (5), Website (2 — "C-Care Medical" Singapore, "Medi-Serve" Hong Kong CPAP/BiPAP/oximeter distributor, currently unknown domain).
   - LinkedIn-only rows where Contact Person + Role are already filled (~65+ rows) — single search each: `"<name>" "<company>" linkedin`.
   - Contact Person/Role for the ~20 small/mid companies Apollo+DeepSeek already tried and failed on this session (rows 190, 194, 196, 198, 199, 202, 205, 206, 208, 210, 211, 212, 214, 215, 216, 218, 223, 224, 225, 226, 227) — one more WebSearch attempt each, low probability but cheap.
   - Personal Email fields — lowest priority/ROI, expect near-zero yield; only chase opportunistically if it surfaces alongside another search.
4. Apply findings directly to the FINAL.xlsx with a backup first; never fabricate.

## Suggested prompt to paste into the new chat
> Read `output/web_expand/global_wearable_medical_devices_market_global/HANDOFF_remaining_contact_fill.md` for full context. First re-run the two Apollo/office-scrape scripts with `--resume`. Then test whether WebSearch-agent fan-out is viable (org spend limit may still be capped) with a small canary agent before committing to a full batch. Prioritize factual gaps and LinkedIn-only lookups over chasing personal emails. Apply real findings directly to the FINAL.xlsx (backup first), never fabricate a name/email/LinkedIn URL.
