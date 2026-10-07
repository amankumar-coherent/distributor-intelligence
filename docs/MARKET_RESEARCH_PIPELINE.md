# How the market-research pipeline works

How one market string becomes a few hundred verified distributor rows with
all 20 columns filled — the queries it sends, how each next query differs
from the last, how it keeps going until the target, and why a single
question returns a company name *together with* all its details instead of
just a list of names.

Everything below is what the code does, with the constants that control
it. File references are clickable.

---

## 1. The shape of the whole thing

```
  MARKET string  ("Global Polyurethane Market - Distributors Only")
        │
        ▼
  ┌───────────────────────────────────────────────────────────┐
  │ DISCOVERY LOOP          run_market_research.py            │
  │                                                           │
  │   batch (target 20 companies)                             │
  │     └─ round 1..8  ── one AI Mode query each ── 10 asked  │
  │          │                                                │
  │          ├─ parse JSON  → 20 fields per company           │
  │          ├─ gate: verdict must be "distributor"           │
  │          ├─ gate: not a duplicate (name / website+country)│
  │          ├─ gate: every firmographic field present        │
  │          └─ accept → state file (checkpoint)              │
  │                                                           │
  │   repeat batches until target, or the market dries up     │
  └───────────────────────────────────────────────────────────┘
        │
        ▼
  ┌───────────────────────────────────────────────────────────┐
  │ VERIFICATION PASS      verify_distributors_ai_mode.py     │
  │   chunks of 10 · KEEP/DROP · confidence ≥ 70 · fail-closed│
  └───────────────────────────────────────────────────────────┘
        │
        ▼
  <market>_verified_distributors.xlsx
```

Two separate passes on purpose. Discovery asks *who is in this market*;
verification asks *is this specific company really a distributor here*.
A model answering both at once tends to justify whatever it just said.

---

## 2. The first query

Built by `build_draft_prompt()` in
[discovery.py](../src/vendor_intel/pipeline/market_research/discovery.py).
It is one prompt with five parts, sent to Google AI Mode as a `udm=50`
search URL.

### 2.1 System framing

```
You are a precise market-research analyst. You only list real, verifiable
companies and only state facts you actually know with confidence. You never
invent company names, websites, emails, phone numbers, or people's names —
for anything you are not sure of, use an empty string rather than guess.
```

### 2.2 What a distributor IS (the role rule)

Chosen by `distributor_definition_for(market)`. Physical-goods markets get:

```
Give me 20 DISTRIBUTOR companies of this market ONLY. A qualifying
company's PRIMARY business must be buying in volume from manufacturers and
reselling/distributing onward (wholesale / authorized / value-added /
stocking distributor). That is what it IS, not something it also does.

STRICTLY EXCLUDE, even if related to this market:
- CHANNEL PARTNERS, resellers, VARs, dealers, agents, brokers, sales reps…
- Manufacturers/OEMs (a manufacturer's sales subsidiary is still the
  manufacturer).
- SYSTEM INTEGRATORS …
- Pure service providers, consultancies …
If unsure whether a company is a true wholesale distributor, EXCLUDE it.
```

Software/cloud markets get the **inverted** rule, because there is no
physical stock to warehouse:

```
This is a SOFTWARE/CLOUD market, so there is no physical stock: here a
distributor IS the channel partner … authorized reseller, master agent /
master distributor, value-added reseller, white-label or wholesale
aggregator …
EXCLUDE: the PLATFORM VENDORS themselves …
```

Routed by `_SOFTWARE_MARKET_HINTS` — `cpaas`, `saas`, `software`, `cloud`,
`api`, … So CPaaS returns Twilio's *resellers*, never Twilio.

### 2.3 What it must SELL (the product rule)

`product_scope_for(market)`. The role rule alone is not enough: an
ophthalmic **equipment** distributor satisfies every word of "wholesale
distributor" while selling no dry-eye drug. Measured before this existed:
98 of 175 Dry Eyes companies were device distributors.

```
PRODUCT SCOPE (required): "dry eye" AND ("pharmaceuticals" OR "ophthalmic
emulsion" OR "lubricant eye drops") AND ("distributor" OR "wholesaler").
It must distribute dry-eye DRUGS: artificial tears, lubricant drops,
ophthalmic emulsions/gels, cyclosporine (Restasis/Ikervis/Cequa),
lifitegrast (Xiidra), perfluorohexyloctane (Miebo) … EXCLUDE ophthalmic
EQUIPMENT/DEVICE/INSTRUMENT distributors (diagnostics, slit lamps,
tear-film analysers, lasers, IPL, phaco, OCT …), contact lens/spectacle/
optical retail … Name the dry-eye drug in key_products_technologies; if you
cannot, EXCLUDE the company.
```

The boolean form is deliberate — it states the required *intersection*
more sharply than prose. Every market has its own; add one in
`_PRODUCT_SCOPES`.

### 2.4 Global coverage

```
GLOBAL COVERAGE IS A HARD REQUIREMENT, not a preference. Distributors
exist in every region — Latin America, Africa, the Middle East, South and
Southeast Asia, Eastern Europe, the Nordics, Oceania — and a list drawn
only from the US, Germany, Japan, Taiwan and China is INCOMPLETE.
Prefer a real regional/national distributor in an under-covered country
over yet another large US or Western-European name.
```

### 2.5 The field spec — this is what makes one query return everything

The key design decision: **the query does not ask for company names.** It
asks for a JSON array where each object already carries all 20 columns.
One round-trip, not "name → then look each one up".

```
Reply with ONLY a JSON array, no markdown fences, no commentary, one
object per company:
[{"company": "", "verdict": "", "verdict_reason": "", "website": "",
"founded": "", "headquarters": "", "continent_geography": "",
"operational_presence": "", "ownership": "", "employees": "",
"core_categories": "", "specialty_focus": "",
"key_products_technologies": "", "online_offline_both": "",
"contact_person": "", "role": "", "email": "", "linkedin": "",
"office_no": "", "office_email": ""}]
```

Each field has its own rule. The ones that matter most:

| Field | Rule in the prompt |
|---|---|
| `verdict` | `"distributor"` / `channel_partner` / `manufacturer` / `system_integrator` / `other` / `unknown`. **Include the company ONLY if `distributor`.** Never guess. |
| `continent_geography` | EVERY country with **verified physical presence** (HQ, office, plant, warehouse). Format `REGION(Country, Country)`. Never "Global" or "50+ countries". |
| `operational_presence` | EVERY country it does business in — **broader**, includes all of the above plus more. |
| `employees` | Always a RANGE (`1,000-5,000`), never a bare number. |
| `core_categories` | Real product categories. **"NEVER repeat the market name — it is NOT a category and will be rejected."** |
| 6 contact fields | **Empty string required** unless the real published value is known. Never build an email from the domain, a LinkedIn URL from a name, or a phone number from a country code. *"A wrong contact detail is worse than an empty one."* |

Note the asymmetry: firmographic fields **must** be filled (a blank drops
the company), contact fields **must** be blank unless genuinely known.
That split is what stopped the model fabricating `info@<domain>` for 94%
of rows.

---

## 3. How the next query differs

Each round sends a *different* prompt. Four things change:

### 3.1 The exclusion list grows

```
Do NOT repeat any of these (incl. their subsidiaries, aliases or former names):
- Quimidroga S.A.
- Magpol Kimya A.Ş.
- …
- ...and 340 more already-covered companies: do not repeat ANY well-known
  distributor you would expect to have been listed already; prefer
  smaller/regional distributors not yet mentioned.
```

Only the **25 most recent** names are listed (`_MAX_EXCLUDED_IN_PROMPT`);
the rest collapse into that trailing instruction. Listing 400 names would
blow the URL limit (§3.4).

### 3.2 The geographic steer rotates

`format_geo_steer()` counts collected HQ countries and names gaps:

```
ALREADY WELL COVERED — do not return more from here: North America, Western Europe.
PRIORITISE distributors headquartered in: Kuwait, Tunisia, Taiwan, Czechia, Hungary, Romania.
```

Only **6 countries per call**, and the slice **rotates**:

```python
offset = (len(rows) // 5) % max(1, len(absent))
rotated = (absent + absent)[offset : offset + 6]
```

So every ~5 companies the window advances — successive calls target
different countries, covering all ~22 target countries across batches at a
fraction of the prompt cost. This is what produced **54-80 countries per
market** instead of a US/Europe list.

Silent below 25 rows, so early batches are not steered on noise.

### 3.3 The ask shrinks as the batch fills

```
round 1/8: asking for 10 more companies (0/20 filled so far)...
round 2/8: asking for 10 more companies (8/20 filled so far)...
round 3/8: asking for 3 more companies (17/20 filled so far)...
```

`request_size = min(remaining, 10)` — never more than
`_MAX_DRAFT_REQUEST_SIZE = 10`, because a 20-company answer with 20 fields
each measured ~267s and often timed out.

### 3.4 It is re-sized to fit Google's URL limit

The whole prompt rides in `?q=`, and Google 400s past ~8 KB.
`MAX_URL_CHARS = 7800`, minus `_URL_SAFETY_MARGIN = 250`.

Truncation is not graceful: `build_ai_mode_url` cuts from the **end**,
which removes the JSON reply template — so the model has no output format
and returns nothing. The builder therefore *measures* the assembled
request (system + prompt, since `chat()` concatenates them) and shrinks in
priority order:

1. named exclusions (degrade smoothly — the "…and N more" line remains)
2. the geo steer (costs coverage breadth, not correctness)
3. a compact distributor definition

The product scope and the reply template are **never** dropped.

---

## 4. Reaching the target

### 4.0 Round vs batch — the two nested loops

Three levels, and the words are easy to mix up:

```
RUN      = keep making batches until target (400) or the market dries up
  BATCH  = collect up to 20 accepted companies          (batch_size)
    ROUND  = ONE AI Mode query, asking for up to 10     (_MAX_DRAFT_REQUEST_SIZE)
```

**Why a round asks 10 when a batch wants 20.** Asking for 20 companies
with 20 fields each measured ~267 s per call and frequently timed out, and
a longer prompt is likelier to be truncated past the URL limit. Two
smaller queries are faster and far more reliable than one large one.

**Why a batch needs up to 8 rounds for 20 companies.** Most drafts are
rejected. A round asking for 10 typically yields 2-6 acceptances — the
rest are duplicates, manufacturers, or rows missing a required field. So
the batch keeps asking, shrinking the request each time:

```
round 1/8: asking for 10 more companies (0/20 filled so far)...
  → got 10 drafts, accepted 8          (2 were manufacturers)
round 2/8: asking for 10 more companies (8/20 filled so far)...
  → got 10 drafts, accepted 9          (1 duplicate)
round 3/8: asking for 3 more companies (17/20 filled so far)...
  → got 3 drafts, accepted 3           → batch full, stop early
batch 12 done: +20 new (total now 240/400)
```

`request_size = min(batch_size - accepted, 10)` — round 3 asks for only 3
because that is all the batch still needs. A batch ends when it is full,
when the 8 rounds are spent, or when two consecutive rounds add nothing.

**Why batches at all, rather than one long round loop?** The batch is the
checkpoint boundary: state is saved and the Excel refreshed after each
one, so a crash costs at most the batch in flight. It is also the unit
the stop rules measure (§4.2) — "this batch yielded 2" is a meaningful
signal, whereas a single weak round is just noise.

### 4.1 Rounds inside a batch — [batch_runner.py](../src/vendor_intel/pipeline/market_research/batch_runner.py)

```
for round in 1..8:                      # _MAX_DRAFT_ROUNDS
    ask for min(remaining, 10)
    for each draft returned:
        verdict != "distributor"      → reject
        no website                     → reject
        duplicate name                 → skip
        duplicate website+country      → skip
        any firmographic field blank   → reject
        otherwise                      → ACCEPT
```

A round returning nothing is retried **once** before the batch gives up —
an empty answer is as likely to be a transient failure as an exhausted
market, and stopping on the first one cost a live batch 14 of its 20
companies.

### 4.2 Batches until target

```
run_until_target(market, target=400, batch_size=20)
```

After each batch the state file is saved and the Excel refreshed, so an
interruption never loses more than the batch in flight.

Three stop conditions:

| Condition | Constant | Why |
|---|---|---|
| 3 consecutive **empty** batches | `_MAX_CONSECUTIVE_EMPTY_BATCHES = 3` | genuinely exhausted |
| 3 consecutive batches yielding **≤2** | `_LOW_YIELD_THRESHOLD = 2`, `_MAX_LOW_YIELD_BATCHES = 3` | dried up without ever being empty |
| target reached | — | done |

The low-yield rule matters: a market can trickle +1, +2, +1 forever while
the model recycles names already held. It needs **three in a row** —
Polyurethane went +4, +1, +2, then **+9**, and a first-dip rule would have
stopped it at 303 instead of 404.

### 4.3 A real yield curve (CPaaS, 55 batches)

```
20 17 18 16  2 16 16  5  4 13 12  6 15 10 15 14 10  2  1 14 15 13 13  5
 9  7  5  6  4  1  5  2  2  4  8  3 10  1  1  3  1  6  1  2  3  2  1  4
 1  3  2  2  1  2 ← three ≤2 in a row → "exhausted at 377"
```

Strong early, transient dips that recover, gradual decline, then a genuine
stop. Not a smooth curve — which is exactly why the stop rule is
conservative.

---

## 5. Why one query returns name **and** details

This is the question people usually ask, so directly:

**Not** `"list 20 companies"` → then 20 follow-up lookups. That would be
21 queries per batch, ~13 s each, and each follow-up would re-ground the
model on a name it may have invented in the first place.

**Instead** one query whose reply *schema* is the whole row. The model
must produce `company` and `website` and `founded` and `headquarters` and
17 more fields in the same object, or the row is rejected. Three
consequences:

1. **20 fields for the price of one query.** A 20-company batch costs ~8
   queries, not ~160.
2. **Self-consistency.** A model that must simultaneously state the
   website, HQ, founding year and product categories is far less likely to
   emit a company that does not exist — there is nothing to attach the
   details to.
3. **The `verdict` field is the gate.** The model classifies its own
   answer (`distributor` / `channel_partner` / `manufacturer` / …) and
   [batch_runner.py](../src/vendor_intel/pipeline/market_research/batch_runner.py)
   accepts **only** `"distributor"`. A manufacturer that slips into the
   candidate set removes itself.

### The completeness rule

`schema.finalize_row()` rejects a row where any
`schema.REQUIRED_FIELDS` value is blank or a placeholder (`N/A`,
`unknown`, `TBD`). So "all columns filled" is not a hope — it is enforced,
and an incomplete company never reaches the sheet.

`REQUIRED_FIELDS` is the 14 firmographic columns: Company, Website,
Founded, Headquarters, Continent / Geography, Operational Presence,
Ownership, Employees, Core Categories, Specialty Focus, Key Products /
Technologies, Online / Offline / Both, Office Website, Country Code /
Region Code. Measured across the six finished markets: **1,891/1,891 rows
complete**.

(Office Website is still in `REQUIRED_FIELDS` but was dropped from the
exported workbooks on request, so the shipped files have 20 columns.)

The six contact fields are deliberately exempt, for the reason in §2.5.

---

## 6. Duplicate control

Two keys, because one is not enough:

| Key | Catches |
|---|---|
| normalised **name** (`company_dedupe_key`) | punctuation, legal suffixes — `Acme Ltd` = `Acme Limited` |
| **website + country** (`website_identity`) | genuinely different names for one company |

The second was added after finding this in live data:

```
powernsun.com   PowernSun
                Powernsun (Power n Sun Solar Systems)
                Power & Sun
                Power n Sun
                Power & Sun FZCO
                Powernsun Group
                Power N Sun Solar Equipment's Trading LLC     ← 7 rows, 1 company
```

Name matching caught 5 of 29 such cases; website matching caught the rest.
Across five markets it removed **308 redundant rows (13-20%)**.

**Country is part of the key** on purpose — a multinational runs one domain
across genuinely distinct national distributors:

```
krannich-solar.com   Krannich Solar GmbH & Co. KG    Germany       ← kept
                     Krannich Solar Turkey           Turkey        ← kept
                     Krannich Solar d.o.o.           Croatia       ← kept
                     Krannich Solar Energy (Pty)     South Africa  ← kept
                     Krannich Solar Private Limited  India         ← kept
                     Krannich Solar Group            Germany       ← merged
```

Country aliases are normalised (`UAE` = `United Arab Emirates`, `USA` =
`United States`), or a duplicate escapes under the other spelling.

---

## 7. Verification — [verify_distributors_ai_mode.py](../verify_distributors_ai_mode.py)

A second, independent pass over the finished sheet, in chunks of 10:

```
Verify these 10 companies … for each return:
{"company":"", "in_market":true|false, "is_distributor":true|false,
 "fits_criteria":true|false, "confidence":0-100, "reason":""}
```

KEEP requires **all three flags true AND confidence ≥ 70**. Anything else
drops. Typical outcome:

```
DONE: 387 kept, 17 dropped
   Novergy Clean Tech    → "primarily an OEM / manufacturer"
   Sunsynk Distribution  → "primarily a manufacturer, OEM"
   Getz Pharma           → "primarily research-driven"
```

**Fail-closed**: an unparseable answer drops the whole chunk rather than
guessing. But a chunk that failed for *backend* reasons (browser died,
quota hit) is **not** checkpointed — it is marked `backend_failure`, the
offset rewinds, and a re-run re-asks. Otherwise a browser crash would
silently record 10 real companies as rejected distributors.

---

## 8. Running it

```powershell
# one market
.venv\Scripts\python.exe run_market_research.py `
  --market "Global Polyurethane Market - Distributors Only" `
  --target 400 --batch-size 20 --export-xlsx --verify

# a queue, one market at a time (they share one browser profile and one IP)
.venv\Scripts\python.exe run_all_markets.py --target 400
```

Everything resumes: discovery checkpoints per batch, verification per
chunk of 10. Re-run the same command after any interruption.

### Adding a market

1. Add the string to `MARKETS` in [run_all_markets.py](../run_all_markets.py)
   — it must end `- Distributors Only` (the state filename derives from it).
2. Add a product scope to `_PRODUCT_SCOPES` in
   [discovery.py](../src/vendor_intel/pipeline/market_research/discovery.py).
   Without one, nothing constrains *what* the company sells.
3. If it is software/cloud, confirm a `_SOFTWARE_MARKET_HINTS` term matches
   so it gets the inverted definition.

---

## 9. Failure modes worth knowing

| Symptom | Cause | Handling |
|---|---|---|
| `got 0 draft(s) back`, every round | prompt truncated past 7,800 chars → reply template lost | builder measures and shrinks; guard test pins ≥150 chars headroom |
| `Core Categories = "Global X Market - Distributors Only"` | prompt interpolated the market into the field it asked about | prompt reworded; `_clean_category` blanks an echo |
| `phone 85 -> 85`, no errors | Google quota hit; echoed template read as a valid empty answer | `is_ai_mode_rate_limited()` stops the run. **Cure: reset the profile — the quota is tied to session cookies, not the IP** |
| CAPTCHA | anonymous AI Mode | captcha-raptor extension auto-solves in ~3-11 s; the query is re-sent |
| `Something went wrong` ×3 | Google-side generation failure | `reset_session()` rebuilds the profile (staged extensions survive) |
| 0 kept / 0 dropped on a full file | banner pushed the header below row 1 | `find_header_row()` scans for `Company`; prints `NOTHING TO VERIFY` |

---

## 10. Results from six markets

| Market | Verified | Countries |
|---|---|---|
| Silicon Carbide | 298 | 63 |
| Wearable Glucometer | 298 | 80 |
| Dry Eyes Treatment | 327 | 76 |
| Solar Rooftop | 311 | 73 |
| Polyurethane | 348 | 54 |
| CPaaS | 309 | 70 |
| **Total** | **1,891** | |

All 20 columns present, **0** vague geography values, **0** fabricated
LinkedIn URLs, **0** duplicate rows, **1,891/1,891** complete on every
required field.
