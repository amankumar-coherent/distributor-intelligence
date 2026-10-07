# Extracting Company Data from Google AI Mode

How to write prompts that make AI Mode return **structured, real company
data**, and how to parse it reliably.

Companion to `GOOGLE_AI_MODE_INTEGRATION.md` (that one covers the browser
plumbing — Patchright, CAPTCHAs, pacing). This one is only about
**queries in, companies out**.

Reference implementation:
`src/vendor_intel/pipeline/market_research/discovery.py`

---

## 1. The shape of the whole thing

```
market name + exclusion list
        ↓
one GET request:  google.com/search?q=<entire prompt>&udm=50
        ↓
rendered page text  (prompt echoed back, then the answer)
        ↓
strip echo → strip fences → repair newlines → json.loads
        ↓
per-company dict → validate → dedupe → append to state
        ↓
repeat with the new names added to the exclusion list
```

Three properties of AI Mode drive every design decision below:

1. **Stateless.** There is no conversation. "Give me 20 more" means nothing
   — every request must carry its own exclusion list.
2. **The whole prompt travels in the URL.** Google 400s past ~8 KB, so the
   prompt has a hard size budget that anything growing per-iteration must
   respect.
3. **It echoes the prompt back.** The page text contains your question
   *before* the answer — which breaks naive JSON parsing (§5).

---

## 2. Anatomy of a working extraction prompt

Seven parts, in this order. The order matters: constraints stated *after*
the output template get less reliably applied.

```
1. Market / subject          "Market: "Global LNG Market""
2. What you want             "Give me 20 DISTRIBUTOR companies"
3. Inclusion definition      what the category IS, positively
4. Exclusion list (types)    what must be rejected, with reasons
5. Exclusion list (names)    companies already collected
6. Field-by-field rules      one line per output field
7. Output format             explicit JSON template
```

### Part 2-3: define the category by its PRIMARY business

Vague category names return adjacent companies. Anchor on what the company
*is*, not what it *also does*:

```
A qualifying company's PRIMARY business must be buying in volume from
manufacturers and reselling/distributing onward (wholesale / authorized /
value-added / stocking distributor). That is what it IS, not something it
also does.
```

That last sentence does real work: without it you get manufacturers who
also wholesale, and integrators who also resell.

### Part 4: exclusions need reasons, not just labels

Listing forbidden categories is not enough — say *why* each fails, or the
model applies its own looser definition:

```
STRICTLY EXCLUDE, even if related to this market:
- CHANNEL PARTNERS, resellers, VARs, dealers, agents, brokers, sales reps,
  referral partners, "solution providers" — selling/representing a vendor's
  products WITHOUT stocking/wholesaling is NOT a distributor here.
- Manufacturers/OEMs (a manufacturer's sales subsidiary is still the
  manufacturer).
- SYSTEM INTEGRATORS — primary business is designing/building/installing/
  integrating solutions, even if they resell hardware as part of a project.
- Pure service providers, consultancies, managed-services firms.
If unsure whether a company is a true wholesale distributor, EXCLUDE it.
```

The parentheticals are the useful part. *"A manufacturer's sales subsidiary
is still the manufacturer"* closes a loophole the model otherwise walks
straight through.

### Part 6: a self-reported verdict field is the best filter

Rather than trusting the model's category judgement implicitly, make it
**declare** the classification as data you can gate on:

```
- verdict: "distributor" only for a genuine wholesale/authorized
  distributor; "channel_partner" for a reseller/VAR/dealer/agent/solution
  provider; "manufacturer"; "system_integrator"; "other"; or "unknown" if
  you truly do not know. NEVER guess this field. Include the company in
  your reply ONLY if verdict is "distributor".
- verdict_reason: one sentence on why it is a wholesale distributor rather
  than a channel partner or reseller.
```

Then reject in code, so a mislabel cannot slip through:

```python
if item.get("verdict", "").strip().lower() != "distributor":
    return None      # channel_partner, manufacturer, unknown -> all dropped
```

Two things make this work: `"NEVER guess this field"` (the only field
carrying that instruction), and `verdict_reason`, which forces the model to
justify itself and measurably improves the verdict's accuracy.

### Part 7: always end with an explicit template

AI Mode has no JSON mode but follows an explicit instruction reliably:

```
Reply with ONLY a JSON array, no markdown fences, no commentary, one object
per company:
[{"company": "", "verdict": "", "website": "", "founded": "", ...}]
```

The empty-string template is what keeps keys consistent. **It is also what
breaks parsing if you are careless** — see §5.

---

## 3. Field rules that actually change the output

Each of these exists because the loose version produced unusable data.

| Field | Rule that works | Failure it fixes |
|---|---|---|
| `website` | "real official domain (not a directory/social page)" | LinkedIn/Bloomberg pages returned as the website |
| `founded` | "4-digit year" | "Founded in 1998", "1998 (as Acme Ltd)" |
| `employees` | "ALWAYS a RANGE like `1,000-5,000` — NEVER a bare number. Convert ~23,000 to `20,000-25,000`" | false precision from an estimate |
| `headquarters` | "city, country" | country only, or a full street address |
| geography | see below | "Global", "50+ countries" |

### Geography: name every country, ban the summary words

The single most common failure is a plausible-looking summary instead of
data. Ban the specific words, give a format, and give an example:

```
- continent_geography: EVERY country where it has VERIFIED PHYSICAL
  PRESENCE (HQ, offices, plants, R&D, service centres, warehouses). Format
  as REGION(Country, Country); e.g. "APAC(India, Singapore, Japan);
  EU(Germany, France)". Name every country individually. NEVER write
  "Global", "Worldwide", a bare region, or a count like "50+ countries".
  Physical presence only — not places it merely sells to.

- operational_presence: EVERY country where it meaningfully does business
  even with no office (selling, distributing, supplying, servicing).
  BROADER than continent_geography: all of those countries PLUS any extra.
```

Three techniques worth copying:

- **Enumerate what counts** ("HQ, offices, plants, R&D, service centres,
  warehouses") rather than saying "presence".
- **Ban the exact strings** you keep receiving. "NEVER write Global" works;
  "be specific" does not.
- **Define the pair relationally** — "BROADER than X, all of those PLUS
  extra" stops two columns coming back identical.

Then enforce it in code, because the model still slips:

```python
_VAGUE_GEO_RE = re.compile(
    r"^\s*(global|worldwide|international|multinational|na|eu|apac|emea|"
    r"north america|europe|asia|africa|middle east|latin america|"
    r"\d+\+?\s*countries?)\s*$", re.I)
```

---

## 4. The rule that matters most: never demand a full row

> **Do not write "no field may be blank".** The model will invent values to
> comply.

Measured in this project, with that instruction in place across 493
companies:

| Field | Result |
|---|---|
| Office Email | **94%** were `info@` or `sales@` at the company's own domain |
| LinkedIn | URLs built from the contact's name (`/in/<name-slug>`) |
| Office No. | numbers generated from the HQ country's calling code |

None of it was researched. All of it looked researched.

The fix is to split the fields into two classes with opposite instructions:

```
Fill every FIRMOGRAPHIC field below for every company (no blank/"N/A"/
"unknown"). The six CONTACT fields at the end are the exception and MUST
be left "" when the real value is not known:

The SIX CONTACT FIELDS below are different from every field above: for
these, an EMPTY STRING IS REQUIRED unless you actually know the real
published value. NEVER construct, pattern-generate, infer or guess them.
Do NOT build an email from the domain. Do NOT build a LinkedIn URL from a
person's name. Do NOT invent a phone number from a country calling code.
A wrong contact detail is worse than an empty one.

- contact_person: real named sales/procurement leadership contact ONLY if
  you actually know who holds that role. Otherwise "".
- email: ONLY a real published email you actually know. Otherwise "".
  Never derive it from the website domain.
- linkedin: ONLY the contact person's real profile URL you actually know.
  Otherwise "". Never construct one from their name.
- office_no: ONLY a real published telephone number. Otherwise "".
```

Effect of that change on the same market, before → after:

| Field | Before (invented) | After (real only) |
|---|---|---|
| Contact Person | 100% | 9.6% |
| Office No. | 100% | 48.6% |
| Office Email | 100% | 25.0% |
| LinkedIn | 100% | **0%** |

Lower coverage, but the remaining values are real — specific mailboxes like
`post@bng.no` and `S.Operations@fuelng.com.sg`, real switchboard numbers
with correct country codes. Only 8 of 104 emails were generic patterns,
down from 94%.

**LinkedIn going to exactly 0% is the honest answer**, not a bug: a specific
person's profile URL is the hardest thing to know, and every candidate was
constructed. A column of plausible dead URLs is worse than an empty one.

### Fabrication detector

A field that is overwhelmingly **one pattern** was generated, not
researched. Checking a column's distribution takes seconds:

```python
from collections import Counter
locals_ = Counter(e.split("@")[0] for e in emails if "@" in e)
print(locals_.most_common(5))
# [('info', 367), ('sales', 96), ...]   <- 94% of 493: formula, not research
```

---

## 5. Parsing the answer (four defences, all required)

AI Mode returns **rendered page text**, not JSON. Four things must be
handled, in this order.

### 5.1 Strip the echoed prompt — the subtle one

AI Mode renders your question *before* the answer. Since the prompt **ends
with an empty JSON template**, a naive `text.find("[")` locks onto the echo
and returns zero companies.

```python
def _strip_echoed_prompt(raw: str) -> str:
    text = raw or ""
    # Whitespace-tolerant: the echoed template is compact (`[{"company"`)
    # but a real answer is often pretty-printed (`[\n  {"company"`).
    starts = [m.start() for m in re.finditer(r'\[\s*\{\s*"company"', text)]
    if starts:
        return text[starts[-1]:]          # FROM the last opening, not past it
    obj = text.find('{"company"')
    return text[obj:] if obj > 0 else text
```

Two bugs lived here:

- **A literal `'[{"company"'` marker** matched only the compact echo and
  missed the pretty-printed real answer → 0 companies from a good reply.
- **Slicing `starts[-1] + 1`** instead of `starts[-1]` dropped the first
  record of **every** batch — 3 companies silently became 2. Found only by
  a failing test.

### 5.2 Strip Google's wrappers

```python
FENCE = chr(96) * 3          # a markdown code fence, written this way so
                             # the snippet can live inside one
if raw.startswith(FENCE):
    raw = raw.strip(chr(96))
    if raw.lower().startswith("json"):
        raw = raw[4:]                     # stray language tag
raw = raw.replace("Use code with caution.", "")   # Google's own UI footer
```

### 5.3 Repair literal newlines inside strings

Reading an answer out of rendered text can leave **real newlines inside
string values** (observed: `"company": "\nDatumstruct\n"`). That is illegal
JSON, so `json.loads` rejects the entire array and every company is lost.

```python
try:
    data = json.loads(candidate)
except (json.JSONDecodeError, ValueError):
    data = json.loads(re.sub(r"\s*\n\s*", " ", candidate))   # retry once
```

### 5.4 Rebuild the array if the opening bracket was consumed

After echo-stripping, the text can start mid-array:

```python
start, end = raw.find("["), raw.rfind("]")
if start < 0 or end <= start:
    first, last = raw.find("{"), raw.rfind("}")
    if first < 0 or last <= first:
        return []
    candidate = "[" + raw[first:last + 1] + "]"
else:
    candidate = raw[start:end + 1]
```

### 5.5 Normalise every value

```python
def _clean_value(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()

name = _clean_value(item.get("company"))
name = re.sub(r"^\d+[.)]\s*", "", name)      # drop "1. " / "1) " numbering
founded = re.sub(r"[^\d]", "", founded)[:4]  # "Founded in 1998" -> "1998"
```

---

## 6. Batching and continuation (a stateless "give me 20 more")

There is no conversation, so continuation is entirely prompt-side: embed
the names already collected in **every** request.

```
Do NOT repeat any of these (incl. their subsidiaries, aliases or former
names):
- TD SYNNEX
- Ingram Micro
...
```

### The URL budget problem, and why trimming is safe

The exclusion list is the only part that grows without bound — one line per
company found, heading for 500. Live failure: at **148 companies** the
prompt hit 9,086 chars, the URL 10,701, and **every request 400'd**
mid-run.

So cap it:

```python
_MAX_EXCLUDED_IN_PROMPT = 40      # newest-first
_MAX_PROMPT_CHARS = 1500          # budget for the exclusion block alone

def _format_excluded(excluded: list[str]) -> str:
    if not excluded:
        return "(none yet)"
    recent = list(reversed(excluded))[:_MAX_EXCLUDED_IN_PROMPT]
    lines, used = [], 0
    for name in recent:
        line = f"- {name}"
        if used + len(line) + 1 > _MAX_PROMPT_CHARS:
            break
        lines.append(line)
        used += len(line) + 1

    omitted = len(excluded) - len(lines)
    if omitted:
        lines.append(
            f"- ...and {omitted} more already-covered companies: do not "
            f"repeat ANY well-known distributor you would expect to have "
            f"been listed already; prefer smaller/regional distributors "
            f"not yet mentioned.")
    return "\n".join(lines)
```

**Trimming is safe for correctness** because the prompt list is only a
*hint*. The authoritative duplicate guard is a state check against the full
history, so a trimmed name that comes back is silently skipped. Two details
that make the cap work well:

- **Newest first** — a repeat is likeliest among companies just returned.
- **Disclose the omission with a strategy** — telling the model *"and 453
  more, prefer smaller/regional ones not yet mentioned"* keeps it
  productive far longer than silence.

### Batch sizes, measured

| Size | Behaviour |
|---|---|
| 20 fields × 20 companies | answers truncate |
| **10 per request** | reliable for the full 19-field schema |
| 15 per request | fine for a 6-field contact-only lookup |

Ask for fewer companies per request than you want per batch, and loop.

### Diminishing returns are real

At 493 companies in one market, only ~40 exclusion names fit the prompt.
Nearly everything returned was already collected, filtered locally, and the
batch came back empty. That is a **genuine ceiling**, not a bug — but see
§8, because it is easily confused with a rate-limit block.

---

## 7. Validate in code; never trust the reply

The prompt is a request, not a guarantee. Every rule needs a code-side gate.

```python
def build_row(item: dict) -> dict | None:
    # 1. hard category gate on the self-reported verdict
    if item.get("verdict", "").strip().lower() != "distributor":
        return None

    # 2. a company with no website is unusable
    if not item.get("website"):
        return None

    # 3. normalise what is fixable
    employees = normalize_employees(item["employees"])   # 23,000 -> 20,000-25,000

    # 4. REJECT rather than repair contact fields — a repaired value is
    #    indistinguishable from a real one downstream
    return {
        "Company": item["company"],
        "Employees": employees,
        "LinkedIn": real_linkedin_or_blank(item.get("linkedin")),  # /in/ only
        "Email": real_or_blank(item.get("email")),
        # ... remaining columns
    }
```

> **Repair logic can manufacture data.** Two helpers here silently
> fabricated: one rebuilt a LinkedIn URL from a person's name, another
> copied `Email` into `Office Email`. Both produced values that looked
> independently researched. Prefer rejecting a malformed value to "fixing"
> it.

Split required from optional so blanks do not delete good rows:

```python
CONTACT_FIELDS  = ("Contact Person", "Role", "Email", "LinkedIn",
                   "Office No.", "Office Email")          # may be empty
REQUIRED_FIELDS = tuple(c for c in COLUMNS
                        if c != "#" and c not in CONTACT_FIELDS)

def finalize_row(candidate, *, index):
    for col in REQUIRED_FIELDS:
        if is_placeholder(candidate.get(col)):
            return None                    # drop the company
    for col in CONTACT_FIELDS:
        row[col] = "" if is_placeholder(candidate.get(col)) else clean(candidate[col])
```

Without that split, switching off fabrication would have deleted all 493
companies.

### Deduplication

Normalise before comparing, or `Acme Inc` and `Acme` are two companies:

```python
def company_dedupe_key(name: str) -> str:
    raw = normalize_name(name).lower()       # strips Inc/Ltd/GmbH/etc.
    compact = re.sub(r"[^a-z0-9]", "", raw)
    return ALIASES.get(compact, raw)         # Dr Reddys -> drreddyslaboratories
```

Dedupe at three levels: within one reply, against the accepted set of the
current batch, and against all persisted history.

---

## 8. Empty batch: three causes, opposite responses

Getting this wrong caused a chain of bugs where each fix created the next.

| Cause | Evidence | Correct response |
|---|---|---|
| Rate limited | `/sorry/` URL, "unusual traffic" | cool off 5+ min |
| Model refused | "no response available for this search" | retry now, shorter prompt |
| Genuinely exhausted | neither | stop |

A run once reported **"318/500, genuinely exhausted"** while actually
CAPTCHA'd — and the same detector later reported *"no sign of blocking"* on
a batch whose own log said *"CAPTCHA cleared"*, because blocking was
detected via a counter that a **successful solve resets to zero**.

Two rules:

1. Track a `last_captcha_at` timestamp that a solve does **not** clear.
2. Log the evidence with the conclusion — `"no sign of blocking (3 CAPTCHAs
   this run, last 41 s ago)"` — so a wrong inference is visible rather
   than merely asserted.

---

## 9. Second-pass enrichment

Fields the bulk prompt leaves empty can be re-asked in a **focused**
follow-up. One short question about 15 companies grounds far better than a
20-company request for 19 fields each.

Recovered on a 416-company market this way: **+183 phone numbers, +88
office emails, +26 named contacts** — none previously known.

Three rules for a second pass:

- **Only ask about gaps.** Re-querying filled rows wastes queries (each one
  risks a CAPTCHA) and risks overwriting verified data.
- **Treat an existing value as authoritative.** Never overwrite or blank it
  on the strength of a fresh lookup.
- **Scope which columns may be written**, so a run targeting five fields
  cannot blank the sixth.

```python
if only_empty and existing_value:
    continue                        # authoritative — do not touch
if column not in requested_columns:
    continue                        # out of scope — never written
```

> **This is not independent verification.** Re-asking the same model can
> repeat a plausible invention. It works because the prompt now *rewards*
> "unknown" instead of forbidding it. For real confirmation, fetch the
> company's own website and read the contact page — a different source.

---

## 10. Checklist

```
PROMPT
[ ] Define the category by PRIMARY business ("what it IS, not what it does")
[ ] Exclude adjacent types WITH reasons, not just labels
[ ] Add a self-reported `verdict` field marked "NEVER guess this field"
[ ] Add `verdict_reason` — justification improves the verdict
[ ] One explicit rule per field (format, units, banned strings)
[ ] Ban vague geography by exact string; give a format AND an example
[ ] NEVER say "no field may be blank" — say "" is preferred over a guess
[ ] End with an explicit empty-value JSON template
[ ] Embed the exclusion list in EVERY request (AI Mode is stateless)
[ ] Cap the exclusion list; newest-first; disclose the omitted count

PARSING
[ ] Strip the echoed prompt: LAST `\[\s*\{\s*"key"`, slice FROM it
[ ] Strip ``` fences, a stray "json" tag, "Use code with caution."
[ ] Retry json.loads after collapsing newlines inside strings
[ ] Rebuild the array from { } if the opening [ was consumed
[ ] Collapse whitespace in every value; strip "1. " numbering

VALIDATION
[ ] Gate on verdict in code — never trust the category implicitly
[ ] Reject malformed contact values; never repair them
[ ] Separate REQUIRED from optional fields so blanks don't drop rows
[ ] Dedupe on a normalised key, at all three levels
[ ] Check column distributions for fabrication (one pattern = generated)

LOOP
[ ] ≤10 companies per request for a wide schema
[ ] Distinguish blocked / refused / exhausted before stopping
[ ] Persist after every batch — a run will be interrupted
```
