# Google AI Mode as a Free LLM Backend — Integration Guide

How this project drives **Google AI Mode** (`udm=50`) as a no-account,
no-API-key generative backend, and how to lift it into another project.

Reference implementation: `src/vendor_intel/scraping/google_ai_mode.py`
(single file, no project-internal imports beyond `patchright`).

**Companion doc:** `AI_MODE_COMPANY_EXTRACTION.md` covers the other half —
how to write prompts that return structured company data, and how to parse
the reply. This doc is the browser plumbing; that one is queries in,
records out.

---

## 1. What this is, and when to use it

Google AI Mode is the generative tab on Google Search. Navigating to
`https://www.google.com/search?q=<prompt>&udm=50` returns a real
model-generated answer — no login, no API key, no per-account quota.

| | AI Mode | ChatGPT via OAuth/Codex | Paid API |
|---|---|---|---|
| Cost | free | free tier | per token |
| Account needed | **no** | yes | yes |
| Quota | none (rate-limited by IP) | ~dozens of calls | billing |
| Latency (measured) | **5–13 s** | ~166 s | 1–5 s |
| Failure mode | CAPTCHA | "usage limit reached" | 429 |

**Use it for** bulk enrichment, classification, and research passes where
throughput matters more than latency and an occasional CAPTCHA is
acceptable.

**Do not use it for** user-facing request paths (a CAPTCHA blocks
everything), anything needing an SLA, or high concurrency — it is one
browser, one query at a time.

### Honest limitations

- **Not a metered API.** No token counts, no `temperature`, no system role,
  no streaming. One prompt string in, rendered page text out.
- **It will CAPTCHA you.** Pacing reduces this; nothing eliminates it.
- **It can refuse.** AI Mode sometimes answers "no response available for
  this search" — a normal page, not a block (see §7).
- **Answers are ungrounded by default.** If you need verified facts, verify
  them against a real source, and tell the model that `""` is an acceptable
  answer (see §9).

---

## 2. The one non-negotiable dependency: Patchright

```txt
patchright>=1.61.0     # NOT playwright
```

```bash
pip install patchright
patchright install chromium    # only if you need the bundled build
```

[Patchright](https://github.com/Kaliiiiiiiiii-Vinyzu/patchright-python) is a
stealth-patched drop-in fork of Playwright. **This is the single reason the
whole approach works.**

Measured on the same machine, same network, same minute:

| Driver | Result |
|---|---|
| `patchright.sync_api` | answer in ~5 s, no CAPTCHA |
| `playwright.sync_api` | instant `/sorry/` CAPTCHA, every time |

The import is otherwise identical, which makes the failure easy to
introduce by accident:

```python
from patchright.sync_api import sync_playwright    # correct
# from playwright.sync_api import sync_playwright  # instantly blocked
```

Also prefer **real Chrome** over bundled Chromium (`channel="chrome"`).
Measured: real Chrome answers cleanly; Chromium got CAPTCHA'd on nearly
every query. Chromium is only needed if you must load an extension (§8).

---

## 3. Minimal working implementation

This is the whole mechanism in ~50 lines. Everything else in this guide is
hardening.

```python
import time, urllib.parse
from patchright.sync_api import sync_playwright

MAX_URL_CHARS = 7800   # Google 400s past ~8 KB

def build_url(prompt: str) -> str:
    return ("https://www.google.com/search"
            f"?q={urllib.parse.quote_plus(prompt)}&udm=50&hl=en")

def ask(prompt: str, profile_dir: str = "./ai_mode_profile") -> str:
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            profile_dir,
            channel="chrome",          # real Chrome, not Chromium
            headless=False,            # a CAPTCHA needs a visible window
            viewport={"width": 1280, "height": 900},
            args=["--profile-directory=Default"],
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(build_url(prompt), wait_until="domcontentloaded", timeout=30_000)

        if "/sorry/" in page.url:
            ctx.close()
            raise RuntimeError("CAPTCHA")

        # Wait for generation: completion marker, else text that stops growing
        deadline, last, stable = time.monotonic() + 90, -1, 0
        while time.monotonic() < deadline:
            if "ai mode response is ready" in page.inner_text("body").lower():
                break
            n = len(page.locator("main").first.inner_text(timeout=1000).strip())
            stable = stable + 1 if n > 0 and n == last else 0
            if stable >= 2:
                break
            last = n
            page.wait_for_timeout(1500)

        text = page.locator("main").first.inner_text()
        ctx.close()
        return text
```

Key choices, each of which was a bug once:

- **`launch_persistent_context`, not `launch`.** A warm profile with cookies
  and history looks human; a fresh incognito context gets challenged
  immediately.
- **`headless=False`.** A CAPTCHA can only be solved in a visible window,
  and MV3 extensions do not run reliably headless.
- **`wait_until="domcontentloaded"`.** AI Mode client-side-redirects right
  after load (appending `&sei=`), which can interrupt Playwright's own
  navigation wait. Treat `"interrupted by another navigation"` as benign.

---

## 4. Waiting for the answer (the expensive mistake)

There is **no single reliable completion signal.** Use three, in order:

1. **`"AI Mode response is ready"`** in the body text — strongest, present
   on normal completions.
2. **A `"Thinking"` indicator appearing and then clearing** — often never
   appears at all, so it cannot be relied on alone.
3. **Answer text unchanged across two consecutive polls** — catches
   responses that render with neither marker.

> **Measured regression worth avoiding:** an earlier version required *both*
> no-"Thinking" *and* >400 characters of text. Every answer shorter than
> 400 chars spun until the 90 s timeout — **~94 s per query instead of
> ~5 s.** Never gate completion on answer length.

---

## 5. Extracting the answer, including URLs

`innerText` silently drops every hyperlink target. Link labels are often
just "Official Website", with the domain only in the `href` — so a naive
read loses all URLs.

Rewrite links in place, read the text, then restore the DOM:

```python
page.evaluate("""
() => {
  const main = document.querySelector('main') || document.body;
  const links = Array.from(main.querySelectorAll('a[href^="http"]'))
    .filter(a => !(a.getAttribute('href') || '').includes('google.com'));
  const originals = links.map(a => a.textContent);
  for (const a of links) {
    a.textContent = `${a.textContent.trim()} (${a.getAttribute('href')})`;
  }
  const hidden = main.querySelectorAll('script, style');
  const display = Array.from(hidden).map(el => el.style.display);
  hidden.forEach(el => { el.style.display = 'none'; });
  const text = main.innerText || '';
  hidden.forEach((el, i) => { el.style.display = display[i]; });
  links.forEach((a, i) => { a.textContent = originals[i]; });
  return text;
}
""")
```

Two traps: hide `<script>`/`<style>` or their source leaks into the text,
and mutate the **attached** DOM — an unattached clone falls back to
`textContent` and leaks script source anyway.

---

## 6. Rate limiting: pacing is the only real lever

Google throttles by **request rate per IP**. Space queries with jitter:

```python
_DEFAULT_DELAY = 10.0   # seconds
_JITTER        = 5.0

def pace(last_at: float) -> float:
    target = _DEFAULT_DELAY + random.uniform(0, _JITTER)
    elapsed = time.monotonic() - last_at
    if last_at and elapsed < target:
        time.sleep(target - elapsed)
    return time.monotonic()
```

Restart the clock from the **end** of each query, so generation time counts
toward the gap instead of adding to it.

Measured guidance:

| Delay | Behaviour |
|---|---|
| 0 s | CAPTCHA within a few queries |
| 10 s | ~56 queries before a CAPTCHA |
| 25 s | recommended for long unattended runs |

**Also: never run two scrapers at once.** Two processes hitting Google from
one IP was the direct cause of repeated throttling in this project.

### URL length limit

The whole prompt rides in `?q=`, and percent-encoding inflates it. Google
returns **HTTP 400 "malformed or illegal request"** past roughly 8 KB.

Cap the final URL at **~7,800 characters** and truncate with a loud
warning. Watch any prompt section that grows per iteration — an exclusion
list of 148 names produced a 10,701-char URL and every request failed
mid-run.

---

## 7. Three distinct failure modes — do not conflate them

This caused a chain of bugs in which each fix created the next. An empty
result has **three** causes needing **opposite** responses:

| Cause | Evidence | Correct response |
|---|---|---|
| **Rate limited** | `/sorry/` in URL; "unusual traffic" | long cool-off (5+ min) — waiting helps |
| **Model refused** | "no response available for this search" | retry now, shorter/reworded prompt — waiting is useless |
| **Genuinely nothing** | neither of the above | stop |

```python
CAPTCHA_MARKERS = ("unusual traffic", "detected unusual traffic",
                   "/sorry/index", "our systems have detected")

REFUSAL_MARKERS = ("no response available for this search",
                   "try asking something else", "i can't help with that",
                   "couldn't generate a response")
```

Guard the refusal check with a length limit (`len(text) < 600`) so a long
real answer that happens to contain the phrase is not discarded.

### CAPTCHA handling: poll, then re-send

```python
def wait_for_captcha_clear(page, timeout=60) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        page.wait_for_timeout(2000)
        if "/sorry/" not in page.url:
            return True      # solved -> caller MUST re-issue the query
    return False
```

Three bugs to avoid, all of which shipped here first:

1. **Blind `sleep(60)` then `continue` without re-navigating.** The next
   attempt re-reads the stale `/sorry/` page, so a *successfully solved*
   CAPTCHA still returns nothing. Always re-issue `page.goto()`.
2. **Letting a solve consume the retry budget.** With `retries=1`, attempt 0
   hit the CAPTCHA and attempt 1 read the stale page. Give CAPTCHA rounds
   their own allowance.
3. **Using the consecutive-CAPTCHA counter to detect blocking.** It resets
   on a successful solve, so a solved CAPTCHA left *no evidence* and the
   caller reported "no results, market exhausted". Track a separate
   `last_captcha_at` timestamp that a solve does not clear.

Log the evidence with the conclusion — `"no sign of blocking (3 CAPTCHAs
this run, last 41 s ago)"` — so a wrong inference is visible rather than
merely asserted.

---

## 8. Loading a CAPTCHA-solver extension (captcha-raptor)

**Read the limitation first, or you will waste a day.** A solver extension
only helps when Google renders a *solvable* reCAPTCHA widget. captcha-raptor
injects **only** into `*://*.google.com/recaptcha/*` and
`*://*.recaptcha.net/recaptcha/*` (its own manifest), so on the `/sorry/`
page that reports **"Cannot contact reCAPTCHA"** no widget ever loads and
there is no URL for it to match. That block is prevented by pacing (§6),
not solved by an extension.

It *does* help when Google shows a real checkbox or image-tile challenge:
the run continues unattended instead of waiting for a human.

### 8.1 The extension

[captcha-raptor / Captcha Plugin](https://captchaplugin.com) — free, MV3,
runs ONNX models locally (no API key, no data leaves the machine).

| Fact | Value |
|---|---|
| Name | `Free AI ReCaptcha Solver by Raptor (Captcha Plugin)` |
| Manifest | v3, service worker `background.js` |
| Injects into | `*://*.google.com/recaptcha/*`, `*://*.recaptcha.net/recaptcha/*` |
| Permissions | `alarms`, `declarativeNetRequest`, `idle`, `storage` |
| Min Chrome | 111 |

**The folder you point at must directly contain `manifest.json`.** Most
extension repos keep it in a subfolder — captcha-raptor's is in
`captcha-raptor/extension/`, others use `public/`, `src/` or `dist/`.
Pointing at the repo root loads **nothing, silently**.

### 8.2 The blocker: managed Chrome ignores `--load-extension`

Measured on Chrome 152, throwaway profiles, ground truth read from
`chrome://extensions-internals`:

| Browser | Flags | Loaded? |
|---|---|---|
| Real Chrome | `--load-extension` | **No** |
| Real Chrome | `--load-extension` + `--disable-extensions-except` | **No** |
| Bundled Chromium | same flags | **Yes** |

Cause: enterprise-managed Chrome refuses sideloaded unpacked extensions.
Check for it:

```powershell
Get-ChildItem "HKLM:\SOFTWARE\Policies\Google\Chrome"
# any ExtensionInstall* key => --load-extension will be ignored
```

So **extensions force a build switch to Chromium**, and that is a real
trade-off: Chromium is CAPTCHA'd noticeably more often than real Chrome.
Decide deliberately.

```python
def resolve_channel(extensions: list[str]) -> str | None:
    """Chromium when extensions are wanted, else real Chrome."""
    if extensions and managed_chrome_blocks_extensions():
        return None          # None => Playwright's bundled Chromium
    return "chrome"
```

Chromium must also get its **own profile directory** — Chrome and Chromium
have incompatible profile state and cookie encryption, so sharing one
corrupts it:

```python
def profile_dir_for(channel: str | None) -> str:
    base = "./data/ai_mode_profile"
    return base if channel else f"{base}_chromium"
```

### 8.3 Install it so it survives (the part that actually works)

`--load-extension` is a **command-line-only** load: Chrome never records it
in the profile, so it disappears on any profile rebuild and cannot be
managed from `chrome://extensions`. Copy the extension to a directory you
control and load it from there every launch.

Keep that staging directory a **sibling** of the profile, never inside it —
inside, a profile reset has to move it out and back, and Chrome/Chromium
profiles cannot share it:

```python
import shutil
from pathlib import Path

_GENERIC = ("extension", "extensions", "src", "public", "dist", "build")

def staging_root(profile_dir: str) -> Path:
    base = Path(profile_dir)
    return base.parent / f"{base.name}_extensions"      # SIBLING, not child

def install_extension(source: str, profile_dir: str) -> bool:
    """Copy an unpacked extension next to the profile so it persists."""
    src = Path(source.strip().strip('"'))
    if not (src / "manifest.json").is_file():
        nested = src / "extension"                      # common repo layout
        if not (nested / "manifest.json").is_file():
            print(f"no manifest.json in {src} or {src}/extension")
            return False
        src = nested

    # Name the copy after the EXTENSION, not the folder: selecting
    # `captcha-raptor/extension` would otherwise stage it as "extension",
    # colliding with the next repo that uses the same layout.
    name = src.name
    if name.lower() in _GENERIC:
        name = src.parent.name or name

    dest = staging_root(profile_dir) / name
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        shutil.rmtree(dest, ignore_errors=True)
    shutil.copytree(src, dest)
    print(f"installed: {dest}")
    return True

def installed_extensions(profile_dir: str) -> list[str]:
    root = staging_root(profile_dir)
    if not root.is_dir():
        return []
    return [str(c.resolve()) for c in sorted(root.iterdir())
            if (c / "manifest.json").is_file()]
```

Copying rather than referencing the download folder means the original can
be moved or deleted without breaking the automation profile.

### 8.4 Launch with it

```python
extensions = installed_extensions(profile_dir)
channel    = resolve_channel(extensions)
profile    = profile_dir_for(channel)

args = ["--profile-directory=Default"]
if extensions:
    joined = ",".join(extensions)
    args.append(f"--load-extension={joined}")
    # Do NOT add --disable-extensions-except unless you mean it: it
    # disables every extension the profile itself installed, which makes
    # chrome://extensions look EMPTY and "Load unpacked" appear broken.

ctx = pw.chromium.launch_persistent_context(
    profile,
    channel=channel,        # None => Chromium (required for extensions)
    headless=False,         # MV3 workers are unreliable headless
    args=args,
)
```

Two flag notes:

- **`headless=False` is required.** MV3 service workers do not run reliably
  in headless Chrome; a solver that never starts is worse than none.
- **Skip `--disable-extensions-except`.** It hides the profile's own
  extensions, and it is what makes manual "Load unpacked" appear to fail.

### 8.5 Verify it loaded — against ground truth

**Do not count service workers.** Chrome always runs its own component
extensions, so `len(context.service_workers)` is never zero:

| Component extension | ID |
|---|---|
| Chrome PDF Viewer | `mhjfbmdgcfjbbpaeojofohoefgiehjai` |
| Google Hangouts | `nkeimhogjdpnpccoofpliimaahmaaome` |
| Google Network Speech | `fignfifoniblkonapihmkfakmlgkbkcf` |

> A never-loaded extension was reported as **"active: 1 background worker"**
> for several rounds here, because that one worker was Chrome's own network
> speech component. The bare count is a false positive.

Two reliable checks. First, ask Chrome directly:

```python
page = ctx.new_page()
page.goto("chrome://extensions-internals/", wait_until="load", timeout=20_000)
rows = [e for e in json.loads(page.inner_text("body"))
        if e.get("location") != "COMPONENT"]
for e in rows:
    print(e["name"], e["location"])     # expect: ... [COMMAND_LINE]
assert rows, "extension is genuinely not installed"
```

Second, attribute the worker by deriving the extension's ID. Chrome computes
it from the **absolute path**: SHA-256, first 16 bytes, each nibble mapped
`0-f -> a-p`. On Windows the path is hashed as **UTF-16LE with the drive
letter uppercased**:

```python
import hashlib, os
from pathlib import Path

def unpacked_extension_id(path: str) -> str:
    raw = str(Path(path).resolve())
    if os.name == "nt":
        if len(raw) > 1 and raw[1] == ":":
            raw = raw[0].upper() + raw[1:]
        digest = hashlib.sha256(raw.encode("utf-16-le")).hexdigest()
    else:
        digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()
    return "".join(chr(ord("a") + int(c, 16)) for c in digest[:32])

def our_workers(ctx, paths) -> list[str]:
    want = {unpacked_extension_id(p) for p in paths} - {""}
    urls = [w.url for w in ctx.service_workers] + \
           [p.url for p in ctx.background_pages]
    return [u for u in urls if any(i in u for i in want)]
```

A correct load looks like this — the ID matches the derivation, and Chrome
names the extension itself:

```
chrome-extension://<derived-id>/background.js  <-- ours
chrome://extensions-internals: Free AI ReCaptcha Solver by Raptor [COMMAND_LINE]
```

The ID is path-dependent, so it changes if you move the staged folder —
always compare against `unpacked_extension_id(path)` rather than a
hardcoded literal.

MV3 workers register **lazily** — poll for ~8 s before concluding failure;
a 2 s check produces false negatives.

### 8.6 Adding one by hand

Manual installs persist only if the profile is not rebuilt and
`--disable-extensions-except` is absent:

```python
page.goto("chrome://extensions/")
# 1. toggle Developer mode (top right)
# 2. Load unpacked
# 3. pick the folder CONTAINING manifest.json
```

If that page looks empty, `--disable-extensions-except` is set — that is the
cause, not a broken profile.

### 8.7 Checklist

```
[ ] Point at the folder CONTAINING manifest.json (often extension/, public/, src/)
[ ] Check HKLM\SOFTWARE\Policies\Google\Chrome for ExtensionInstall* policy
[ ] If managed: switch to Chromium (channel=None) + its own profile dir
[ ] Copy the extension to a SIBLING dir of the profile; load from there
[ ] headless=False
[ ] Omit --disable-extensions-except
[ ] Verify via chrome://extensions-internals, filtering location==COMPONENT
[ ] Poll ~8 s for the MV3 worker
[ ] Accept: it cannot clear "Cannot contact reCAPTCHA" — pacing does that
```


## 9. Prompting AI Mode

It has no JSON mode, but reliably follows an explicit instruction.

```
Reply with ONLY a JSON array, no markdown fences and no commentary:
[{"name": "", "website": ""}]
```

Parsing the reply needs three defences:

```python
text = raw.replace("Use code with caution.", "")        # Google's UI footer
if text.lstrip().lower().startswith("json"):            # stray language tag
    text = text.lstrip()[4:]

# AI Mode ECHOES the question. If your prompt ends with an empty JSON
# template, find("[") locks onto the echo and yields zero results.
# Take the LAST array that starts with a real key:
starts = [m.start() for m in re.finditer(r'\[\s*\{\s*"name"', text)]
if starts:
    text = text[starts[-1]:]

candidate = text[text.find("["):text.rfind("]") + 1]
try:
    data = json.loads(candidate)
except json.JSONDecodeError:                            # literal newlines in strings
    data = json.loads(re.sub(r"\s*\n\s*", " ", candidate))
```

> **Slice *from* the marker, not past it.** An off-by-one here silently
> dropped the first item of every batch — found only by a failing test.

### Never instruct the model to fill every field

If a prompt says *"no field may be blank"*, the model **invents** values to
comply. Measured in this project: **94%** of one email column came back as
`info@<their-own-domain>`, LinkedIn URLs were built from people's names, and
phone numbers were generated from country calling codes.

Invert it:

```
An empty string is CORRECT and preferred whenever you do not genuinely
know a value. NEVER construct an email from a domain. NEVER build a
LinkedIn URL from a person's name. NEVER invent a phone number.
A wrong value is much worse than an empty one.
```

Then reject constructed shapes in code — a plausible invention cannot be
distinguished from real data downstream, so it is worse than a blank. Also
avoid "repair" helpers that manufacture data (rebuilding a LinkedIn URL
from a name, copying one field into another): prefer rejecting a malformed
value to fixing it into something indistinguishable from real data.

**A field that is 94% one pattern was generated, not researched.** Checking
a column's frequency distribution is a fast fabrication test.

---

## 10. Porting checklist

```
[ ] pip install patchright   (NOT playwright)
[ ] Copy google_ai_mode.py — it has no project-internal imports
[ ] Import from patchright.sync_api
[ ] channel="chrome", headless=False, launch_persistent_context
[ ] Profile dir outside the repo (or .gitignore it — it holds cookies)
[ ] Completion: marker -> Thinking-cleared -> text-stable. Never length
[ ] Extract via evaluate() with links inlined
[ ] Pace 10–25 s with jitter; restart the clock at query END
[ ] Cap the URL at 7,800 chars
[ ] Handle CAPTCHA / refusal / empty as three separate cases
[ ] Re-issue page.goto() after a solved CAPTCHA
[ ] One scraper per IP at a time
[ ] Reuse one browser per process (launch costs seconds; warm = fewer CAPTCHAs)
[ ] Prompt: "" is preferred over a guess
[ ] Tests: force the backend OFF by default so a missed mock fails loudly
```

### Environment variables used here

| Variable | Default | Purpose |
|---|---|---|
| `GOOGLE_AI_MODE_ENABLED` | `false` | Master switch |
| `GOOGLE_AI_MODE_HEADLESS` | `false` | Keep false for CAPTCHAs |
| `GOOGLE_AI_MODE_BROWSER` | `chrome` | `chromium` for the bundled build |
| `GOOGLE_AI_MODE_QUERY_DELAY` | `10` | Seconds between queries |
| `GOOGLE_AI_MODE_PROFILE_DIR` | `data/ai_mode_chrome_profile` | Persistent profile |
| `GOOGLE_AI_MODE_EXTENSIONS` | *(unset)* | `;`-separated unpacked paths |
| `GOOGLE_AI_MODE_MAX_FAILURES` | `5` | Consecutive failed queries before an automatic session reset |
| `GOOGLE_AI_MODE_MAX_RESETS` | `3` | Cap on automatic resets per process |
| `GOOGLE_AI_MODE_RESET_COOLOFF` | `60` | Seconds to wait after a reset before the first query |

### Automatic session recovery

A session can be alive but persistently failing — every query CAPTCHA'd,
refused, timing out, or returning an empty answer pane. Relaunching the
same profile inherits the problem, because Google's rate-limit flag rides
on the profile's **cookies**.

`reset_session()` closes the browser, deletes the profile, and re-seeds it
from the real Chrome profile, giving the run a fresh identity. It fires
automatically after `GOOGLE_AI_MODE_MAX_FAILURES` consecutive failures,
then retries the prompt that tripped it.

Three details that make it safe:

- **A successful answer clears the streak**, so intermittent failures never
  accumulate into a needless reset.
- **Resets are capped.** A broken environment (no network, Chrome missing)
  would otherwise loop forever rebuilding profiles and hide the real error.
- **Installed extensions survive** — they live in a sibling directory
  (§8.3), specifically so recovery cannot cost the user their solver.

The consecutive-CAPTCHA circuit breaker also routes through it now. It
previously returned `""` *before* the retry loop, which meant a fully
flagged profile — the exact case a reset fixes — could never trigger
recovery.

### Security notes

- The profile directory holds **real cookies and session tokens**. Keep it
  out of version control, and delete any temporary copy immediately after
  use.
- Never pass `--log-path` to chromedriver while debugging: it dumps full
  Chrome Preferences including live session cookies in plaintext.

### Testing

Force the backend off in a session-wide fixture, so a test that forgets a
mock fails on an assertion instead of silently driving a real browser:

```python
@pytest.fixture(autouse=True)
def _no_real_browser(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_ENABLED", "false")
```

*This exact fixture exists because a real `.env` flag once made the suite
launch Chrome windows and hang indefinitely.*

---

## 11. Drop-in OpenAI-shaped adapter

Wrap `ask()` so existing callers need no changes:

```python
def chat(system: str, user: str, **_ignored) -> str:
    """`model`, `max_tokens`, `temperature` are accepted and IGNORED —
    AI Mode is a web UI, not a metered API. Keeping them in the signature
    is what makes this a true drop-in swap."""
    prompt = f"{system.strip()}\n\n{user.strip()}" if system.strip() else user.strip()
    text = ask(prompt)
    if not text:
        raise RuntimeError("AI Mode unavailable (CAPTCHA, timeout, or refusal)")
    return text

def chat_json(system: str, user: str, **kw):
    return parse_json_answer(chat(system, user, **kw))
```

---

## Appendix: bugs this project hit, in order

Each was real, each cost time, and several were caused by the previous fix.

| Symptom | Root cause |
|---|---|
| Instant CAPTCHA on every query | plain Playwright instead of Patchright |
| ~94 s per query | completion gated on >400 chars of answer text |
| All URLs missing from answers | `innerText` drops `href`s |
| HTTP 400 mid-run | 10,701-char URL from a growing exclusion list |
| First item of every batch dropped | echo-stripper sliced *past* the marker |
| Solved CAPTCHA changed nothing | no `page.goto()` re-issue after the solve |
| "Market exhausted" while CAPTCHA'd | blocking detected via a counter a solve resets |
| 300 s cool-off that never helped | model *refusal* misclassified as a block |
| Extension "active" but never ran | counted Chrome's own component workers |
| 94% of a column identical | prompt forbade blank values, so the model invented them |

### The meta-lesson

Every one of those was found by measuring, not reasoning. When something
fails, read the authoritative state — the actual URL, the actual
`chrome://extensions-internals` list, the actual column distribution —
rather than the nearest correlated signal. Several of the bugs above
persisted precisely because a proxy signal (a worker count, a CAPTCHA
counter, an answer length) looked healthy while the real thing was broken.
