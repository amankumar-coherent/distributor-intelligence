#!/usr/bin/env python3
"""Classify each unique company's role via Google AI Mode.

Reads cards_unique_companies.xlsx, asks AI Mode one question per company
(mirroring the query the user ran by hand for TigerBoost), and writes a
"Role" column holding a single word: Brand, Manufacturer, Distributor,
Trader, Supplier, Retailer, Service, or Unknown.

WEBSITE IN THE QUERY: the prompt names the company together with its
website, because "TigerBoost" alone is ambiguous while
"TigerBoost (tigerboost.es)" is not. The website comes from the Website
column; when that cell is empty it is derived from the email domain --
exactly the TigerBoost case (robert@tigerboost.es -> tigerboost.es).
Free mail domains (gmail, 163, qq...) are NOT usable as a company
website, so those companies are asked by name alone rather than being
given a wrong domain.

NO GUESSED ROLES -- this is the point of the script, not a nicety. Every
Role cell is read out of what AI Mode actually said about that company;
nothing is inferred from the company's name, its country, its email
domain, or what similar companies usually do. "Unknown" is a CORRECT and
preferred answer: a blank can be re-queried or researched by hand, whereas
a plausible invented role is indistinguishable from a real one once it is
in the spreadsheet. Concretely, a role is recorded only when:

  - the model gave its own labelled verdict ("Role: Brand"), or a bare
    one-word verdict line; failing that,
  - a sentence PREDICATED OF THIS COMPANY names a role ("X is a
    manufacturer"). Sentences about the industry ("nutraceutical
    ingredients are made by manufacturers") or about customers ("supplies
    to manufacturers") are excluded.

and never when the model disclaimed knowledge ("I could not find
information about this company") -- role words in such an answer describe
the sector, not the firm, so the result is Unknown.

WHAT THE SCRIPT CANNOT GUARANTEE: that the answer is about the RIGHT
company. Asked by bare name, AI Mode sometimes answers about a
same-or-similar-named firm (verified live: "Xi'an Zhengfa" returned an
answer about Zhejiang Jiande Zhengfa). Parsing cannot detect this, so the
"Check" column in the Role Audit sheet flags every row where the model
restated the name, and every row matched by name alone. Those rows need a
human look before the role is trusted.

Progress is checkpointed to data/company_roles/roles_checkpoint.json after
every company, so an interrupted run (CAPTCHA, closed browser, Ctrl-C)
resumes where it stopped instead of re-querying from company 1.

Usage:
    python run_company_roles.py                 # run/resume all 530
    python run_company_roles.py --limit 10      # smoke-test on 10
    python run_company_roles.py --export-only   # rebuild xlsx from checkpoint
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

IN_XLSX = ROOT / "cards_unique_companies.xlsx"
OUT_XLSX = ROOT / "cards_unique_companies_with_role.xlsx"
STATE_DIR = ROOT / "data" / "company_roles"
STATE = STATE_DIR / "roles_checkpoint.json"

# Mail providers whose domain says nothing about the company, so it must
# never be presented to the model as "the company's website".
FREE_MAIL = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.co.jp", "yahoo.com.tw",
    "yahoo.co.in", "yahoo.com.cn", "hotmail.com", "outlook.com", "live.com",
    "msn.com", "aol.com", "icloud.com", "me.com", "qq.com", "163.com",
    "126.com", "yeah.net", "vip.163.com", "sina.com", "sina.cn", "sohu.com",
    "foxmail.com", "aliyun.com", "naver.com", "daum.net", "hanmail.net",
    "nate.com", "yandex.ru", "mail.ru", "rambler.ru", "rediffmail.com",
    "gmx.de", "gmx.net", "web.de", "protonmail.com", "hotmail.co.jp",
    "hotmail.com.tw", "yahoo.fr", "yahoo.de", "outlook.jp",
}

# One word out, ranked so that a sentence naming several roles resolves to
# the most specific one the answer actually asserts.
ROLE_WORDS = [
    ("Manufacturer", r"\bmanufactur\w*\b|\bmaker\b|\bproducer\b|\bOEM\b"),
    ("Distributor", r"\bdistribut\w*\b"),
    ("Brand", r"\bbrand\b"),
    ("Trader", r"\btrader\b|\btrading\b"),
    ("Retailer", r"\bretail\w*\b|\be-?commerce\b|\bmarketplace\b"),
    ("Supplier", r"\bsupplier\b|\bsupplies\b"),
    ("Service", r"\bservice\w*\b|\bconsult\w*\b|\bCRO\b|\bCDMO\b|\blaboratory\b"),
]

_QUESTION = (
    " what is the function of this company and after function definition "
    "also tell me what is the role of the company it is brand, manufacture, "
    "distributor or what? only in one word tell me what is the role of company"
)

# With a website the company is unambiguous, so the name plus domain is
# all the model needs.
PROMPT = "{name} ({site})" + _QUESTION

# Without a domain the bare name is ambiguous -- "INGRAM" or "BJ Inter"
# could be any industry, and the model would be guessing at which company
# is meant. Naming the industry is context we actually have (every row in
# this file is a nutraceutical/ingredients trade-show card), not an
# invented fact about the company, and it steers the lookup to the right
# firm instead of a same-named one in another sector.
PROMPT_NO_SITE = "{name} it is a nutraceutical company," + _QUESTION


def website_for(website, email):
    """Return (website, source). Empty website means none is usable."""
    web = (website or "").strip() if isinstance(website, str) else ""
    web = re.sub(r"^https?://", "", web).strip().strip("/")
    if web:
        return web, "website-column"
    em = (email or "").strip() if isinstance(email, str) else ""
    m = re.search(r"@([A-Za-z0-9.-]+\.[A-Za-z]{2,})", em)
    if m:
        dom = m.group(1).lower().rstrip(".")
        if dom not in FREE_MAIL:
            return dom, "email-domain"
    return "", "none"


# Lines AI Mode wraps around the real answer: the Google nav bar above it
# and the citation/disclaimer block below. Matching a role word inside any
# of these produces a verdict the model never gave, so they are dropped
# before parsing. Verified against live answers: the nav block ends at the
# echoed prompt, and the answer ends at the disclaimer.
_PROMPT_ECHO = re.compile(r"^AI Mode conversation:.*$", re.I)
_TAIL_MARKERS = (
    "AI can make mistakes",
    "AI Mode response is ready",
    "double-check responses",
)
# The model labels its one-word verdict, e.g. "Role\nBrand" or
# "Role: Brand" (often with a leading emoji). That label is the verdict --
# far more reliable than guessing from position in the text.
# "Role: Brand" or "Role\nBrand". The negative lookahead keeps it from
# matching a section HEADING like "Role of the Company", which announces
# the discussion rather than giving the verdict -- that heading previously
# consumed the match and let a wrong role through.
_ROLE_LABEL = re.compile(
    r"(?:^|\n)[^\w\n]*Role\b(?!\s+(?:of|in|as|for|and)\b)\s*[:\-]?\s*\n?\s*"
    r"([A-Za-z][A-Za-z /&-]{2,40})"
)

# The model saying it does not know this company. Any role word in such an
# answer describes the industry in general, so no role may be recorded.
_NO_KNOWLEDGE = re.compile(
    r"\b(?:"
    r"could not find|couldn't find|cannot find|can't find|no (?:specific |public |reliable |verifiable )?"
    r"information|not (?:enough|much) information|unable to (?:find|verify|locate|determine)|"
    r"no (?:results|record|records|details|data) (?:found|available)|"
    r"does not appear to (?:exist|be)|there is no (?:widely|publicly|readily)|"
    r"i (?:do not|don't) have (?:any )?(?:specific )?information|"
    r"limited(?:ly)? (?:public )?information|not clearly (?:documented|established)"
    r")\b",
    re.I,
)

# The model saying the entity is not a commercial company at all (a
# university, government body, trade association). Verified live on
# "Prince of Songkia University", where the answer said it is "a public
# educational and research institution ... not a commercial corporation"
# and the parser still returned Distributor.
_NOT_A_COMPANY = re.compile(
    r"\b(?:not a (?:commercial|private|for-profit|manufacturing|trading)?\s*"
    r"(?:corporation|company|business|firm|entity|manufacturer|distributor|brand)|"
    r"is (?:actually )?a (?:public |state |national )?"
    r"(?:university|educational|academic|research institution|government|"
    r"non-?profit|trade association|industry association)|"
    r"rather than a (?:commercial|company|business|corporation)|"
    r"if forced to choose)\b",
    re.I,
)

# A clause that predicates something OF THE COMPANY, rather than of its
# customers or its market. Matches "X is a manufacturer", "the company
# operates as a distributor", "they manufacture", "it produces".
_IS_ABOUT_SUBJECT = re.compile(
    r"\b(?:is|acts? as|operates? as|functions? as|serves? as|"
    r"remains?|became?|considered|classified as|primarily|mainly|"
    r"the company|this company|they|its role)\b",
    re.I,
)

# Sentences about the market/industry at large rather than this company.
# "Nutraceutical ingredients are made by manufacturers worldwide" states a
# fact about the sector; scoring it would attribute an industry-wide role
# to whichever company happened to be asked about.
_GENERIC_SUBJECT = re.compile(
    r"\b(?:companies|firms|manufacturers|distributors|suppliers|brands|"
    r"products|ingredients|the (?:industry|sector|market))\b\s+(?:are|is|can|"
    r"typically|generally|usually|often|worldwide|globally)",
    re.I,
)


def strip_page_chrome(answer: str, prompt: str = "") -> str:
    """Return just the model's answer, without Google's surrounding UI text."""
    text = (answer or "").replace("\xa0", " ").strip()
    if not text:
        return ""
    lines = text.splitlines()
    # Everything up to and including the echoed prompt is Google's chrome.
    # The echo sometimes carries an "AI Mode conversation:" prefix and
    # sometimes appears bare, so match the prompt text itself as well.
    # Google echoes the prompt TWICE -- once prefixed "AI Mode
    # conversation:" and once bare -- so this must not stop at the first
    # match. A leftover echo is not cosmetic: the prompt itself contains
    # "brand, manufacture, distributor", which the prose fallback would
    # then score as if the model had said it.
    start = 0
    needle = " ".join(prompt.split())[:60].lower()
    for i, line in enumerate(lines):
        flat = " ".join(line.split()).lower()
        if _PROMPT_ECHO.match(line.strip()) or (needle and needle in flat):
            start = i + 1
    body = lines[start:]
    # Cut at the first disclaimer/citation marker.
    for i, line in enumerate(body):
        if any(mark.lower() in line.lower() for mark in _TAIL_MARKERS):
            body = body[:i]
            break
    return "\n".join(body).strip()


# Google's own failure page, returned INSTEAD of an answer. It is not an
# empty string, so an "answer is blank" check sails straight past it and
# records Unknown as though the model had considered the company and had
# nothing to say. Those are opposite situations: a real Unknown is a
# result to keep, a backend failure is a query that must be retried.
_BACKEND_FAILURE = re.compile(
    r"something went wrong|an AI response wasn't generated|"
    r"couldn't generate a response|try again later|"
    r"AI responses? (?:may be|are) unavailable|"
    # Google's RATE-LIMIT notice, which is a distinct message from the
    # generic error page above and was missed by it: "You've reached the
    # request limit for AI responses at the moment. Try again in a little
    # while." Two companies were saved as Unknown off this text before it
    # was added -- the query was never answered at all.
    r"reached the request limit|request limit for AI responses|"
    r"try again in a little while|"
    r"you(?:'ve| have) (?:reached|hit) (?:the|your) limit",
    re.I,
)


def is_backend_failure(answer: str, prompt: str = "") -> bool:
    """True when Google returned an error page rather than an answer."""
    body = strip_page_chrome(answer, prompt)
    return bool(body) and bool(_BACKEND_FAILURE.search(body))


def parse_role(answer: str, prompt: str = "") -> str:
    """Pull the one-word role out of AI Mode's reply, or 'Unknown'.

    Reads the model's own labelled verdict ("Role: Brand") when present,
    since that is what the prompt asks for. Only if there is no label does
    it fall back to scanning the prose -- and even then it must not pick up
    'manufacturers' from a description of WHO THE COMPANY SELLS TO, which
    is why the fallback prefers a short standalone verdict line.
    """
    body = strip_page_chrome(answer, prompt)
    if not body:
        return "Unknown"

    m = _ROLE_LABEL.search(body)
    if m:
        for role, pattern in ROLE_WORDS:
            if re.search(pattern, m.group(1), re.I):
                return role

    # A bare verdict line, e.g. "Manufacturer." on its own (BIO ACTIVES
    # put it first, so position alone can't be trusted).
    for line in body.splitlines():
        stripped = line.strip().strip(".*_# ")
        if 3 <= len(stripped) <= 24:
            for role, pattern in ROLE_WORDS:
                if re.fullmatch(pattern, stripped, re.I):
                    return role

    # A line that OPENS with the verdict and then qualifies it, e.g.
    # "Distributor (while the parent group handles manufacturing...)" or
    # "Supplier (specifically an ingredient supplier)." The leading word is
    # the answer; the qualifier often names other roles, so ranking the
    # whole line by ROLE_WORDS order picks the wrong one.
    for line in body.splitlines():
        opener = line.strip().strip("*_# ")
        for role, pattern in ROLE_WORDS:
            if re.match(rf"(?:{pattern})\s*(?:[(.,:\-]|$)", opener, re.I):
                return role

    # Last resort: the model gave no labelled verdict and no standalone
    # one-word line, so any role here has to be read out of free prose.
    # That is where a fabricated role would come from, so it is fenced in
    # two ways.
    #
    # 1. If the answer disclaims knowledge of the company, return Unknown
    #    even though role words appear -- "I could not find this company,
    #    but nutraceutical distributors generally..." describes the
    #    INDUSTRY, not this firm.
    if _NO_KNOWLEDGE.search(body):
        return "Unknown"

    # The model rejecting the premise: the entity is a university, an
    # institute, an association -- not a commercial company at all. Its
    # answer then discusses commercial terms only to say none of them fit
    # ("if forced to choose from brand, manufacturer, or distributor...").
    # Scoring that would put a trading role on a research institution.
    if _NOT_A_COMPANY.search(body):
        return "Unknown"

    # 2. Only accept a role from a sentence that is actually about this
    #    company ("X is a manufacturer of..."), not one about its customers
    #    or its market. Take the last such sentence, since the prompt puts
    #    the verdict after the description.
    best, best_pos = None, -1
    for sentence in re.split(r"(?<=[.!?])\s+|\n", body):
        if _GENERIC_SUBJECT.search(sentence):
            continue
        if not _IS_ABOUT_SUBJECT.search(sentence):
            continue
        for role, pattern in ROLE_WORDS:
            hits = list(re.finditer(pattern, sentence, re.I))
            if hits and body.find(sentence) + hits[-1].start() > best_pos:
                best, best_pos = role, body.find(sentence) + hits[-1].start()
    # No qualifying sentence -> no role. An Unknown is honest and can be
    # re-queried; a role guessed from surrounding prose is indistinguishable
    # from a researched one once it is in the spreadsheet.
    return best or "Unknown"


# The model correcting the name it was given, e.g. "the company you are
# referring to is accurately registered as ...". Seen live on Xi'an
# Zhengfa, where the answer was about a DIFFERENT firm in another
# province. The role may be perfectly real and still belong to the wrong
# company, which no amount of parsing can detect -- so it is surfaced for
# human review rather than silently accepted or silently dropped.
_NAME_SUBSTITUTED = re.compile(
    r"\b(?:you are referring to is|actually (?:known|registered|named) as|"
    r"correct name is|properly known as|refers to|also known as|"
    r"appears to be a (?:reference|typo)|did you mean)\b",
    re.I,
)


def answer_warning(company, rec) -> str:
    """Flag rows whose role needs a human look. Empty means no concern."""
    if not rec:
        return ""
    role = rec.get("role") or ""
    body = strip_page_chrome(rec.get("answer") or "", rec.get("prompt") or "")
    notes = []
    if role == "Unknown":
        notes.append("no role stated" if body else "no answer returned")
    if body and _NAME_SUBSTITUTED.search(body):
        notes.append("VERIFY: answer may describe a different company")
    if rec.get("website_source") == "none" and role != "Unknown":
        notes.append("no website - matched by name only")
    return "; ".join(notes)


def load_state() -> dict:
    if STATE.exists():
        try:
            return json.loads(STATE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_state(state: dict) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")


def read_rows():
    import openpyxl

    wb = openpyxl.load_workbook(IN_XLSX, read_only=True)
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(max_row=1))]
    return header, list(ws.iter_rows(min_row=2, values_only=True))


def export(header, rows, state) -> Path:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Unique Companies"
    # "Website Used" is deliberately NOT on this sheet -- it duplicates the
    # existing Website column for most rows. It stays on the Role Audit
    # sheet, which is where the record of what was actually queried belongs
    # (it is the only place showing the email-domain fallback).
    ws.append(list(header) + ["Role"])
    audit = wb.create_sheet("Role Audit")
    audit.append([
        "serial number", "Company", "Website Used", "Website Source", "Role",
        "Check", "Prompt Sent", "AI Mode Answer",
    ])

    for r in rows:
        key = str(r[0])
        rec = state.get(key) or {}
        ws.append(list(r) + [rec.get("role", "")])
        audit.append([
            r[0], r[3], rec.get("website", ""), rec.get("website_source", ""),
            rec.get("role", ""), answer_warning(r[3], rec),
            rec.get("prompt", ""), (rec.get("answer") or "")[:1500],
        ])

    for col, w in zip("ABCDEFGHIJ", [8, 22, 26, 40, 20, 20, 30, 28, 55, 14]):
        ws.column_dimensions[col].width = w
    for col, w in zip("ABCDEFGH", [8, 40, 28, 16, 14, 26, 60, 110]):
        audit.column_dimensions[col].width = w
    wb.save(OUT_XLSX)
    return OUT_XLSX


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--limit", type=int, default=None, help="Only process this many un-answered companies (smoke test)")
    ap.add_argument("--export-only", action="store_true", help="Rebuild the xlsx from the checkpoint without querying")
    ap.add_argument("--redo-unknown", action="store_true", help="Re-query companies whose role came back Unknown")
    ap.add_argument("--reparse", action="store_true", help="Re-score stored answers with the current parser (no queries)")
    ap.add_argument(
        "--real-chrome",
        action="store_true",
        help="Use real Chrome instead of Chromium. Far fewer CAPTCHAs, but "
        "the captcha-raptor extension will NOT load (verified against "
        "chrome://extensions-internals), so nothing auto-solves a CAPTCHA.",
    )
    ap.add_argument(
        "--cooloff-minutes",
        type=int,
        default=30,
        help="On a wall of backend errors (a per-IP limit), pause this long "
        "and resume automatically. 0 = stop instead of waiting.",
    )
    ap.add_argument(
        "--max-cooloffs",
        type=int,
        default=4,
        help="How many times to cool off before giving up for the session.",
    )
    ap.add_argument(
        "--recycle-every",
        type=int,
        default=25,
        help="Restart the browser every N queries (0 = never). The AI Mode "
        "session degrades with use; recycling avoids the error-page runs.",
    )
    ap.add_argument(
        "--fresh",
        action="store_true",
        help="Start over from company 1. The existing checkpoint is renamed "
        "(never deleted) so previous answers stay recoverable.",
    )
    ap.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Seconds to wait between queries (0 = no pacing, the default here). "
        "Raise it if Google starts CAPTCHA-ing the run.",
    )
    args = ap.parse_args()

    apply_env_overrides()

    # Pacing off by default for THIS script only: it is set in-process,
    # after apply_env_overrides() so it wins over .env, and it does not
    # touch the shared scraper's own default that other pipelines rely on.
    # Google's rate limiter is the reason the delay exists, so if answers
    # start coming back empty or CAPTCHA'd, re-run with e.g. --delay 15.
    os.environ["GOOGLE_AI_MODE_QUERY_DELAY"] = str(max(0.0, args.delay))

    # STAY ON REAL CHROME. .env points GOOGLE_AI_MODE_EXTENSIONS at the
    # captcha-raptor extension, and requesting an extension makes the
    # scraper switch from real Chrome to bundled Chromium -- which Google
    # CAPTCHAs and error-pages far more aggressively. Measured directly:
    # under Chromium, saanroo / NAMSIANG / solutexcorp returned
    # "Something went wrong" on every attempt across several runs and
    # browser restarts; with this flag set, the same three queries in the
    # same session returned Supplier / Distributor / Manufacturer with
    # zero CAPTCHAs. The CAPTCHA-solver extension was causing more
    # CAPTCHAs than it solved, so it is disabled for this script only --
    # other pipelines reading .env are unaffected.
    if args.real_chrome:
        # Real Chrome: far fewer CAPTCHAs, but the extension will NOT load.
        # Verified against chrome://extensions-internals (not a worker
        # count, which always includes Chrome's own components): a fresh
        # profile CREATED with the extension staged still reported 0
        # non-component extensions on Chrome, while the same extension
        # reported 1 under Chromium.
        #
        # NOTE: do NOT also clear GOOGLE_AI_MODE_EXTENSIONS here. Tried
        # that to silence the "no extension background worker detected"
        # warning; it does not silence it (the copy staged in the profile
        # loads independently of the env var) and it pushed the scraper
        # down a different code path that raised "Playwright Sync API
        # inside the asyncio loop" and cost a real answer.
        os.environ["GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM"] = "false"
    else:
        # DEFAULT: bundled Chromium, which is the only build here that
        # loads the captcha-raptor solver (1 background worker, confirmed
        # via chrome://extensions-internals as "Free AI ReCaptcha Solver
        # by Raptor [COMMAND_LINE]"). Chosen deliberately for unattended
        # runs: Google challenges Chromium far more often, but the solver
        # can clear a challenge with no human present, and the runner's
        # backend-failure guard never records an error page as a result.
        # Pair this with a non-zero --delay; back-to-back querying on
        # Chromium is what produced the long "Something went wrong" runs.
        os.environ["GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM"] = "true"
    header, rows = read_rows()

    if args.fresh and STATE.exists():
        # Archive rather than delete: a full re-run costs hours of queries,
        # and the raw answers in here can always be re-scored offline with
        # --reparse if only the parser changed.
        import datetime

        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = STATE.with_name(f"roles_checkpoint.{stamp}.json")
        STATE.rename(backup)
        print(f"Previous progress archived to: {backup.name}")

    state = load_state()

    if args.reparse:
        # Raw answers are checkpointed, so a parser improvement never needs
        # the 530 queries to be repeated.
        changed = 0
        for key, rec in state.items():
            new = parse_role(rec.get("answer") or "", rec.get("prompt") or "")
            if new != rec.get("role"):
                rec["role"] = new
                changed += 1
        save_state(state)
        path = export(header, rows, state)
        print(f"Re-parsed {len(state)} stored answers, {changed} roles changed. Exported {path}")
        return

    if args.export_only:
        path = export(header, rows, state)
        print("Exported " + str(path))
        return

    # Refuse to run twice at once. Two instances share one automation
    # profile and one checkpoint file: they fight over the Chrome profile
    # lock (producing spurious "Playwright Sync API inside the asyncio
    # loop" errors), and the second one's save can clobber the first's.
    # Observed live -- a --limit 2 test started alongside a full run
    # reported a company as Unknown that the other run had answered
    # correctly seconds earlier.
    lock = STATE_DIR / "run.lock"
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    if lock.exists():
        age = time.time() - lock.stat().st_mtime
        if age < 900:
            print(
                f"Another run appears to be active (started {age/60:.0f} min ago).\n"
                f"Finish or stop it first, or delete {lock} if it is stale."
            )
            sys.exit(1)
        print("Found a stale lock (>15 min old); continuing.")
    lock.write_text(str(os.getpid()), encoding="utf-8")

    from vendor_intel.scraping import google_ai_mode as aim

    if not aim.ai_mode_enabled():
        print("GOOGLE_AI_MODE_ENABLED is not true in .env -- aborting.")
        sys.exit(1)

    todo = []
    for r in rows:
        rec = state.get(str(r[0]))
        if rec and rec.get("role") and not (args.redo_unknown and rec["role"] == "Unknown"):
            continue
        todo.append(r)
    if args.limit:
        todo = todo[: args.limit]

    already = len(rows) - len(todo)
    print(f"{len(rows)} companies total, {already} already answered, {len(todo)} to query.")
    consecutive_fail = 0
    # The AI Mode session degrades the longer it is used: one page is
    # reused for every query, and past a few dozen queries Google starts
    # returning its error page and CAPTCHA-ing. Restarting the browser
    # clears it (measured: a query failing in the old session succeeded
    # immediately in a fresh one). Recycling on a schedule keeps the run
    # ahead of the degradation instead of paying several failed attempts
    # per company once it sets in.
    since_restart = 0
    # Set once the run gives up on Chromium and switches to real Chrome,
    # so it does not flip back and forth for the rest of the run.
    fell_back = False
    cooloffs_left = args.max_cooloffs

    try:
        for i, r in enumerate(todo, 1):
            serial, company = r[0], (r[3] or "").strip()
            web, source = website_for(r[7], r[6])
            prompt = (
                PROMPT.format(name=company, site=web)
                if web
                else PROMPT_NO_SITE.format(name=company)
            )

            # "Something went wrong" is a POISONED BROWSER SESSION, not a
            # rate limit. Measured: a query that failed repeatedly in the
            # running browser succeeded immediately in a fresh one, with
            # zero CAPTCHAs. The scraper reuses a single page for every
            # query (google_ai_mode._get_page returns pages[0]), so once
            # that page goes bad, every later query fails and retrying
            # inside the same session cannot recover it.
            #
            # So: retry once in-session (cheap, catches a genuine blip),
            # then restart the browser, which is what actually fixes it.
            if since_restart >= args.recycle_every > 0:
                print(f"  -- {since_restart} queries on this browser session, restarting it --")
                try:
                    aim.shutdown()
                except Exception:
                    pass
                time.sleep(5)
                since_restart = 0

            since_restart += 1
            answer = aim.ask(prompt)
            if is_backend_failure(answer, prompt):
                print(f"[{i}/{len(todo)}] #{serial} {company[:44]:44s} backend error -- retrying in 10s")
                time.sleep(10)
                answer = aim.ask(prompt)
            if is_backend_failure(answer, prompt):
                # Restarting alone does NOT fix this on Chromium: measured,
                # the same query kept failing across restart after restart,
                # then succeeded first try on real Chrome. Google serves the
                # error page to Chromium's fingerprint, and the CAPTCHA
                # solver cannot help because the problem is not a CAPTCHA.
                # So the restart also FALLS BACK to real Chrome (losing the
                # extension, which is the right trade -- an unattended run
                # that completes beats one that stalls with a solver loaded).
                if not args.real_chrome and not fell_back:
                    fell_back = True
                    os.environ["GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM"] = "false"
                    print(
                        f"[{i}/{len(todo)}] #{serial} {company[:44]:44s} "
                        "repeated errors -- falling back to real Chrome "
                        "(CAPTCHA solver off; use --real-chrome to start this way)"
                    )
                else:
                    print(f"[{i}/{len(todo)}] #{serial} {company[:44]:44s} still failing -- restarting browser")
                try:
                    aim.shutdown()  # clears _context; next ask() relaunches
                except Exception:
                    pass
                time.sleep(5)
                since_restart = 1
                answer = aim.ask(prompt)

            # A backend failure is NOT a result. Recording it would write
            # Unknown into the sheet as though the model had considered the
            # company, and -- worse -- checkpoint it, so a resumed run would
            # skip the company forever. Leave it unanswered so re-running
            # picks it up.
            if is_backend_failure(answer, prompt):
                consecutive_fail += 1
                print(f"[{i}/{len(todo)}] #{serial} {company[:44]:44s} {web[:28]:28s} -> BACKEND ERROR (not saved)")
                if consecutive_fail >= 5:
                    # A wall of consecutive errors is a CUMULATIVE PER-IP
                    # limit, not a per-query or per-browser fault. Proven:
                    # after ~250 queries in a day, a company that answered
                    # perfectly that morning (TigerBoost) also began
                    # failing, on a fresh browser, on BOTH Chrome and
                    # Chromium, with no CAPTCHA. Nothing local fixes that
                    # -- only time. Retrying or switching browser just
                    # burns queries against a closed door.
                    if args.cooloff_minutes > 0 and cooloffs_left > 0:
                        cooloffs_left -= 1
                        mins = args.cooloff_minutes
                        print(
                            f"\n{consecutive_fail} consecutive backend errors -- Google is refusing "
                            f"queries from this IP.\nCooling off for {mins} min "
                            f"({cooloffs_left} cool-off(s) left after this), then resuming. "
                            "Ctrl-C is safe; progress is saved."
                        )
                        try:
                            aim.shutdown()
                        except Exception:
                            pass
                        time.sleep(mins * 60)
                        consecutive_fail = 0
                        since_restart = 0
                        continue
                    print("\nStopping: AI Mode keeps refusing queries from this IP.")
                    print("This is a cumulative daily/session limit, not a bug -- a company")
                    print("that answered fine earlier will also fail right now. Nothing was")
                    print("recorded for the failed companies, so re-running picks them up.")
                    print("Best fix: wait a few hours (or try a different network), then re-run.")
                    break
                continue

            role = parse_role(answer, prompt)
            state[str(serial)] = {
                "company": company, "website": web, "website_source": source,
                "role": role, "prompt": prompt,
                "answer": (answer or "").strip()[:2000],
            }
            save_state(state)
            print(f"[{i}/{len(todo)}] #{serial} {company[:44]:44s} {web[:28]:28s} -> {role}")

            # Distinguish "this company genuinely has no clear role" from
            # "Google has stopped answering us": only an EMPTY answer means
            # the backend failed. A parsed Unknown off a real answer is a
            # result, not a failure, so it must not trip the guard.
            consecutive_fail = 0 if (answer or "").strip() else consecutive_fail + 1
            if consecutive_fail >= 8:
                print("\nStopping: 8 consecutive empty answers -- AI Mode is blocked or the browser died.")
                print("Progress is checkpointed; re-run this script to resume.")
                break
    except KeyboardInterrupt:
        print("\nInterrupted -- progress saved, re-run to resume.")
    finally:
        try:
            aim.shutdown()
        except Exception:
            pass
        try:
            lock.unlink()
        except Exception:
            pass
        path = export(header, rows, state)
        done = sum(1 for r in rows if state.get(str(r[0]), {}).get("role"))
        unknown = sum(1 for r in rows if state.get(str(r[0]), {}).get("role") == "Unknown")
        print(f"\nAnswered {done}/{len(rows)} ({unknown} Unknown). Exported {path}")


if __name__ == "__main__":
    main()
