#!/usr/bin/env python3
"""Classify each Manufacturer as raw-material vs finished-goods, via AI Mode.

Reads output/company_roles/Manufacturer.xlsx (the 433 companies the role
pass classified as Manufacturer) and adds a "Manufacturer Type" column:

    Raw Material    - makes ingredients/actives/excipients sold B2B to
                      other manufacturers (extracts, vitamins, probiotics
                      in bulk, API, bulk powders)
    Finished Goods  - makes the end product a consumer buys (capsules,
                      tablets, gummies, bottled drinks), including
                      contract manufacturers / OEM / ODM who produce
                      finished dosage forms for other brands
    Both            - genuinely does both at meaningful scale. NOT a
                      dumping ground: it is recorded only when the model
                      says so, because in this industry "both" is a real
                      and common business model (an extract maker that
                      also runs a capsule line), and forcing a binary
                      would misrepresent those firms.
    Unknown         - the model did not say, or did not know the company

REUSES run_company_roles.py for website resolution and page-chrome
stripping, so the prompt sees exactly the same company identity as the
role pass (Website column, else non-free email domain, else name +
industry). Keeping one implementation avoids the two scripts drifting
apart on which website a company was asked about.

NO GUESSED TYPES: same discipline as the role pass. A type is recorded
only from the model's own labelled verdict or a sentence predicated of
this company; answers that disclaim knowledge yield Unknown. A blank is
correct and re-queryable, an invented value is not.

Progress is checkpointed after every company to
data/manufacturer_type/checkpoint.json, so an interrupted run resumes.

Usage:
    python run_manufacturer_type.py --delay 12 --real-chrome
    python run_manufacturer_type.py --limit 5          # smoke test
    python run_manufacturer_type.py --reparse          # re-score, no queries
    python run_manufacturer_type.py --redo-unknown
    python run_manufacturer_type.py --export-only
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

# Reuse the role runner's helpers rather than copying them, so website
# resolution and answer cleaning cannot drift between the two passes.
_spec = importlib.util.spec_from_file_location("_roles", ROOT / "run_company_roles.py")
_roles = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_roles)

website_for = _roles.website_for
strip_page_chrome = _roles.strip_page_chrome
is_backend_failure = _roles.is_backend_failure
_NO_KNOWLEDGE = _roles._NO_KNOWLEDGE

IN_XLSX = ROOT / "output" / "company_roles" / "Manufacturer.xlsx"
OUT_XLSX = ROOT / "output" / "company_roles" / "Manufacturer_by_type.xlsx"
SPLIT_DIR = ROOT / "output" / "company_roles" / "manufacturer_type"
STATE_DIR = ROOT / "data" / "manufacturer_type"
STATE = STATE_DIR / "checkpoint.json"

_QUESTION = (
    " is this company a raw material manufacturer (it makes ingredients, "
    "extracts, actives or bulk powders that it sells to other manufacturers) "
    "or a finished goods manufacturer (it makes the final product a consumer "
    "buys, such as capsules, tablets, gummies or bottled drinks, including "
    "contract manufacturing of finished products for other brands)? First "
    "explain what it actually produces, then answer on the last line with "
    "exactly one of: Raw Material, Finished Goods, or Both"
)

PROMPT = "{name} ({site})" + _QUESTION
PROMPT_NO_SITE = "{name} it is a nutraceutical company," + _QUESTION

# What to write when the model says a company does BOTH raw material and
# finished goods. The model is still ASKED for "Both" -- forcing it to
# pick one would make it guess -- but a genuine "both" is recorded as
# plain "Manufacturer", per the user's instruction. The model's own word
# stays visible in the Type Audit sheet's answer column, so the mapping
# is reversible and nothing is lost.
BOTH_LABEL = "Manufacturer"

# Batch mode: ask about N companies in one query. Each company is NUMBERED
# and the reply must repeat the number, so an answer can be tied back to
# the company it is about. Without that, a reply listing fewer or
# reordered companies would silently shift every classification by one --
# the worst possible failure here, since every row would look plausible
# and be wrong.
BATCH_PROMPT = (
    "For each nutraceutical company listed below, decide whether it is a "
    "RAW MATERIAL manufacturer (makes ingredients, extracts, actives or "
    "bulk powders sold to other manufacturers) or a FINISHED GOODS "
    "manufacturer (makes the final product a consumer buys - capsules, "
    "tablets, gummies, bottled drinks - including contract manufacturing "
    "of finished products for other brands), or Both.\n\n"
    "Reply with one line per company, in this exact format and nothing "
    "else:\n"
    "<number>. <company name> - <Raw Material|Finished Goods|Both|Unknown>\n\n"
    "Use Unknown if you do not know the company - do not guess.\n\n"
    "{companies}"
)

# "3. Acme Ltd - Raw Material", "3) Acme - Finished Goods", and also the
# UNNUMBERED form the model actually tends to use: "Acme Ltd (acme.com) -
# Raw Material". The number is optional because AI Mode frequently drops
# it even when asked for it -- an early version required it and rejected
# 10 out of 10 correct answers.
_BATCH_LINE = re.compile(
    r"^\s*(?:(\d{1,3})\s*[.)\]]\s*)?(.+?)\s*[-–—:]\s*"
    r"(Raw\s*Material|Finished\s*Goods|Both|Unknown)\s*\.?\s*$",
    re.I | re.M,
)


def _name_key(s: str) -> str:
    """Comparable form of a company name: letters+digits, lowercased.

    Drops the legal-suffix noise ("Co., Ltd.", "Pvt", "GmbH") that the
    model adds or omits freely, so a match is decided on the distinctive
    part of the name.
    """
    s = re.sub(r"\(.*?\)", " ", s or "")  # strip the echoed website
    s = re.sub(
        r"\b(?:co|ltd|limited|inc|corp|corporation|pvt|private|gmbh|bv|nv|"
        r"sa|ag|kk|plc|llc|lp|group|company|technologies|technology)\b",
        " ", s, flags=re.I,
    )
    return re.sub(r"[^a-z0-9]", "", s.lower())


def parse_batch(answer: str, prompt: str, batch: list) -> dict[int, str]:
    """Map batch position (1-based) -> type, for lines we can trust.

    Only accepts a line whose number is in range AND whose echoed company
    name plausibly matches the company at that position. A mismatch means
    the model reordered or skipped, so that line is dropped rather than
    applied to the wrong company.
    """
    body = strip_page_chrome(answer, prompt)
    if not body or _NO_KNOWLEDGE.search(body) and len(body) < 200:
        return {}

    canon = {"rawmaterial": "Raw Material", "finishedgoods": "Finished Goods",
             "both": BOTH_LABEL, "unknown": "Unknown"}
    keys = [_name_key(r[3] or "") for r in batch]
    out: dict[int, str] = {}

    for num, name, label in _BATCH_LINE.findall(body):
        canon_key = re.sub(r"[^a-z]", "", label.lower())
        if canon_key not in canon:
            continue

        # Match on the ECHOED NAME first -- it is the reliable anchor. The
        # position number is only a tiebreaker, because the model drops or
        # renumbers it freely. Assigning a type to the wrong company is the
        # one unrecoverable error here, so an ambiguous line is skipped.
        got = _name_key(name)
        idx = None
        if got:
            matches = [
                n for n, k in enumerate(keys, 1)
                if k and (k == got or (len(k) >= 5 and len(got) >= 5 and (k in got or got in k)))
            ]
            if len(matches) == 1:
                idx = matches[0]
            elif len(matches) > 1 and num and int(num) in matches:
                idx = int(num)  # several similar names; trust the number
        if idx is None and num and 1 <= int(num) <= len(batch) and not got:
            idx = int(num)
        if idx is None:
            continue
        out[idx] = canon[canon_key]
    return out

# Ordered longest-first so "Raw Material" is tested before bare "Raw".
TYPE_WORDS = [
    (BOTH_LABEL, r"\bboth\b"),
    ("Raw Material", r"\braw[\s-]?materials?\b|\bingredients?\s+manufacturer\b|\bbulk\b|\bAPI\b|\bactives?\b|\bexcipients?\b"),
    ("Finished Goods", r"\bfinished[\s-]?goods?\b|\bfinished[\s-]?products?\b|\bfinished[\s-]?dosage\b|\bconsumer[\s-]?products?\b"),
]

# The model's own verdict line, e.g. "Answer: Raw Material" or a bare
# "Finished Goods" on its own line. Mirrors the role pass's _ROLE_LABEL,
# with the same guard against matching a section HEADING.
_TYPE_LABEL = re.compile(
    # Start of a line OR just after a sentence end, because the verdict is
    # usually on its own line but not always ("...capsules. Answer: Finished
    # Goods" appears inline too).
    r"(?:^|\n|(?<=[.!?])\s)[^\w\n]*"
    r"(?:Answer|Type|Classification|Verdict|Category)\b"
    r"(?!\s+(?:of|in|as|for|and)\b)\s*[:\-]?\s*\n?\s*"
    r"([A-Za-z][A-Za-z /&-]{2,40})",
    re.I,
)

# The model saying NEITHER option applies -- the company does not
# manufacture at all (a media firm, a trade-show organiser, a distributor
# that was mis-bucketed upstream). Any type word in such an answer sits
# inside a negation, so nothing may be recorded.
_NEITHER = re.compile(
    r"(?:^|\n)\s*Neither\b"
    r"|\bneither\b[^.\n]{0,60}\bnor\b"
    r"|\bdoes not manufacture\b[^.\n]{0,80}\b(?:or|nor)\b[^.\n]{0,40}"
    r"\b(?:finished|consumer|raw)\b"
    r"|\bis not (?:a |an )?[\w\s,\-]{0,40}?"
    r"(?:manufacturer|manufacturing (?:entity|company|firm|business))\b"
    r"|\brather than a (?:physical |actual )?[\w\s]{0,20}?manufacturer\b",
    re.I,
)

# A clause predicated of THIS company, not of its customers or its market.
_IS_ABOUT_SUBJECT = re.compile(
    r"\b(?:is|acts? as|operates? as|functions? as|serves? as|produces?|"
    r"manufactures?|makes?|supplies|specialis|specializ|primarily|mainly|"
    r"the company|this company|they|its role)\b",
    re.I,
)
_GENERIC_SUBJECT = re.compile(
    r"\b(?:companies|firms|manufacturers|distributors|suppliers|brands|"
    r"products|ingredients|the (?:industry|sector|market))\b\s+(?:are|is|can|"
    r"typically|generally|usually|often|worldwide|globally)",
    re.I,
)


def parse_type(answer: str, prompt: str = "") -> str:
    """Read the model's raw-material / finished-goods verdict, or Unknown."""
    body = strip_page_chrome(answer, prompt)
    if not body:
        return "Unknown"
    if _NO_KNOWLEDGE.search(body):
        return "Unknown"
    # The model rejecting both options -- "Neither", or "it does not
    # manufacture raw materials or finished consumer goods". Verified
    # live: a media/publishing company answered exactly that, and matching
    # "raw materials" inside the NEGATION recorded it as Raw Material.
    if _NEITHER.search(body):
        return "Unknown"

    # 1. The labelled verdict the prompt asked for.
    m = _TYPE_LABEL.search(body)
    if m:
        for label, pattern in TYPE_WORDS:
            if re.search(pattern, m.group(1), re.I):
                return label

    # 2. A bare verdict line ("Finished Goods." on its own).
    for line in body.splitlines():
        stripped = line.strip().strip(".*_# ")
        if 3 <= len(stripped) <= 30:
            for label, pattern in TYPE_WORDS:
                if re.fullmatch(pattern, stripped, re.I):
                    return label

    # 3. A line that OPENS with the verdict and then qualifies it.
    for line in body.splitlines():
        opener = line.strip().strip("*_# ")
        for label, pattern in TYPE_WORDS:
            if re.match(rf"(?:{pattern})\s*(?:[(.,:\-]|$)", opener, re.I):
                return label

    # 4. Last resort: a sentence about THIS company. Take the last one, and
    #    skip sentences describing the industry or the customers.
    best, best_pos = None, -1
    for sentence in re.split(r"(?<=[.!?])\s+|\n", body):
        if _GENERIC_SUBJECT.search(sentence) or not _IS_ABOUT_SUBJECT.search(sentence):
            continue
        for label, pattern in TYPE_WORDS:
            hits = list(re.finditer(pattern, sentence, re.I))
            if hits and body.find(sentence) + hits[-1].start() > best_pos:
                best, best_pos = label, body.find(sentence) + hits[-1].start()
    return best or "Unknown"


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


def export(header, rows, state):
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Manufacturers"
    ws.append(list(header) + ["Manufacturer Type"])
    audit = wb.create_sheet("Type Audit")
    audit.append([
        "serial number", "Company", "Website Used", "Website Source",
        "Manufacturer Type", "Check", "Prompt Sent", "AI Mode Answer",
    ])

    for r in rows:
        rec = state.get(str(r[0])) or {}
        ws.append(list(r) + [rec.get("type", "")])
        note = ""
        if rec:
            if rec.get("type") == "Unknown":
                note = "no type stated"
            elif rec.get("website_source") == "none":
                note = "no website - matched by name only"
        audit.append([
            r[0], r[3], rec.get("website", ""), rec.get("website_source", ""),
            rec.get("type", ""), note, rec.get("prompt", ""),
            (rec.get("answer") or "")[:1500],
        ])

    for col, w in zip("ABCDEFGHIJK", [8, 22, 26, 40, 20, 20, 30, 28, 55, 14, 18]):
        ws.column_dimensions[col].width = w
    for col, w in zip("ABCDEFGH", [8, 40, 28, 16, 18, 26, 60, 110]):
        audit.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_XLSX)

    # Also split into one workbook per type, matching split_roles_excel.py.
    SPLIT_DIR.mkdir(parents=True, exist_ok=True)
    buckets: dict[str, list] = {}
    for r in rows:
        rec = state.get(str(r[0])) or {}
        buckets.setdefault(rec.get("type") or "Unknown", []).append(r)
    for label, items in buckets.items():
        out = openpyxl.Workbook()
        sh = out.active
        sh.title = re.sub(r"[^A-Za-z0-9 _-]", "", label)[:31] or "Sheet1"
        sh.append(list(header) + ["Manufacturer Type"])
        for i, r in enumerate(items, 1):
            sh.append([i] + list(r[1:]) + [label])
        for col, w in zip("ABCDEFGHIJK", [8, 22, 26, 40, 20, 20, 30, 28, 55, 14, 18]):
            sh.column_dimensions[col].width = w
        sh.freeze_panes = "A2"
        out.save(SPLIT_DIR / f"{re.sub(r'[^A-Za-z0-9 _-]', '', label) or 'Unknown'}.xlsx")

    return OUT_XLSX


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--export-only", action="store_true")
    ap.add_argument("--reparse", action="store_true")
    ap.add_argument("--redo-unknown", action="store_true")
    ap.add_argument("--real-chrome", action="store_true",
                    help="Use real Chrome (recommended; far fewer CAPTCHAs, no solver)")
    ap.add_argument("--delay", type=float, default=12.0)
    ap.add_argument("--cooloff-minutes", type=int, default=30)
    ap.add_argument("--max-cooloffs", type=int, default=4)
    ap.add_argument("--recycle-every", type=int, default=25)
    ap.add_argument("--batch", type=int, default=5,
                    help="Companies per query (1 = one at a time). Batching saves "
                         "queries, which matters because AI Mode has a cumulative "
                         "per-IP daily limit. Smaller batches give the model fewer "
                         "companies to confuse and so defer fewer to the slower "
                         "one-at-a-time pass. Anything a batch cannot resolve "
                         "confidently falls back to an individual query.")
    args = ap.parse_args()

    apply_env_overrides()
    os.environ["GOOGLE_AI_MODE_QUERY_DELAY"] = str(max(0.0, args.delay))
    # Must be set BEFORE the scraper launches a browser -- it decides the
    # build (and therefore which profile directory is used) at launch.
    # Verified: without this, --real-chrome still created the *_chromium
    # profile, because .env requests the extension and that forces Chromium.
    os.environ["GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM"] = (
        "false" if args.real_chrome else "true"
    )
    # GOOGLE_AI_MODE_BROWSER in .env currently says "chromium", which picks
    # the build (and profile dir) independently of the extension switch
    # above. Without setting it here, --real-chrome silently still ran
    # Chromium -- verified by the extension loading, which only Chromium
    # does. Set it explicitly so the flag means what it says.
    if args.real_chrome:
        os.environ["GOOGLE_AI_MODE_BROWSER"] = "chrome"

    header, rows = read_rows()
    state = load_state()

    if args.reparse:
        # Batched rows must be re-scored with parse_batch: their stored
        # answer is the WHOLE batch reply, so the single-company parser
        # would read another company's verdict. Re-running the batch
        # parser needs the original batch, which is not stored, so those
        # rows are left as they are -- use --redo-unknown to re-query them
        # individually if a batch verdict looks wrong.
        changed = skipped = 0
        for rec in state.values():
            if rec.get("batched"):
                skipped += 1
                continue
            new = parse_type(rec.get("answer") or "", rec.get("prompt") or "")
            if new != rec.get("type"):
                rec["type"] = new
                changed += 1
        save_state(state)
        print(f"({skipped} batch-classified rows left untouched)")
        print(f"Re-parsed {len(state)} answers, {changed} changed. Exported {export(header, rows, state)}")
        return

    if args.export_only:
        print(f"Exported {export(header, rows, state)}")
        return

    lock = STATE_DIR / "run.lock"
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    if lock.exists() and time.time() - lock.stat().st_mtime < 900:
        print(f"Another run appears active. Delete {lock} if stale.")
        sys.exit(1)
    lock.write_text(str(os.getpid()), encoding="utf-8")

    from vendor_intel.scraping import google_ai_mode as aim

    todo = []
    for r in rows:
        rec = state.get(str(r[0]))
        if rec and rec.get("type") and not (args.redo_unknown and rec["type"] == "Unknown"):
            continue
        todo.append(r)
    if args.limit:
        todo = todo[: args.limit]

    print(f"{len(rows)} manufacturers, {len(rows) - len(todo)} already typed, {len(todo)} to query.")
    consecutive_fail = 0
    cooloffs_left = args.max_cooloffs
    since_restart = 0

    # ---- batch mode -------------------------------------------------
    # Ask about --batch companies per query. Anything the batch does not
    # confidently resolve (missing line, name mismatch, Unknown) falls
    # through to the per-company path below, so batching only ever saves
    # queries -- it never costs accuracy.
    if args.batch > 1 and todo:
        remaining = []
        batches = [todo[i:i + args.batch] for i in range(0, len(todo), args.batch)]
        for bi, batch in enumerate(batches, 1):
            if since_restart >= args.recycle_every > 0:
                try:
                    aim.shutdown()
                except Exception:
                    pass
                time.sleep(5)
                since_restart = 0
            since_restart += 1

            listing = "\n".join(
                f"{n}. {(r[3] or '').strip()}"
                + (f" ({website_for(r[7], r[6])[0]})" if website_for(r[7], r[6])[0] else "")
                for n, r in enumerate(batch, 1)
            )
            prompt = BATCH_PROMPT.format(companies=listing)
            answer = aim.ask(prompt)
            if is_backend_failure(answer, prompt):
                time.sleep(10)
                answer = aim.ask(prompt)

            got = {} if is_backend_failure(answer, prompt) else parse_batch(answer, prompt, batch)
            saved = 0
            for n, r in enumerate(batch, 1):
                mtype = got.get(n)
                if not mtype or mtype == "Unknown":
                    remaining.append(r)  # resolve individually below
                    continue
                web, source = website_for(r[7], r[6])
                state[str(r[0])] = {
                    "company": (r[3] or "").strip(), "website": web,
                    "website_source": source, "type": mtype,
                    "prompt": prompt, "answer": (answer or "").strip()[:2000],
                    "batched": True,
                }
                saved += 1
            save_state(state)
            print(f"[batch {bi}/{len(batches)}] {len(batch)} companies -> {saved} typed, "
                  f"{len(batch) - saved} deferred to individual queries")

        print(f"\nBatch pass done. {len(remaining)} companies need an individual query.")
        todo = remaining

    try:
        for i, r in enumerate(todo, 1):
            serial, company = r[0], (r[3] or "").strip()
            web, source = website_for(r[7], r[6])
            prompt = PROMPT.format(name=company, site=web) if web else PROMPT_NO_SITE.format(name=company)

            if since_restart >= args.recycle_every > 0:
                try:
                    aim.shutdown()
                except Exception:
                    pass
                time.sleep(5)
                since_restart = 0
            since_restart += 1

            answer = aim.ask(prompt)
            if is_backend_failure(answer, prompt):
                print(f"[{i}/{len(todo)}] #{serial} {company[:42]:42s} backend error -- retrying in 10s")
                time.sleep(10)
                answer = aim.ask(prompt)

            if is_backend_failure(answer, prompt):
                consecutive_fail += 1
                print(f"[{i}/{len(todo)}] #{serial} {company[:42]:42s} -> BACKEND ERROR (not saved)")
                if consecutive_fail >= 5:
                    if args.cooloff_minutes > 0 and cooloffs_left > 0:
                        cooloffs_left -= 1
                        print(f"\nPer-IP limit hit. Cooling off {args.cooloff_minutes} min "
                              f"({cooloffs_left} left), then resuming. Ctrl-C is safe.")
                        try:
                            aim.shutdown()
                        except Exception:
                            pass
                        time.sleep(args.cooloff_minutes * 60)
                        consecutive_fail = 0
                        since_restart = 0
                        continue
                    print("\nStopping: AI Mode keeps refusing queries from this IP.")
                    print("Wait a few hours (or change network) and re-run to resume.")
                    break
                continue

            consecutive_fail = 0
            mtype = parse_type(answer, prompt)
            state[str(serial)] = {
                "company": company, "website": web, "website_source": source,
                "type": mtype, "prompt": prompt,
                "answer": (answer or "").strip()[:2000],
            }
            save_state(state)
            print(f"[{i}/{len(todo)}] #{serial} {company[:42]:42s} {web[:26]:26s} -> {mtype}")
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
        from collections import Counter

        done = Counter(rec["type"] for rec in state.values() if rec.get("type"))
        print(f"\nTyped {sum(done.values())}/{len(rows)}: {dict(done)}")
        print(f"Exported {export(header, rows, state)}")


if __name__ == "__main__":
    main()
