#!/usr/bin/env python3
"""India-based leadership (CTO / AI Head / CEO / Founder / Co-founder / COO / MD)
for a list of companies, via Google AI Mode in N parallel browsers.

Input: a workbook whose first sheet has a "Company" column (header row 1).
For global companies the India head / India leadership is wanted, never
global leaders based abroad. Up to 3 people per company, CTO and AI Head
first, then CEO, Founder, Co-founder, COO, MD.

Never edits the input. Output workbook = the input plus, on the first
sheet, Person 1-3 columns, and a "People" sheet (one row per person with
designation, location, LinkedIn and source).

Usage:
  python scripts/india_leaders_ai_mode.py --input X.xlsx --output Y.xlsx --workers 15
  python scripts/india_leaders_ai_mode.py --input X.xlsx --status
  python scripts/india_leaders_ai_mode.py --input X.xlsx --output Y.xlsx --build

Answers are checkpointed per worker in data/india_leaders/<input stem>/,
so re-running the same command resumes. STATUS_WALL_WAITS /
STATUS_WALL_WAIT_SEC control the pause-and-retry when Google blocks.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from dotenv import load_dotenv  # noqa: E402

load_dotenv(ROOT / ".env")

import openpyxl  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

from status_check_deliverables import extract_json, internet_up  # noqa: E402

CHUNK = 3
MAX_PEOPLE = 3
MAX_CONSECUTIVE_EMPTY = 3
MAX_PASSES = 4
# Bumped when the prompt asks for new fields; older answers are re-asked.
ANSWER_VERSION = 2
_EMAIL = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")

_PROMPT = (
    "For each company below, find up to 3 senior leaders who are CURRENTLY "
    "in the role and BASED IN INDIA (working in India).\n"
    "Wanted designations, in priority order: CTO (Chief Technology Officer), "
    "AI Head (Chief AI Officer / Head of AI / VP or Director of AI), CEO, "
    "Founder, Co-founder, COO, Managing Director (MD).\n"
    "Rules:\n"
    "- For a global / foreign company give its INDIA leadership (e.g. CTO "
    "India, Head of AI India, CEO India, MD India, Country Head India). Never "
    "give global leaders based outside India.\n"
    "- Fill CTO and AI Head first, then CEO, Founder, Co-founder, COO, MD.\n"
    "- Only real, named people you found evidence for; do not guess. If the "
    "company has no India-based person in these roles, return an empty list.\n"
    "- linkedin: the person's LinkedIn profile URL only if you actually found "
    "it, else \"\".\n"
    "- email: the person's work email ONLY if it is publicly listed (company "
    "site, press release, directory); never construct or guess one from a "
    "name pattern; else \"\".\n"
    "- location: the person's city in India.\n"
    "- source: where this is stated (website / news / LinkedIn), short.\n"
    "Answer every id. Reply with ONLY this JSON:\n"
    '{"results":[{"id":1,"Company":"","india_entity":"","people":[{"name":"",'
    '"designation":"","email":"","location":"","linkedin":"","source":""}]}]}'
)

_RETRY_HINT = (
    "\nA previous search found nobody for these companies, so search harder: "
    "the company's Indian subsidiary or India office (e.g. \"<Company> India "
    "Pvt Ltd\", its India site or MCA/Zauba director list), LinkedIn profiles "
    "with \"<Company>\" and an Indian city, and Indian business news. If the "
    "name is not a company at all (a sponsor tier, category or label), return "
    "an empty list.\n"
)

# Designation -> (rank, label). First match in this order wins per pattern;
# a combined title ("Founder & CEO") takes its best rank.
_ROLES = [
    (1, "CTO", r"\bc\.?t\.?o\b|chief technolog|chief technical"),
    (2, "AI Head", r"\bcaio\b|chief (data (and|&) )?(ai|artificial intelligence)"
                   r"|head\b.{0,25}\b(ai|artificial intelligence|genai|gen ai)\b"
                   r"|\b(ai|artificial intelligence|genai)\b.{0,15}\bhead"
                   r"|(vp|vice president|director)\b.{0,25}\b(ai|artificial intelligence|genai)\b"),
    (3, "CEO", r"\bc\.?e\.?o\b|chief executive"),
    (5, "Co-founder", r"co-?\s?founder"),
    (4, "Founder", r"(?<!co-)(?<!co)(?<!co )\bfounder"),
    (6, "COO", r"\bc\.?o\.?o\b|chief operating"),
    (7, "MD", r"managing director|\bm\.?d\b"),
    # Overall India leadership only ("Country Head", "President - India",
    # "Head of India"); a functional head (sales, marketing, ...) is not.
    (7, "India Head", r"country (head|manager)(?!\s+of\s+(?!india)\w)"
                      r"|^(?!.*\b(sales|marketing|services|consumer|ads|originals|hr|finance|legal"
                      r"|business development|partner|channel|engineering)\b)"
                      r"(.*head\W{1,4}(of\W+)?india\b|.*\bindia\W{0,3}(country\W)?head"
                      r"|.*president\W{1,4}(for (the )?)?(\w+\W+){0,2}india\b|.*leading india)"),
]

_ABROAD = re.compile(
    r"\b(usa|united states|u\.s\.|uk|united kingdom|london|new york|california|"
    r"singapore|dubai|uae|germany|france|japan|china|australia|canada|kenya|"
    r"nigeria|denmark|sweden|netherlands|switzerland|hong kong)\b", re.I)


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def _norm(s: object) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


def role_of(designation: str) -> tuple[int, str] | None:
    d = designation.lower()
    hits = [(rank, lab) for rank, lab, pat in _ROLES if re.search(pat, d)]
    return min(hits) if hits else None


def designation_label(designation: str) -> str:
    """Every wanted role in the title, in priority order ("Co-founder / COO")."""
    d = designation.lower()
    labels = sorted({(rank, lab) for rank, lab, pat in _ROLES if re.search(pat, d)})
    if len(labels) > 1:  # "India Head" only when no listed role applies
        labels = [x for x in labels if x[1] != "India Head"]
    return " / ".join(lab for _, lab in labels)


# -- input / checkpoints -------------------------------------------------------

def read_companies(xlsx: Path) -> list[str]:
    ws = openpyxl.load_workbook(xlsx, read_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = [str(c or "").strip() for c in next(rows)]
    ci = header.index("Company")
    out, seen = [], set()
    for r in rows:
        name = str(r[ci] or "").strip() if ci < len(r) else ""
        if name and _norm(name) not in seen:
            seen.add(_norm(name))
            out.append(name)
    return out


def run_dir(xlsx: Path) -> Path:
    d = ROOT / "data" / "india_leaders" / re.sub(r"[^A-Za-z0-9]+", "_", xlsx.stem).strip("_")
    d.mkdir(parents=True, exist_ok=True)
    return d


def load_json(p: Path) -> dict:
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_json(p: Path, data: dict) -> None:
    tmp = p.with_suffix(f".{os.getpid()}.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, p)


def current(answers: dict) -> dict:
    """Answers from the current prompt version, people re-checked against the
    current role rules (so tightening a rule fixes saved answers too)."""
    out = {}
    for c, a in answers.items():
        if a.get("v") != ANSWER_VERSION:
            continue
        people = []
        for p in a["people"]:
            role = role_of(p["designation"])
            if role:
                people.append({**p, "rank": role[0], "role": role[1]})
        people.sort(key=lambda x: x["rank"])
        out[c] = {**a, "people": people}
    return out


def all_answers(xlsx: Path) -> dict[str, dict]:
    merged: dict[str, dict] = {}
    for p in sorted(run_dir(xlsx).glob("shard_*.json")):
        merged.update(current(load_json(p)))
    # --retry answers only replace a first-pass answer that found nobody.
    for p in sorted(run_dir(xlsx).glob("retry_*.json")):
        merged.update({c: a for c, a in current(load_json(p)).items() if a["people"]})
    # Reviewer corrections: {"Company": {"people": [], "note": why}}.
    for c, fix in load_json(run_dir(xlsx) / "overrides.json").items():
        if c in merged and isinstance(fix, dict):
            merged[c] = {**merged[c], **fix}
    # Once --verify-linkedin has run, only links it confirmed are kept.
    checks = link_checks(xlsx)
    if checks:
        for a in merged.values():
            a["people"] = [p if checks.get(p["linkedin"], {}).get("status") == "verified"
                           else {**p, "linkedin": ""} for p in a["people"]]
    return merged


# -- LinkedIn link check -------------------------------------------------------

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/129.0 Safari/537.36")
_CONTROL = "https://www.linkedin.com/in/satyanadella"


def _profile_title(url: str) -> tuple[int, str]:
    import requests

    # No session: LinkedIn starts answering 999 to everything once a
    # cookie from an earlier response is sent back.
    try:
        r = requests.get(url, timeout=25, headers={"User-Agent": _UA, "Accept-Language": "en-US,en"})
    except Exception:  # noqa: BLE001 - network error = unknown, retried later
        return 0, ""
    m = re.search(r"<title>([^<]*)", r.text) if r.status_code == 200 else None
    return r.status_code, (m.group(1).strip() if m else "")


def _name_matches(person: str, title: str) -> bool:
    shown = re.split(r"\s+[-–|]\s+", title)[0]
    want = {t for t in re.findall(r"[a-z]+", person.lower()) if len(t) > 1}
    have = set(re.findall(r"[a-z]+", shown.lower()))
    # LinkedIn may shorten a surname to its initial ("Atul K.").
    initials = {t for t in have if len(t) == 1}
    hits = len(want & have) + sum(1 for t in want - have if t[0] in initials)
    return bool(want & have) and hits >= min(2, len(want))


def verify_linkedin(xlsx: Path) -> bool:
    """Open every LinkedIn link: keep it only if the public profile exists
    and its name matches the person. LinkedIn answers 999 both for a
    profile that does not exist and when rate-limiting, so a 999 counts as
    "not found" only while a known-real control profile still loads."""
    path = run_dir(xlsx) / "linkedin_check.json"
    checks = load_json(path)
    raw: dict[str, dict] = {}
    for p in sorted(run_dir(xlsx).glob("shard_*.json")) + sorted(run_dir(xlsx).glob("retry_*.json")):
        raw.update(current(load_json(p)))
    todo = {}
    for a in raw.values():
        for p in a["people"]:
            if p["linkedin"] and p["linkedin"] not in checks:
                todo[p["linkedin"]] = p["name"]
    log(f"{len(todo)} LinkedIn links to check ({len(checks)} already checked)")
    import random

    # LinkedIn blocks the IP after ~25 quick requests, so go slowly and, when
    # blocked, wait it out on the same link rather than skipping it.
    for n, (url, name) in enumerate(todo.items(), 1):
        fails = blocks = 0
        while True:
            time.sleep(random.uniform(10, 15))
            code, title = _profile_title(url)
            if code == 200 and title:
                status = "verified" if _name_matches(name, title) else "wrong_person"
                break
            if code == 404:
                status = "not_found"
                break
            time.sleep(5)
            ctrl, _ = _profile_title(_CONTROL)
            if ctrl == 200:
                # LinkedIn is answering, so the link itself failed; a
                # second failure makes it "not found".
                fails += 1
                if fails >= 2:
                    status = "not_found"
                    break
                continue
            blocks += 1
            if blocks > 6:
                log(f"STOPPING: still blocked after {blocks - 1} waits; re-run to resume "
                    f"({len(checks)} checked)")
                return False
            log(f"blocked by LinkedIn (control {ctrl}); waiting 10 min ({blocks}/6)")
            time.sleep(600)
        checks[url] = {"name": name, "status": status, "title": title}
        save_json(path, checks)
        if n % 25 == 0 or status != "verified":
            log(f"{n}/{len(todo)} {status}: {name} | {url} | {title[:70]}")
    by: dict[str, int] = {}
    for v in checks.values():
        by[v["status"]] = by.get(v["status"], 0) + 1
    log(f"LinkedIn check: {by}")
    return True


def link_checks(xlsx: Path) -> dict[str, dict]:
    """url -> check. A direct page check (linkedin_check.json) beats an
    AI Mode one (linkedin_ai_*.json)."""
    merged: dict[str, dict] = {}
    for p in sorted(run_dir(xlsx).glob("linkedin_ai_*.json")):
        merged.update(load_json(p))
    merged.update(load_json(run_dir(xlsx) / "linkedin_check.json"))
    return merged


def all_links(xlsx: Path) -> dict[str, dict]:
    """url -> the person it was given for (first-pass and retry answers)."""
    links: dict[str, dict] = {}
    for p in sorted(run_dir(xlsx).glob("shard_*.json")) + sorted(run_dir(xlsx).glob("retry_*.json")):
        for c, a in current(load_json(p)).items():
            for person in a["people"]:
                if person["linkedin"]:
                    links.setdefault(person["linkedin"], {"name": person["name"], "company": c,
                                                          "designation": person["designation"]})
    return dict(sorted(links.items()))


_LINK_PROMPT = (
    "Check whether each LinkedIn profile URL below really exists. Search for "
    "the EXACT URL / its /in/ slug - do not assume it exists because a person "
    "with that name works at that company; AI tools often invent plausible "
    "slugs.\n"
    "exists: true only if you found this exact profile URL; false if no "
    "profile with this exact slug can be found.\n"
    "profile_name: the name shown on that exact profile, else \"\".\n"
    "headline: its headline / current company, else \"\".\n"
    "Answer every id. Reply with ONLY this JSON:\n"
    '{"results":[{"id":1,"exists":false,"profile_name":"","headline":""}]}'
)
LINK_CHUNK = 10


def link_ai_worker(xlsx: Path, shard: int, shards: int) -> None:
    from vendor_intel.scraping import google_ai_mode as gam

    links = list(all_links(xlsx).items())[shard::shards]
    ck = run_dir(xlsx) / f"linkedin_ai_{shard}.json"
    done = load_json(ck)
    empty = 0
    try:
        for pass_no in range(1, 3):
            todo = [(u, v) for u, v in links if u not in done]
            if not todo:
                break
            log(f"[L{shard}] pass {pass_no}: {len(todo)} links")
            for i in range(0, len(todo), LINK_CHUNK):
                batch = todo[i:i + LINK_CHUNK]
                payload = [{"id": n + 1, "url": u, "expected_person": v["name"],
                            "company": v["company"]} for n, (u, v) in enumerate(batch)]
                parsed = extract_json(gam.ask(
                    f"{_LINK_PROMPT}\n\nProfiles:\n{json.dumps(payload, ensure_ascii=False)}",
                    retries=1) or "")
                if not parsed:
                    empty += 1
                    log(f"[L{shard}] empty answer ({empty}/{MAX_CONSECUTIVE_EMPTY})")
                    if empty >= MAX_CONSECUTIVE_EMPTY:
                        log(f"[L{shard}] STOPPING: blocked. Re-run to resume.")
                        return
                    continue
                empty = 0
                for r in parsed.get("results") or []:
                    try:
                        url, v = batch[int(r.get("id")) - 1]
                    except (TypeError, ValueError, IndexError, AttributeError):
                        continue
                    shown = str(r.get("profile_name") or "").strip()
                    exists = r.get("exists") is True or str(r.get("exists")).lower() == "true"
                    if not exists:
                        status = "not_found"
                    elif shown and _name_matches(v["name"], shown):
                        status = "verified"
                    else:
                        status = "wrong_person" if shown else "unconfirmed"
                    done[url] = {"name": v["name"], "status": status, "via": "ai_mode",
                                 "title": f"{shown} - {str(r.get('headline') or '').strip()}"}
                save_json(ck, done)
                log(f"[L{shard}] {sum(u in done for u, _ in links)}/{len(links)} links checked")
    finally:
        gam.shutdown()


def empty_companies(xlsx: Path) -> list[str]:
    merged: dict[str, dict] = {}
    for p in sorted(run_dir(xlsx).glob("shard_*.json")):
        merged.update(current(load_json(p)))
    return sorted(c for c, a in merged.items() if not a["people"])


# -- one worker ----------------------------------------------------------------

def clean_people(raw: object) -> list[dict]:
    """Keep India-based people in a wanted role, best rank first, max 3."""
    people, seen = [], set()
    for p in raw if isinstance(raw, list) else []:
        if not isinstance(p, dict):
            continue
        name = str(p.get("name") or "").strip()
        desig = str(p.get("designation") or "").strip()
        loc = str(p.get("location") or "").strip()
        role = role_of(desig)
        if not name or not role or _norm(name) in seen:
            continue
        if loc and _ABROAD.search(loc) and "india" not in loc.lower():
            continue
        li = str(p.get("linkedin") or "").strip()
        if "linkedin.com/in/" not in li.lower():
            li = ""
        email = str(p.get("email") or "").strip().lower()
        if not _EMAIL.match(email):
            email = ""
        seen.add(_norm(name))
        people.append({"name": name, "designation": desig, "role": role[1], "rank": role[0],
                       "email": email,
                       "location": loc, "linkedin": li,
                       "source": str(p.get("source") or "").strip()})
    people.sort(key=lambda x: x["rank"])
    return people[:MAX_PEOPLE]


def _parse(answer: str, batch: list[str]) -> dict[str, dict]:
    parsed = extract_json(answer or "")
    if not parsed:
        return {}
    by_name = {_norm(c): c for c in batch}
    got: dict[str, dict] = {}
    for r in parsed.get("results") or []:
        if not isinstance(r, dict):
            continue
        try:
            company = batch[int(r.get("id")) - 1]
        except (TypeError, ValueError, IndexError):
            company = by_name.get(_norm(r.get("Company")))
        if not company or company in got:
            continue
        got[company] = {"v": ANSWER_VERSION, "india_entity": str(r.get("india_entity") or "").strip(),
                        "people": clean_people(r.get("people"))}
    return got


def run_worker(xlsx: Path, shard: int, shards: int, retry: bool = False) -> None:
    from vendor_intel.scraping import google_ai_mode as gam

    if retry:
        # Second look at companies the first pass found nobody for: fewer
        # per query and a harder-search hint; asked once, saved separately.
        mine = empty_companies(xlsx)[shard::shards]
        ck = run_dir(xlsx) / f"retry_{shard}.json"
        prompt_head, chunk, passes = _PROMPT + _RETRY_HINT, 2, 1
    else:
        mine = [c for i, c in enumerate(read_companies(xlsx)) if i % shards == shard]
        ck = run_dir(xlsx) / f"shard_{shard}.json"
        prompt_head, chunk, passes = _PROMPT, CHUNK, MAX_PASSES
    answers = current(load_json(ck))
    if not retry:
        elsewhere = all_answers(xlsx)
        mine = [c for c in mine if c in answers or c not in elsewhere]
    log(f"[w{shard}] {len(mine)} companies, {sum(c in answers for c in mine)} already done")

    queries = empty = waits_used = 0
    wall_waits = int(os.getenv("STATUS_WALL_WAITS") or 3)
    wall_wait_sec = int(os.getenv("STATUS_WALL_WAIT_SEC") or 600)
    try:
        for pass_no in range(1, passes + 1):
            todo = [c for c in mine if c not in answers]
            if not todo:
                break
            log(f"[w{shard}] pass {pass_no}: {len(todo)} to research")
            for i in range(0, len(todo), chunk):
                batch = todo[i:i + chunk]
                payload = [{"id": n + 1, "Company": c} for n, c in enumerate(batch)]
                prompt = f"{prompt_head}\n\nCompanies:\n{json.dumps(payload, ensure_ascii=False)}"
                got = _parse(gam.ask(prompt, retries=1), batch)
                queries += 1
                if not got:
                    if not internet_up():
                        log(f"[w{shard}] no network, waiting 60s")
                        time.sleep(60)
                        continue
                    empty += 1
                    log(f"[w{shard}] empty answer ({empty}/{MAX_CONSECUTIVE_EMPTY})")
                    if empty >= MAX_CONSECUTIVE_EMPTY:
                        if waits_used < wall_waits:
                            waits_used += 1
                            log(f"[w{shard}] blocked — pausing {wall_wait_sec // 60} min "
                                f"({waits_used}/{wall_waits}), then fresh profile")
                            gam.shutdown()
                            time.sleep(wall_wait_sec)
                            if not gam.reset_session("blocked: pause before resuming"):
                                gam.shutdown()
                            empty = 0
                            continue
                        log(f"[w{shard}] STOPPING: {empty} empty answers in a row. Re-run to resume.")
                        return
                    continue
                empty = 0
                answers.update(got)
                save_json(ck, answers)
                n_people = sum(len(v["people"]) for v in got.values())
                done = sum(c in answers for c in mine)
                log(f"[w{shard}] q{queries}: {len(got)}/{len(batch)} answered, "
                    f"{n_people} people | {done}/{len(mine)} done")
        log(f"[w{shard}] finished: {queries} queries, {sum(c not in answers for c in mine)} unanswered")
    finally:
        gam.shutdown()


# -- build output --------------------------------------------------------------

def build(xlsx: Path, out: Path) -> None:
    answers = all_answers(xlsx)
    wb = openpyxl.load_workbook(xlsx)
    ws = wb.worksheets[0]
    header = [str(c.value or "").strip() for c in ws[1]]
    # The request column ("Designations to find/needed ...") is not wanted in the output.
    for j in reversed([j for j, h in enumerate(header)
                       if re.match(r"designations (to find|needed)", h.lower())]):
        ws.delete_cols(j + 1)
        del header[j]
    ci = header.index("Company") + 1
    start = len(header) + 1
    cols = []
    for n in range(1, MAX_PEOPLE + 1):
        cols += [f"Person {n} Name", f"Person {n} Designation", f"Person {n} Email"]
    cols += ["People found", "Research status"]
    for j, h in enumerate(cols):
        c = ws.cell(row=1, column=start + j, value=h)
        c.font = Font(bold=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(start + j)].width = 24

    pws = wb.create_sheet("People (India leaders)", 1)
    pws.append(["No.", "Company", "Name", "Designation", "Email"])
    for c in pws[1]:
        c.font = Font(bold=True)

    totals = {"with": 0, "none": 0, "pending": 0, "people": 0}
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(row=r, column=ci).value or "").strip()
        if not name:
            continue
        a = answers.get(name)
        if a is None:
            totals["pending"] += 1
            ws.cell(row=r, column=start + len(cols) - 1, value="Not researched")
            continue
        people = a["people"]
        totals["with" if people else "none"] += 1
        totals["people"] += len(people)
        for k, p in enumerate(people):
            row = [p["name"], designation_label(p["designation"]) or p["role"], p.get("email") or None]
            for j, v in enumerate(row):
                ws.cell(row=r, column=start + 3 * k + j, value=v)
            pws.append([ws.cell(row=r, column=1).value, name, *row])
        ws.cell(row=r, column=start + len(cols) - 2, value=len(people))
        ws.cell(row=r, column=start + len(cols) - 1,
                value="Found" if people else "No India-based person found")
    for col, w in zip("ABCDE", (6, 40, 30, 24, 36)):
        pws.column_dimensions[col].width = w
    pws.freeze_panes = "A2"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    log(f"companies with people {totals['with']}, none found {totals['none']}, "
        f"not researched {totals['pending']}; {totals['people']} people -> {out}")


def show_status(xlsx: Path) -> None:
    companies = read_companies(xlsx)
    answers = all_answers(xlsx)
    done = [answers[c] for c in companies if c in answers]
    found = sum(1 for a in done if a["people"])
    log(f"researched {len(done)}/{len(companies)} | with people {found} | "
        f"people {sum(len(a['people']) for a in done)} | {len(companies) - len(done)} left")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True, type=Path)
    ap.add_argument("--output", type=Path)
    ap.add_argument("--workers", type=int, default=0)
    ap.add_argument("--shard", type=int, help=argparse.SUPPRESS)
    ap.add_argument("--shards", type=int, default=1, help=argparse.SUPPRESS)
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--build", action="store_true")
    ap.add_argument("--retry", action="store_true",
                    help="re-ask only companies the first pass found nobody for")
    ap.add_argument("--verify-linkedin-ai", action="store_true",
                    help="check every LinkedIn link via AI Mode (use with --workers)")
    ap.add_argument("--verify-linkedin", action="store_true",
                    help="open every LinkedIn link; keep only real, name-matching profiles")
    args = ap.parse_args()
    xlsx = args.input.resolve()

    if args.verify_linkedin_ai:
        if args.shard is not None:
            link_ai_worker(xlsx, args.shard, args.shards)
            return
        procs = []
        for i in range(args.workers or 1):
            env = dict(os.environ, GOOGLE_AI_MODE_PROFILE_DIR=str(ROOT / "data" / f"ai_mode_status_{i}"))
            procs.append(subprocess.Popen(
                [sys.executable, __file__, "--input", str(xlsx), "--verify-linkedin-ai",
                 "--shard", str(i), "--shards", str(args.workers or 1)], env=env))
            time.sleep(15)
        for p in procs:
            p.wait()
        checks = link_checks(xlsx)
        left = sum(u not in checks for u in all_links(xlsx))
        by: dict[str, int] = {}
        for v in checks.values():
            by[v["status"]] = by.get(v["status"], 0) + 1
        log(f"LinkedIn check: {by}, {left} unchecked")
        if args.output and not left:
            build(xlsx, args.output.resolve())
        return
    if args.verify_linkedin:
        # Unchecked links are blanked by all_answers, so build only when done.
        if verify_linkedin(xlsx) and args.output:
            build(xlsx, args.output.resolve())
        return
    if args.status:
        show_status(xlsx)
        return
    if args.shard is not None:
        run_worker(xlsx, args.shard, args.shards, args.retry)
        return
    if args.workers:
        procs = []
        for i in range(args.workers):
            env = dict(os.environ, GOOGLE_AI_MODE_PROFILE_DIR=str(ROOT / "data" / f"ai_mode_status_{i}"))
            procs.append(subprocess.Popen(
                [sys.executable, __file__, "--input", str(xlsx), "--shard", str(i),
                 "--shards", str(args.workers)] + (["--retry"] if args.retry else []), env=env))
            time.sleep(15)
        for p in procs:
            p.wait()
        show_status(xlsx)
    if args.output and (args.workers or args.build):
        build(xlsx, args.output.resolve())


if __name__ == "__main__":
    main()
