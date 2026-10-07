"""Standalone leadership-contact lookup pipeline.

Two-step design per company:

1. Ask the Google AI Overview scraper a plain, natural-language question
   (no JSON constraint) — the exact phrasing the user found gets the
   richest answers, e.g. "give me full name and email of md,ceo,country
   head,cto,ai head,marketing head, founder of {company}". Measured: a
   rigid "reply ONLY as JSON" instruction suppressed the model's normal
   broad web search/synthesis and returned far fewer names than the plain
   question does for the exact same company.
2. Feed that free-text answer to this repo's shared LLM router
   (placeholders/llm.py's llm_complete_json, same provider configured via
   LLM_PROVIDER in .env) to extract the 7 roles into structured
   {name, email} pairs for the Excel export.

Deliberately separate from every other pipeline in this repo: it does not
import or modify company_enrichment.py, kompass_ai_search.py,
website_verification.py, or any other existing module beyond re-using the
generic ask_generative()/llm_complete_json() entry points they already
share.
"""
from __future__ import annotations

import asyncio
import json
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

QUERY_TEMPLATE = (
    "give me full name and email of md,ceo,country head,cto,ai head,"
    "marketing head, founder of {company}"
)

ROLES = ["MD", "CEO", "COUNTRY HEAD", "CTO", "AI HEAD", "MARKETING HEAD", "Founder"]

# Lower-case wording used inside a query, matching the phrasing of
# QUERY_TEMPLATE (which is what the AI Overview answers best).
_ROLE_QUERY_WORDS = {
    "MD": "md",
    "CEO": "ceo",
    "COUNTRY HEAD": "country head",
    "CTO": "cto",
    "AI HEAD": "ai head",
    "MARKETING HEAD": "marketing head",
    "Founder": "founder",
}


def build_gap_query(company: str, missing_roles: list[str]) -> str:
    """Query asking only about the roles a company is still missing.

    Asking for all seven roles when six are already known wastes the
    answer's budget re-stating what we have, and (measured) the AI Overview
    often replies with a generic overview instead of naming the one person
    actually wanted. A narrower question gets a more direct answer.
    """
    if not missing_roles:
        return ""
    words = [_ROLE_QUERY_WORDS[r] for r in missing_roles if r in _ROLE_QUERY_WORDS]
    if not words:
        return ""
    if len(words) == 1:
        return f"give me full name and email of {words[0]} of {company}"
    return f"give me full name and email of {','.join(words)} of {company}"

_EXTRACT_SYSTEM_PROMPT = (
    "You extract structured contact data from a free-text answer about a "
    "company's leadership. Reply with ONLY a JSON object, no markdown, no "
    "commentary."
)

_EXTRACT_USER_TEMPLATE = """Company: {company}

Free-text research answer about this company's leadership (may include
citations, notes, and prose — extract only the factual name/email pairs):
---
{answer}
---

Extract the full name for each of these roles: MD, CEO, COUNTRY HEAD, CTO,
AI HEAD, MARKETING HEAD, Founder. The NAME is the important field — always
fill it in when the text identifies that role, whether or not an email is
present. Only set email when a real, specific address for that person
appears in the text; leave email as an empty string otherwise (ignore
generic company addresses like info@/support@/customercare@ and template
patterns like "firstname.lastname@company.com" — those belong to no
individual).

Match roles generously to the titles in the text, e.g. "Managing
Director" -> MD, "Chief Information Officer"/"Tech Head" -> CTO, "Head of
Digital"/"Head of AI" -> AI HEAD, "CMO"/"Head of Marketing" -> MARKETING
HEAD, "Deputy CEO"/"Regional/India Head" -> COUNTRY HEAD. One person may
fill several roles; repeat the name in each. Leave a role's fields empty
only when the text truly does not identify anyone for it. Never invent a
name or email that isn't in the text above.

For Founder, record whoever founded the company even when that is an
organisation rather than a person — a joint venture "founded by State Bank
of India and Insurance Australia Group" should set Founder to those
parent companies. If the text says the company has no individual founder
but names the founding parent(s), use the parent name(s); only leave
Founder empty when the text identifies no founding person or entity at
all.

Reply with exactly this JSON shape:
{{"MD": {{"name": "", "email": ""}}, "CEO": {{"name": "", "email": ""}},
"COUNTRY HEAD": {{"name": "", "email": ""}}, "CTO": {{"name": "", "email":
""}}, "AI HEAD": {{"name": "", "email": ""}}, "MARKETING HEAD": {{"name":
"", "email": ""}}, "Founder": {{"name": "", "email": ""}}}}"""


@dataclass
class LeadershipRow:
    company: str
    values: dict[str, tuple[str, str]] = field(default_factory=dict)
    raw_reply: str = ""
    error: str = ""
    # How many times the gap-fill pass has queried this company. Persisted
    # so a resumed run can skip companies already retried, instead of
    # re-asking about roles that are genuinely absent (most companies have
    # no named "AI Head" or "Country Head", so those gaps never close and
    # would otherwise be re-queried on every run).
    gapfill_attempts: int = 0


def _balanced_json_spans(text: str) -> list[str]:
    """Return every top-level {...} substring, tracking brace depth and
    string state so braces inside quoted values don't confuse the count."""
    spans = []
    depth = 0
    start = -1
    in_string = False
    escape = False
    for i, ch in enumerate(text):
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and start >= 0:
                    spans.append(text[start : i + 1])
    return spans


def _extract_json(text: str) -> dict:
    """Parse the answer's JSON object out of the raw scrape.

    The reply echoes our own prompt back verbatim before the actual answer
    (e.g. "AI Mode reply for <prompt>..."), and our prompt itself contains
    an example JSON shape as an instruction to the model. A naive
    first-"{"-to-last-"}" slice grabs that echoed example instead of the
    real answer and fails to parse ("Extra data"). Instead, collect every
    balanced {...} span and return the LAST one that parses AND contains at
    least one of our expected role keys, since the real answer always comes
    after the echoed prompt.
    """
    text = (text or "").strip()
    if text.startswith("```"):
        text = text.strip("`").removeprefix("json").strip()
    for span in reversed(_balanced_json_spans(text)):
        try:
            # strict=False: the scraper renders each answer token on its own
            # line, so raw newlines end up inside quoted string values
            # (e.g. "\nAnil Chakravarthy\n") — technically invalid JSON
            # (unescaped control chars), which the default strict parser
            # rejects outright, silently discarding a real, correct answer.
            data = json.loads(span, strict=False)
        except (json.JSONDecodeError, ValueError):
            continue
        if isinstance(data, dict) and any(role in data for role in ROLES):
            return data
    return {}


def _clean_field(value: str) -> str:
    """Collapse the scraper's one-token-per-line rendering into normal
    spacing, e.g. "\\nJohn Warnock\\nand\\nCharles Geschke\\n" ->
    "John Warnock and Charles Geschke"."""
    return " ".join(value.split())


def build_query(company: str) -> str:
    return QUERY_TEMPLATE.format(company=company)


def _row_to_dict(row: LeadershipRow) -> dict:
    return {
        "company": row.company,
        "values": {role: list(pair) for role, pair in row.values.items()},
        "raw_reply": row.raw_reply,
        "error": row.error,
        "gapfill_attempts": row.gapfill_attempts,
    }


def _row_from_dict(data: dict) -> LeadershipRow:
    row = LeadershipRow(company=data.get("company", ""))
    row.values = {role: tuple(pair) for role, pair in (data.get("values") or {}).items()}
    row.raw_reply = data.get("raw_reply", "")
    row.error = data.get("error", "")
    row.gapfill_attempts = int(data.get("gapfill_attempts") or 0)
    return row


class Checkpoint:
    """Persists completed lookups to a JSON file, keyed by company name, so
    a killed or interrupted run can resume without re-querying companies
    it already finished. Written incrementally (one flush per completed
    company) rather than only at the end, so a crash mid-run loses at most
    the in-flight batch."""

    def __init__(self, path: str) -> None:
        self.path = Path(path)
        self._lock = threading.Lock()
        self._done: dict[str, dict] = {}
        if self.path.exists():
            try:
                self._done = json.loads(self.path.read_text(encoding="utf-8")) or {}
            except (json.JSONDecodeError, OSError):
                self._done = {}

    def is_done(self, company: str) -> bool:
        return company.lower() in self._done

    def get(self, company: str) -> "LeadershipRow | None":
        data = self._done.get(company.lower())
        return _row_from_dict(data) if data else None

    def all_rows(self) -> list[LeadershipRow]:
        return [_row_from_dict(d) for d in self._done.values()]

    def record(self, row: LeadershipRow) -> None:
        with self._lock:
            self._done[row.company.lower()] = _row_to_dict(row)
            self.path.parent.mkdir(parents=True, exist_ok=True)
            payload = json.dumps(self._done, indent=2)
            # Temp name derived from the checkpoint's own name: a bare
            # ".tmp" suffix is shared by every checkpoint in the folder, so
            # two runs (e.g. gap-fill and a fresh AI Mode pass) would
            # overwrite each other's staging file.
            tmp = self.path.with_name(self.path.name + ".tmp")
            tmp.write_text(payload, encoding="utf-8")

            # os.replace intermittently fails with WinError 5 on Windows
            # when an indexer or antivirus has the destination open for a
            # moment. Observed live mid-run: the whole pass died with
            # PermissionError after the batch had already been fetched.
            # Retry briefly, then fall back to a direct write — losing the
            # atomicity is much better than losing the run.
            last_error: Exception | None = None
            for attempt in range(5):
                try:
                    tmp.replace(self.path)
                    return
                except PermissionError as exc:  # noqa: PERF203
                    last_error = exc
                    time.sleep(0.2 * (attempt + 1))
            try:
                self.path.write_text(payload, encoding="utf-8")
                tmp.unlink(missing_ok=True)
            except OSError:
                raise last_error if last_error else RuntimeError("checkpoint write failed")


def lookup_company(company: str) -> LeadershipRow:
    """Sync wrapper around _lookup_company_async — shares its parsing so the
    sequential path also reads markdown_raw instead of the (LLM-cleaned)
    markdown field, which corrupts our strict JSON payload."""
    return asyncio.run(_lookup_company_async(company))


def run_lookup(
    companies: list[str],
    *,
    delay_seconds: float = 1.0,
    progress=print,
    checkpoint: "Checkpoint | None" = None,
) -> list[LeadershipRow]:
    results: list[LeadershipRow] = []
    total = len(companies)
    for i, company in enumerate(companies, start=1):
        if checkpoint is not None and checkpoint.is_done(company):
            progress(f"[{i}/{total}] {company}  -> skipped (already checkpointed)")
            results.append(checkpoint.get(company))
            continue
        progress(f"[{i}/{total}] {company}")
        row = lookup_company(company)
        if row.error:
            progress(f"  -> error: {row.error}")
        results.append(row)
        if checkpoint is not None:
            checkpoint.record(row)
        if i < total and delay_seconds > 0:
            time.sleep(delay_seconds)
    return results


def _extract_structured(company: str, answer: str) -> dict:
    """Second step: turn the free-text AI Overview answer into structured
    {role: {name, email}} JSON via this repo's shared LLM router (whatever
    LLM_PROVIDER is configured in .env — OpenAI/Anthropic/Gemini/Groq/
    OpenCode). Returns {} on any failure so the caller can fall back to the
    regex-based _extract_json instead of losing the row entirely."""
    import os

    from vendor_intel.placeholders.llm import LLM_PROVIDER, llm_complete_json

    # llm.py's _resolve_dispatch_model hardcodes "gpt-4o-mini" for
    # provider=="openai" when no model is passed, ignoring OPENAI_MODEL —
    # this repo's .env actually points OPENAI_BASE_URL at DeepSeek's API
    # (OPENAI_MODEL=deepseek-chat), so the hardcoded default gets rejected
    # with "unsupported model" and every extraction call fails. Pass the
    # configured model explicitly to sidestep that.
    model = os.getenv("OPENAI_MODEL") if LLM_PROVIDER == "openai" else None

    try:
        data = llm_complete_json(
            _EXTRACT_SYSTEM_PROMPT,
            _EXTRACT_USER_TEMPLATE.format(company=company, answer=answer[:6000]),
            model=model,
        )
    except Exception:  # noqa: BLE001
        return {}
    if not isinstance(data, dict) or data.get("status") == "llm_failed":
        return {}
    return _drop_unsupported(data, answer)


def _drop_unsupported(parsed: dict, source: str) -> dict:
    """Delete any name or email the source answer does not actually contain.

    The extraction model occasionally invents a person. Observed live:
    Beatoven.ai's answer named only "Mansoor Rahimat Khan", yet the
    extractor returned "Rukam Sitara" as CEO/Country Head/Founder —
    apparently confabulated from the LinkedIn slug
    "linkedin.com/in/mansoorkhansitar" and/or Rukam Capital, an investor in
    the company. A fabricated name is worse than a blank cell because
    nothing downstream can tell it from researched data, so require every
    exported value to be grounded in the text it came from.
    """
    import re

    source_tokens = set(re.findall(r"[a-z0-9]{2,}", (source or "").lower()))
    if not source_tokens:
        return parsed

    cleaned: dict = {}
    for role, entry in parsed.items():
        if not isinstance(entry, dict):
            cleaned[role] = entry
            continue
        name = str(entry.get("name") or "").strip()
        email = str(entry.get("email") or "").strip()

        # A name is supported when every substantial word in it appears in
        # the source. Initials and one-letter tokens are ignored, since the
        # answer may write "Dr. V S S Kiran" and the extractor "V. S. S.
        # Kiran".
        if name:
            words = [w for w in re.findall(r"[a-z0-9]{2,}", name.lower())]
            if words and any(w not in source_tokens for w in words):
                name = ""
        if email and email.lower() not in (source or "").lower():
            email = ""
        cleaned[role] = {"name": name, "email": email}
    return cleaned


_MAX_ATTEMPTS = 3  # empirically: the same query sometimes gets a truncated
# AI Overview answer (just the "emails are private" disclaimer, missing the
# actual name list) and sometimes gets the full list — non-deterministic
# rendering on Google's side, not something fixable in this client. Retrying
# recovers most of these.


async def _lookup_once(company: str, roles: list[str] | None = None) -> LeadershipRow:
    """One query + extraction. `roles` narrows the question to just those
    roles (used by the gap-fill pass); None asks about all seven."""
    from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_enabled

    row = LeadershipRow(company=company)
    if not google_ai_enabled():
        row.error = "generative web backend not enabled (GOOGLE_AI_SCRAPER_ENABLED)"
        return row

    query = build_gap_query(company, roles) if roles else build_query(company)
    if not query:
        return row

    try:
        data = await google_ai_ask(query)
    except Exception as exc:  # noqa: BLE001
        row.error = f"{type(exc).__name__}: {exc}"
        return row

    if data.get("captcha_blocked"):
        row.error = "captcha_blocked"
    # This step's query is now plain natural language (not "reply ONLY as
    # JSON"), so the cleaned `markdown` is the right field — it's the
    # DeepSeek pass's intended use case (strip nav chrome from a prose
    # answer), unlike the old rigid-JSON query where cleaning corrupted the
    # payload.
    reply = str(data.get("markdown") or data.get("markdown_raw") or "").strip()
    row.raw_reply = reply

    parsed = _extract_structured(company, reply) if reply else {}
    if not parsed:
        # Fall back to the regex-based extractor in case the free-text
        # answer already happens to contain a JSON block (e.g. a stray
        # scraper artifact), rather than silently returning an empty row.
        parsed = _extract_json(reply)

    for role in ROLES:
        entry = parsed.get(role) or {}
        if isinstance(entry, dict):
            name = _clean_field(str(entry.get("name") or ""))
            email = _clean_field(str(entry.get("email") or ""))
        else:
            name, email = "", ""
        row.values[role] = (name, email)
    return row


def _names_found(row: LeadershipRow) -> int:
    return sum(1 for name, _ in row.values.values() if name)


async def _lookup_company_async(
    company: str, roles: list[str] | None = None
) -> LeadershipRow:
    best: "LeadershipRow | None" = None
    for attempt in range(_MAX_ATTEMPTS):
        row = await _lookup_once(company, roles)
        if row.error and row.error != "captcha_blocked":
            return row  # a hard error (backend disabled, exception) — retrying won't help
        if best is None or _names_found(row) > _names_found(best):
            best = row
        if _names_found(row) > 0 or attempt == _MAX_ATTEMPTS - 1:
            break
    return best if best is not None else LeadershipRow(company=company)


def missing_roles(row: LeadershipRow) -> list[str]:
    return [r for r in ROLES if not (row.values.get(r) or ("", ""))[0].strip()]


def merge_row(base: LeadershipRow, found: LeadershipRow) -> LeadershipRow:
    """Fill base's empty roles from `found`, never overwriting a known name.

    The gap query only asks about missing roles, but the AI Overview may
    still volunteer others; keeping the original values means a second pass
    can only ever add information.
    """
    merged = LeadershipRow(company=base.company)
    merged.values = dict(base.values)
    merged.raw_reply = found.raw_reply or base.raw_reply
    merged.error = found.error
    merged.gapfill_attempts = base.gapfill_attempts + 1
    for role in ROLES:
        have_name, have_email = merged.values.get(role) or ("", "")
        new_name, new_email = found.values.get(role) or ("", "")
        merged.values[role] = (
            have_name or new_name,
            have_email or new_email,
        )
    return merged


async def run_gap_fill(
    rows: list[LeadershipRow],
    *,
    concurrency: int = 3,
    progress=print,
    checkpoint: "Checkpoint | None" = None,
) -> list[LeadershipRow]:
    """Second pass: for each company, ask only about the roles still blank.

    Returns the merged rows in the same order. Companies with nothing
    missing are passed through untouched and cost no query.
    """
    targets = [(i, r, missing_roles(r)) for i, r in enumerate(rows)]
    todo = [(i, r, miss) for i, r, miss in targets if miss]
    results: list[LeadershipRow] = list(rows)

    total = len(todo)
    if not total:
        progress("nothing missing — no gap-fill queries needed")
        return results

    progress(f"gap-fill: {total} of {len(rows)} companies have missing roles")
    semaphore = asyncio.Semaphore(concurrency)
    done = 0

    async def worker(index: int, row: LeadershipRow, miss: list[str]) -> None:
        nonlocal done
        async with semaphore:
            found = await _lookup_company_async(row.company, miss)
            merged = merge_row(row, found)
            done += 1
            gained = _names_found(merged) - _names_found(row)
            progress(
                f"[{done}/{total}] {row.company}  missing={len(miss)}  +{gained} new"
                + (f"  -> {merged.error}" if merged.error else "")
            )
            results[index] = merged
            if checkpoint is not None:
                checkpoint.record(merged)

    await asyncio.gather(*(worker(i, r, m) for i, r, m in todo))
    return results


async def run_lookup_concurrent(
    companies: list[str],
    *,
    concurrency: int = 5,
    progress=print,
    checkpoint: "Checkpoint | None" = None,
) -> list[LeadershipRow]:
    """Run lookups with up to `concurrency` companies in flight at once —
    each one opens its own browser tab against the google-ai-scraper
    backend, which is explicitly designed for ~10 concurrent tabs and
    coordinates CAPTCHA back-off across all of them (see
    clients/google_ai_scraper.py's _CaptchaGate)."""
    total = len(companies)
    done = 0
    semaphore = asyncio.Semaphore(concurrency)
    results: list[LeadershipRow | None] = [None] * total

    async def worker(index: int, company: str) -> None:
        nonlocal done
        if checkpoint is not None and checkpoint.is_done(company):
            done += 1
            progress(f"[{done}/{total}] {company}  -> skipped (already checkpointed)")
            results[index] = checkpoint.get(company)
            return
        async with semaphore:
            row = await _lookup_company_async(company)
            done += 1
            progress(f"[{done}/{total}] {company}" + (f"  -> error: {row.error}" if row.error else ""))
            if checkpoint is not None:
                checkpoint.record(row)
            results[index] = row

    await asyncio.gather(*(worker(i, c) for i, c in enumerate(companies)))
    return [r for r in results if r is not None]


def write_excel(rows: list[LeadershipRow], output_path: str, *, names_only: bool = False) -> str:
    """Write the results as a ready-to-read workbook.

    Formatting matters here because the sheet is 17 columns wide: with
    default column widths every name is clipped and each column has to be
    dragged out by hand. Columns are sized to their actual content, the
    header row is frozen and filtered, and cells are top-left aligned
    without wrapping so row heights stay uniform.

    names_only drops the (mostly empty) per-role EMAIL columns, halving the
    width to a sheet that fits on screen.
    """
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    records = []
    for row in rows:
        record = {"Company": row.company}
        for role in ROLES:
            name, email = row.values.get(role, ("", ""))
            record[role] = name
            if not names_only:
                record[f"{role} EMAIL"] = email
        record["Roles Found"] = sum(1 for name, _ in row.values.values() if name)
        if not names_only:
            record["Error"] = row.error
        records.append(record)

    df = pd.DataFrame(records)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheet = "Leadership Contacts"
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]

        header_fill = PatternFill("solid", fgColor="1F3864")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=False)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

        # Size each column to its typical content rather than its longest
        # value: across thousands of rows a handful of outliers (a company
        # name listing three founders, say) would otherwise push every
        # column to the maximum and make the sheet unreadably wide. The 95th
        # percentile fits virtually every cell while keeping columns narrow.
        body_align = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            if lengths:
                typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)]
            else:
                typical = 0
            # Email columns are mostly blank, so their few populated values
            # are long; cap them tighter than name columns since they are
            # secondary information here.
            ceiling = 24 if str(column).endswith("EMAIL") else 30
            width = min(max(typical + 2, min(len(str(column)) + 2, 18), 10), ceiling)
            ws.column_dimensions[get_column_letter(idx)].width = width
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body_align

        ws.freeze_panes = "B2"  # keep the header and company name in view
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 26

    return output_path
