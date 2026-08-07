#!/usr/bin/env python3
"""Fill ONLY empty/wrong Contact Person with CEO name; set Role=Chief Executive Officer for those rows."""
from __future__ import annotations

import json
import re
import shutil
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

ROOT = Path(__file__).resolve().parents[1]
FINAL = (
    ROOT
    / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx"
)
SHEET = "Companies"
NPD = "Not publicly disclosed"
ROLE = "Chief Executive Officer"
WORKERS = 8

NPD_SET = {
    "", "n/a", "na", "none", "-", "—", "not publicly disclosed",
    "not publicly available", "unknown", "tbd", "todo", "full name or empty",
    "title or empty",
}
TITLE_NAME = re.compile(
    r"^(chief\s*executive\s*officer|ceo|cto|cfo|founder|co-?founder|president|"
    r"managing\s*director|md|administrador(\s+único)?|director|sales\s*director|"
    r"national\s*sales.*|chair(man|person)?)$",
    re.I,
)
GARBAGE = re.compile(
    r"\b(said|photo\s+is|linkedin|https?://|www\.|@|appoints|visionary|full name)\b",
    re.I,
)
NAME_OK = re.compile(r"^[A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,4}$")

KNOWN_CEO = {
    "amazon": "Andy Jassy",
    "best buy": "Corie Barry",
    "currys": "Alex Baldock",
    "cdw": "Christine A. Leahy",
    "also holding": "Gustavo Möller-Hergt",
    "ab sa": "Andrzej Przybyło",
    "ingram micro": "Paul Bay",
    "td synnex": "Rich Hume",
    "arrow electronics": "Sean J. Kerins",
    "bluechip infotech": "Johnson Hsiung",
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _clean_name(s: str) -> str:
    s = re.sub(r"\s+", " ", (s or "").strip())
    s = re.sub(r"^(mr\.?|mrs\.?|ms\.?|dr\.?)\s+", "", s, flags=re.I)
    s = re.sub(r"\s+(is\s+)?(founder|ceo|president|managing director).*$", "", s, flags=re.I)
    s = s.strip(" ,;-")
    if not s or _norm(s) in NPD_SET or TITLE_NAME.match(s) or GARBAGE.search(s):
        return ""
    if any(x in s.lower() for x in (" wireless", " ltd", " llc", " inc", " gmbh", " group")):
        if len(s.split()) < 2:
            return ""
    if len(s) < 4 or len(s) > 60 or any(ch.isdigit() for ch in s):
        return ""
    if not NAME_OK.match(s) and not re.match(r"^[\wÀ-ÖØ-öø-ÿ.'’\- ]+$", s, re.I):
        return ""
    return s


def is_bad_contact(contact: str) -> bool:
    c = (contact or "").strip()
    if not c or _norm(c) in NPD_SET:
        return True
    if TITLE_NAME.match(c) or GARBAGE.search(c):
        return True
    if not _clean_name(c):
        return True
    return False


def wiki_text(name: str) -> str:
    try:
        q = urllib.parse.quote(name)
        url = (
            "https://en.wikipedia.org/w/api.php?action=query&prop=extracts"
            f"&explaintext=1&redirects=1&format=json&titles={q}"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "VendorIntelCEO/1.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="ignore"))
        pages = (data.get("query") or {}).get("pages") or {}
        for p in pages.values():
            if int(p.get("pageid") or 0) < 0:
                continue
            t = p.get("extract") or ""
            if len(t) > 40:
                return t[:8000]
    except Exception:
        return ""
    return ""


def extract_ceo(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"(?:chief executive officer|CEO)\s*(?:is|:)\s*([A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
        r"([A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})\s*(?:is|,)\s*(?:the\s+)?(?:current\s+)?(?:CEO|Chief Executive Officer)",
        r"CEO[:\s]+([A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
        r"Managing Director[:\s]+([A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
        r"led by\s+(?:CEO\s+)?([A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            name = _clean_name(m.group(1))
            if name:
                return name
    return ""


def web_snippets(name: str, website: str) -> str:
    host = ""
    if website:
        host = re.sub(r"^https?://", "", website, flags=re.I).split("/")[0].replace("www.", "")
    queries = [f'"{name}" CEO OR "Chief Executive Officer" OR "Managing Director"']
    if host:
        queries.insert(0, f'"{name}" {host} CEO OR "Chief Executive Officer"')
    parts = []
    try:
        from ddgs import DDGS

        for q in queries[:2]:
            try:
                hits = list(DDGS().text(q, max_results=5)) or []
            except Exception:
                hits = []
            for h in hits:
                parts.append(f"{h.get('title','')} {h.get('body','')}")
    except Exception:
        pass
    return "\n".join(parts)[:7000]


def resolve_bad(row: dict) -> dict:
    name = row.get("Company") or ""
    known = KNOWN_CEO.get(_norm(name), "")
    contact = known
    src = "known" if known else ""
    if not contact:
        contact = extract_ceo(wiki_text(name))
        if contact:
            src = "wikipedia"
    if not contact:
        contact = extract_ceo(web_snippets(name, row.get("Website") or ""))
        if contact:
            src = "web"
    if not contact:
        contact = NPD
        src = src or "none"
    return {"company": name, "contact": contact, "role": ROLE, "source": src}


def load_rows(path: Path):
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    headers = []
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            hi = i
            headers = vals
            break
    data = []
    for r in rows[hi + 1 :]:
        d = {
            headers[j]: ("" if j >= len(r) or r[j] is None else str(r[j]).strip())
            for j in range(len(headers))
            if headers[j]
        }
        if d.get("Company"):
            data.append(d)
    return [h for h in headers if h], data


def write_xlsx(path, headers, rows, title, meta):
    wb = load_workbook(path)
    ws = wb[SHEET]
    hr = 2
    for r in range(1, 6):
        if ws.cell(r, 1).value == "Company":
            hr = r
            break
    if ws.max_row > hr:
        ws.delete_rows(hr + 1, ws.max_row - hr)
    for j, h in enumerate(headers, 1):
        ws.cell(hr, j).value = h
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    for i, row in enumerate(rows):
        rr = hr + 1 + i
        for j, h in enumerate(headers, 1):
            cell = ws.cell(rr, j, row.get(h, ""))
            cell.alignment = wrap
            cell.border = thin
    ws.cell(1, 1).value = title
    name = "CEO Bad-Only Audit"
    if name in wb.sheetnames:
        del wb[name]
    aus = wb.create_sheet(name)
    aus.append(["field", "value"])
    for k, v in meta.items():
        aus.append([k, str(v)])
    wb.save(path)


def main():
    headers, rows = load_rows(FINAL)
    bad_idx = [i for i, r in enumerate(rows) if is_bad_contact(r.get("Contact Person") or "")]
    print(f"=== CEO fill BAD/EMPTY only ===", flush=True)
    print(f"  total={len(rows)}  to_fix={len(bad_idx)}  keep={len(rows)-len(bad_idx)}", flush=True)
    for i in bad_idx:
        print(f"  - {rows[i].get('Company')}: {rows[i].get('Contact Person')!r}", flush=True)

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(resolve_bad, rows[i]): i for i in bad_idx}
        for fut in as_completed(futs):
            i = futs[fut]
            try:
                res = fut.result()
            except Exception as err:
                res = {
                    "company": rows[i].get("Company"),
                    "contact": NPD,
                    "role": ROLE,
                    "source": f"error:{err}",
                }
            results[i] = res
            print(
                f"  → {res['company']}: {res['contact']} ({res['source']})",
                flush=True,
            )

    out = []
    audits = []
    for i, row in enumerate(rows):
        fixed = dict(row)
        if i in results:
            res = results[i]
            fixed["Contact Person"] = res["contact"]
            fixed["Role"] = ROLE
            src = fixed.get("Data Sources") or ""
            if "ceo_bad_only_fill" not in src:
                fixed["Data Sources"] = (src + "; ceo_bad_only_fill").strip("; ")
            audits.append({**res, "before": row.get("Contact Person")})
        out.append(fixed)

    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    bak = FINAL.with_suffix(FINAL.suffix + f".bak_ceo_bad_{ts}")
    shutil.copy2(FINAL, bak)
    audit_path = FINAL.parent / f"smartphone_ceo_bad_only_{ts}.json"
    audit_path.write_text(json.dumps({"ts": ts, "fixed": audits}, indent=2, ensure_ascii=False), encoding="utf-8")
    still_bad = sum(1 for r in out if is_bad_contact(r.get("Contact Person") or ""))
    # title-as-name should be 0 among contacts that aren't NPD
    title_left = sum(1 for r in out if TITLE_NAME.match((r.get("Contact Person") or "").strip()))
    write_xlsx(
        FINAL,
        headers,
        out,
        f"Global Smartphone Market — channel partners (CEO bad-only fill {ts})",
        {
            "ts": ts,
            "fixed_rows": len(audits),
            "kept_rows": len(rows) - len(audits),
            "still_bad_or_npd": still_bad,
            "title_as_name_left": title_left,
            "backup": str(bak),
            "audit": str(audit_path),
        },
    )
    print("=== DONE ===", flush=True)
    print(f"  fixed={len(audits)} kept={len(rows)-len(audits)} still_bad/npd={still_bad}", flush=True)
    print(f"  title-as-name left={title_left}", flush=True)
    print(f"  backup={bak.name}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
