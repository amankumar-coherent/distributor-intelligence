#!/usr/bin/env python3
"""Fill Contact Person = CEO name, Role = Chief Executive Officer (Wikipedia + web)."""
from __future__ import annotations

import json
import re
import shutil
import urllib.parse
import urllib.request
from collections import Counter
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
WORKERS = 12

# Well-known current CEOs
KNOWN_CEO: dict[str, str] = {
    "amazon": "Andy Jassy",
    "best buy": "Corie Barry",
    "currys": "Alex Baldock",
    "cdw": "Christine A. Leahy",
    "also holding": "Gustavo Möller-Hergt",
    "ab sa": "Andrzej Przybyło",
    "ingram micro": "Paul Bay",
    "td synnex": "Rich Hume",
    "arrow electronics": "Sean J. Kerins",
    "computacenter": "Mike Norris",
    "dicker data": "David Dicker",
    "esprinet": "Alessandro Cattani",
    "shi international": "Thai Lee",
    "insight enterprises": "Joyce Mullen",
    "softcat": "Graham Charlton",
    "ban leong technologies": "Ronald Teng",
    "asbis": "Siarhei Kostevitch",
    "3g mobile": "AK Shaji",
    "zones": "Firoz Lalji",
    "mediamarktsaturn": "Dr. Karsten Wildberger",
}

NPD_SET = {
    "", "n/a", "na", "none", "-", "—", "not publicly disclosed",
    "not publicly available", "unknown", "tbd", "todo", "full name or empty",
}
TITLE_NAME = re.compile(
    r"^(chief\s*executive\s*officer|ceo|cto|cfo|founder|co-?founder|president|"
    r"managing\s*director|md|administrador(\s+único)?|director|sales\s*director|"
    r"national\s*sales.*|chair(man|person)?)$",
    re.I,
)
GARBAGE = re.compile(
    r"\b(said|photo\s+is|linkedin|https?://|www\.|@|appoints|visionary)\b",
    re.I,
)
NAME_OK = re.compile(r"^[A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+){1,4}$")


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _clean_name(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"^(mr\.?|mrs\.?|ms\.?|dr\.?)\s+", "", s, flags=re.I)
    s = re.sub(r"\s+(is\s+)?(founder|ceo|president|managing director).*$", "", s, flags=re.I)
    s = s.strip(" ,;-")
    if not s or _norm(s) in NPD_SET or TITLE_NAME.match(s) or GARBAGE.search(s):
        return ""
    if not NAME_OK.match(s):
        # allow names with comma Jr etc lightly
        if len(s) < 4 or len(s) > 60 or any(ch.isdigit() for ch in s):
            return ""
    # reject company-like
    if any(x in s.lower() for x in (" wireless", " group", " ltd", " llc", " inc", " gmbh")):
        return ""
    return s


def _looks_bad(contact: str) -> bool:
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
        # opensearch
        surl = (
            "https://en.wikipedia.org/w/api.php?action=opensearch&limit=1&namespace=0"
            f"&format=json&search={q}"
        )
        req2 = urllib.request.Request(surl, headers={"User-Agent": "VendorIntelCEO/1.0"})
        with urllib.request.urlopen(req2, timeout=8) as resp:
            arr = json.loads(resp.read().decode("utf-8", errors="ignore"))
        if isinstance(arr, list) and len(arr) > 1 and arr[1]:
            return wiki_text(arr[1][0])
    except Exception:
        return ""
    return ""


def extract_ceo(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"(?:chief executive officer|CEO)\s*(?:is|:)\s*([A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
        r"([A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})\s*(?:is|,)\s*(?:the\s+)?(?:current\s+)?(?:CEO|Chief Executive Officer)",
        r"CEO[:\s]+([A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
        r"Chief Executive Officer[:\s]+([A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
        r"led by\s+(?:CEO\s+)?([A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+(?:\s+[A-Z][a-zÀ-ÖØ-öø-ÿ.'’\-]+){1,3})",
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
    queries = [
        f'"{name}" CEO OR "Chief Executive Officer"',
        f'"{name}" "Chief Executive Officer" OR CEO named OR appointed',
    ]
    if host:
        queries.insert(0, f'"{name}" {host} CEO OR "Chief Executive Officer"')
    parts: list[str] = []
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


def resolve(row: dict) -> dict:
    name = row.get("Company") or ""
    known = KNOWN_CEO.get(_norm(name), "")
    contact = known
    src = "known" if known else ""

    if not contact:
        wiki = wiki_text(name)
        contact = extract_ceo(wiki)
        if contact:
            src = "wikipedia"

    if not contact:
        web = web_snippets(name, row.get("Website") or "")
        contact = extract_ceo(web)
        if contact:
            src = "web"

    # keep existing if it already looks like a real person name and role was founder/president
    # but user wants CEO — only keep existing if we failed to find and existing is clean person
    existing = _clean_name(row.get("Contact Person") or "")
    if not contact and existing and not _looks_bad(existing):
        contact = existing
        src = "kept_existing"

    if not contact:
        contact = NPD
        src = src or "none"

    return {
        "company": name,
        "contact": contact,
        "role": ROLE,
        "source": src,
    }


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
    name = "CEO Fill Audit"
    if name in wb.sheetnames:
        del wb[name]
    aus = wb.create_sheet(name)
    aus.append(["field", "value"])
    for k, v in meta.items():
        aus.append([k, str(v)])
    wb.save(path)


def main():
    headers, rows = load_rows(FINAL)
    print(f"=== CEO fill (Contact Person + Role) ===", flush=True)
    print(f"  companies={len(rows)} workers={WORKERS}", flush=True)

    # Always set Role=CEO; refresh contact when bad OR always try web for accuracy
    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(resolve, row): row for row in rows}
        done = 0
        for fut in as_completed(futs):
            row = futs[fut]
            done += 1
            try:
                res = fut.result()
            except Exception as err:
                existing = _clean_name(row.get("Contact Person") or "")
                res = {
                    "company": row.get("Company"),
                    "contact": existing or NPD,
                    "role": ROLE,
                    "source": f"error:{err}",
                }
            results[_norm(res["company"])] = res
            if done % 25 == 0 or done == len(rows):
                print(f"  progress {done}/{len(rows)}", flush=True)

    out = []
    audits = []
    filled = 0
    for row in rows:
        res = results[_norm(row.get("Company") or "")]
        fixed = dict(row)
        before = fixed.get("Contact Person") or ""
        fixed["Contact Person"] = res["contact"]
        fixed["Role"] = ROLE
        if res["contact"] != NPD and res["contact"] != before:
            filled += 1
        src = fixed.get("Data Sources") or ""
        if "ceo_web_fill" not in src:
            fixed["Data Sources"] = (src + "; ceo_web_fill").strip("; ")
        out.append(fixed)
        audits.append(res)

    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    bak = FINAL.with_suffix(FINAL.suffix + f".bak_ceo_{ts}")
    shutil.copy2(FINAL, bak)
    contact_c = Counter(r["Contact Person"] for r in out)
    role_c = Counter(r["Role"] for r in out)
    npd_n = sum(1 for r in out if r["Contact Person"] == NPD)
    title_as = sum(1 for r in out if TITLE_NAME.match(r["Contact Person"] or ""))
    audit_path = FINAL.parent / f"smartphone_ceo_fill_{ts}.json"
    audit_path.write_text(
        json.dumps(
            {
                "ts": ts,
                "filled_changed": filled,
                "npd": npd_n,
                "role_counts": dict(role_c),
                "audits": audits,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    write_xlsx(
        FINAL,
        headers,
        out,
        f"Global Smartphone Market — channel partners (CEO contacts {ts})",
        {
            "ts": ts,
            "method": "known+wikipedia+web",
            "rows": len(out),
            "npd_contacts": npd_n,
            "title_as_name_left": title_as,
            "role": ROLE,
            "backup": str(bak),
            "audit": str(audit_path),
        },
    )
    print("=== DONE ===", flush=True)
    print(f"  Role all -> {ROLE}: {dict(role_c)}", flush=True)
    print(f"  Contact NPD: {npd_n}  title-as-name left: {title_as}", flush=True)
    print(f"  changed names: {filled}", flush=True)
    print(f"  backup: {bak.name}", flush=True)
    for want in (
        "Amazon", "Best Buy", "CDW", "ALSO Holding", "AB SA",
        "Amir Group FZCO", "Anything America", "Arrow Electronics", "Currys",
    ):
        r = results.get(_norm(want))
        if r:
            print(f"  {want}: {r['contact']} | {r['role']} ({r['source']})", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
