"""Ask DeepSeek for flexible-packaging DISTRIBUTOR candidates, then DNS/HTTP verify.

Adds only verified channel/distributor rows (not manufacturers/OEMs).
"""
from __future__ import annotations

import json
import os
import re
import socket
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

FP = ROOT / (
    "output/web_expand/global_flexible_packaging_market_global"
    "/global_flexible_packaging_market_global_FINAL.xlsx"
)
OUT_JSON = FP.parent / "channel_add_200_verified.json"
EXISTING = FP.parent / "_existing_names.txt"

TARGET = 200

OEM_BLOCK = re.compile(
    r"(?i)\b("
    r"amcor|mondi|huhtamaki|coveris|constantia|winpak|printpack|sonoco|proampac|"
    r"sealed\s*air|berry\s*global|suedpack|südpack|wipak|novolex|cosmo\s*films|"
    r"uflex|polyplex|jindal\s*poly|bischof|glenroy|orora|sig\b|smurfit|ahlstrom|"
    r"rkw\b|walki|gualapack|perlen|korozo|danaflex|flair\s*flexible|epac|"
    r"american\s*packaging\s*corporation|ccl\s*industries|transcontinental|"
    r"manufacturer|converter|extruder|we\s+manufacture|our\s+plants?\b"
    r")\b"
)

CHANNEL_HINT = re.compile(
    r"(?i)\b("
    r"distributor|distribution|wholesale|wholesaler|reseller|merchant|"
    r"channel|supplier\s+of|stockist|packaging\s+supplies|shipping\s+supplies|"
    r"industrial\s+supplies|packaging\s+distributor|poly\s*bag\s+distributor|"
    r"catalog|next[- ]day\s+delivery|in[- ]stock"
    r")\b"
)


def domain_of(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    host = urlparse(u).netloc.lower().split("@")[-1]
    return host[4:] if host.startswith("www.") else host


def normalize_url(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    return u.rstrip("/")


def probe(url: str) -> tuple[bool, str, str]:
    """Return (ok, final_url_or_note, title_snippet)."""
    u = normalize_url(url)
    host = domain_of(u)
    if not host or "." not in host:
        return False, "bad_host", ""
    try:
        socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
    except Exception:
        try:
            socket.getaddrinfo(host, 80, type=socket.SOCK_STREAM)
        except Exception as e:
            return False, f"dns_fail:{e}", ""
    ctx = ssl.create_default_context()
    last = "http_fail"
    for attempt in (u, "http://" + host):
        try:
            req = Request(attempt, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
            kw = {"timeout": 12}
            if attempt.startswith("https"):
                kw["context"] = ctx
            with urlopen(req, **kw) as resp:
                code = getattr(resp, "status", 200)
                final = resp.geturl()
                body = resp.read(12000).decode("utf-8", "ignore")
                title = ""
                m = re.search(r"<title[^>]*>(.*?)</title>", body, re.I | re.S)
                if m:
                    title = re.sub(r"\s+", " ", m.group(1)).strip()[:120]
                text = re.sub(r"<[^>]+>", " ", body)
                text = re.sub(r"\s+", " ", text).lower()
                if code >= 400:
                    last = f"http_{code}"
                    continue
                return True, final, title + " || " + text[:1500]
        except Exception as e:
            last = str(e)[:100]
            continue
    return False, last, ""


def deepseek_candidates(batch_n: int, avoid: list[str]) -> list[dict]:
    client = OpenAI(
        api_key=os.environ["OPENAI_API_KEY"],
        base_url=os.environ.get("OPENAI_BASE_URL") or "https://api.deepseek.com/v1",
    )
    avoid_s = ", ".join(avoid[:120])
    prompt = f"""
You are a B2B packaging channel researcher.
List {batch_n} REAL companies that are DISTRIBUTORS / WHOLESALERS / CHANNEL PARTNERS
supplying flexible packaging materials (poly bags, stretch film, shrink film, pouches,
mailers, flexible films, packaging supplies) to businesses.

STRICT RULES:
- ONLY distributors, wholesalers, stockists, packaging supply merchants, industrial packaging distributors.
- DO NOT include manufacturers, converters, extruders, or global OEMs (no Amcor, Mondi, Huhtamaki, Sealed Air, ProAmpac, Printpack, Winpak, Sonoco, Coveris, Constantia, UFlex, Berry, Novolex, etc.).
- Companies must sell/distribute packaging to other businesses (channel), not primarily make films in their own plants as OEMs.
- Provide official website if known.
- Diversify geos: USA, Canada, UK, EU, Middle East, India, SEA, Australia, LatAm, Africa.
- Avoid these already-listed names: {avoid_s}

Return ONLY a JSON array of objects:
[{{"company":"...","website":"https://...","country":"...","role":"distributor|wholesaler|channel","why_channel":"one short reason"}}]
"""
    resp = client.chat.completions.create(
        model=os.getenv("OPENAI_MODEL") or "deepseek-chat",
        messages=[
            {"role": "system", "content": "Return valid JSON only. No markdown."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
    )
    raw = (resp.choices[0].message.content or "").strip()
    raw = re.sub(r"^```json\s*|\s*```$", "", raw, flags=re.I).strip()
    data = json.loads(raw)
    if not isinstance(data, list):
        raise ValueError("expected list")
    return data


def load_existing() -> tuple[set[str], set[str]]:
    wb = load_workbook(FP, read_only=True, data_only=True)
    ws = wb.active
    names = set()
    domains = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        n = str(r[0] or "").strip()
        w = str(r[1] or "").strip()
        if n:
            names.add(n.lower())
        d = domain_of(w)
        if d:
            domains.add(d)
    wb.close()
    return names, domains


def is_channel_text(blob: str) -> bool:
    if not blob:
        return False
    if OEM_BLOCK.search(blob) and not CHANNEL_HINT.search(blob):
        # manufacturer-heavy page without distributor language
        return False
    # Prefer channel signals
    return bool(CHANNEL_HINT.search(blob))


def main() -> None:
    existing_names, existing_domains = load_existing()
    print(f"existing rows names={len(existing_names)} domains={len(existing_domains)}")

    candidates: list[dict] = []
    # Ask DeepSeek in chunks for diversity
    for i, n in enumerate((80, 80, 80), 1):
        print(f"DeepSeek candidate batch {i} requesting {n}...")
        try:
            batch = deepseek_candidates(n, sorted(existing_names)[:80] + [c.get("company", "") for c in candidates])
            candidates.extend(batch)
            print(f"  got {len(batch)}")
        except Exception as e:
            print(f"  DeepSeek error: {type(e).__name__}: {e}")
        time.sleep(1)

    # Dedup candidates
    seen = set()
    uniq = []
    for c in candidates:
        name = str(c.get("company") or "").strip()
        web = normalize_url(str(c.get("website") or "").strip())
        if not name or not web:
            continue
        key = name.lower()
        dom = domain_of(web)
        if key in seen or key in existing_names or dom in existing_domains:
            continue
        if OEM_BLOCK.search(name):
            continue
        seen.add(key)
        uniq.append(
            {
                "company": name,
                "website": web,
                "country": str(c.get("country") or "").strip(),
                "role": str(c.get("role") or "distributor").strip(),
                "why_channel": str(c.get("why_channel") or "").strip(),
            }
        )
    print(f"unique candidates after dedupe: {len(uniq)}")

    verified: list[dict] = []
    rejected: list[dict] = []

    def check_one(c: dict) -> tuple[dict, bool, str]:
        ok, note, blob = probe(c["website"])
        if not ok:
            return c, False, f"site_fail:{note}"
        # channel verification
        why = (c.get("why_channel") or "") + " " + (blob or "")
        if OEM_BLOCK.search(c["company"]):
            return c, False, "name_oem"
        # Require channel-ish language OR explicit distributor role from model + live site
        if not is_channel_text(why) and c.get("role", "").lower() not in {"distributor", "wholesaler", "channel"}:
            return c, False, "not_channel_signals"
        # Extra reject if page screams manufacturer-only with plant language and no distributor
        low = (blob or "").lower()
        if ("our manufacturing" in low or "our plant" in low or "we extrude" in low) and not CHANNEL_HINT.search(low):
            return c, False, "manufacturer_page"
        c2 = dict(c)
        c2["verified_website"] = note if note.startswith("http") else c["website"]
        c2["title"] = (blob.split("||")[0].strip() if blob else "")[:120]
        return c2, True, "ok"

    print("Probing websites...")
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = [ex.submit(check_one, c) for c in uniq]
        for fut in as_completed(futs):
            c, ok, reason = fut.result()
            if ok:
                verified.append(c)
            else:
                rejected.append({**c, "reject": reason})

    print(f"verified={len(verified)} rejected={len(rejected)}")
    # Prefer verified with stronger channel why
    verified.sort(key=lambda x: (0 if CHANNEL_HINT.search(x.get("why_channel", "")) else 1, x["company"].lower()))
    selected = verified[:TARGET]
    print(f"selected={len(selected)}")

    OUT_JSON.write_text(json.dumps({"selected": selected, "rejected_sample": rejected[:50]}, indent=2, ensure_ascii=False), encoding="utf-8")

    if len(selected) < TARGET:
        print(f"WARNING: only {len(selected)} verified channel companies (<{TARGET}). Will add what we have.")

    # Append to Excel
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_add_channel_{stamp}.xlsx")
    import shutil

    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hidx = {str(h): i for i, h in enumerate(headers)}

    country_map = {
        "united states": ("US", "NA"),
        "usa": ("US", "NA"),
        "us": ("US", "NA"),
        "canada": ("CA", "NA"),
        "united kingdom": ("GB", "EU"),
        "uk": ("GB", "EU"),
        "germany": ("DE", "EU"),
        "france": ("FR", "EU"),
        "netherlands": ("NL", "EU"),
        "india": ("IN", "APAC"),
        "australia": ("AU", "APAC"),
        "new zealand": ("NZ", "APAC"),
        "singapore": ("SG", "APAC"),
        "uae": ("AE", "ME"),
        "saudi arabia": ("SA", "ME"),
        "mexico": ("MX", "LATAM"),
        "brazil": ("BR", "LATAM"),
        "south africa": ("ZA", "AF"),
        "ireland": ("IE", "EU"),
        "spain": ("ES", "EU"),
        "italy": ("IT", "EU"),
        "poland": ("PL", "EU"),
        "sweden": ("SE", "EU"),
        "belgium": ("BE", "EU"),
        "switzerland": ("CH", "EU"),
        "japan": ("JP", "APAC"),
        "china": ("CN", "APAC"),
        "malaysia": ("MY", "APAC"),
        "indonesia": ("ID", "APAC"),
        "thailand": ("TH", "APAC"),
        "philippines": ("PH", "APAC"),
        "turkey": ("TR", "EU"),
    }

    added = 0
    for c in selected:
        row = [None] * len(headers)
        row[hidx["Company"]] = c["company"]
        row[hidx["Website"]] = c.get("verified_website") or c["website"]
        row[hidx["Distribution Type"]] = "Specialty Distributor"
        row[hidx["Retail / E-commerce / Both"]] = "No"
        row[hidx["Core Categories"]] = "Flexible packaging supplies; films; bags; packaging materials (channel/distributor)"
        row[hidx["Specialty Focus"]] = c.get("why_channel") or "Distributes flexible packaging / packaging supplies to businesses"
        row[hidx["Summary"]] = f"Verified channel partner/distributor candidate. Role={c.get('role')}. {c.get('why_channel','')}".strip()
        country = (c.get("country") or "").strip().lower()
        cc = rc = None
        for k, (a, b) in country_map.items():
            if k in country:
                cc, rc = a, b
                break
        if cc and "Country Code" in hidx:
            row[hidx["Country Code"]] = cc
        if rc and "Region Code" in hidx:
            row[hidx["Region Code"]] = rc
        if country and "Headquarters" in hidx:
            # country only as HQ hint — not inventing city
            row[hidx["Headquarters"]] = c.get("country")
        if rc and "Continent / Geography" in hidx:
            row[hidx["Continent / Geography"]] = f"{rc}({c.get('country')})" if c.get("country") else rc
        ws.append(row)
        added += 1

    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(FP)
    print(f"backup: {backup}")
    print(f"ADDED={added} TOTAL_ROWS={ws.max_row-1}")
    print(f"json: {OUT_JSON}")


if __name__ == "__main__":
    main()
