"""Web-verify renewable landscape: keep distributors/channel partners only.

Uses local SearXNG + Excel text. Resumable checkpoint.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")
load_dotenv(Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env"), override=False)

from openpyxl import load_workbook

from vendor_intel.clients.searxng import searxng_search_any, searxng_urls
from vendor_intel.export.landscape_slim import SLIM_COLUMNS
from vendor_intel.pipeline.web_expand import write_final_xlsx

OUT_DIR = ROOT / "output" / "web_expand" / "global_renewable_energy_technologies_market_global"
XLSX = OUT_DIR / "global_renewable_energy_technologies_market_global_FINAL.xlsx"
OUT_XLSX = OUT_DIR / "global_renewable_energy_technologies_market_global_CHANNEL_ONLY.xlsx"
REPORT = OUT_DIR / "channel_web_verify_report.json"
CKPT = OUT_DIR / "channel_web_verify_checkpoint.json"

OEM_BLOCK = {
    "sungrow", "sungrow power supply", "sungrow europe", "huawei", "huawei fusionsolar",
    "huawei technologies", "growatt", "growatt new energy", "delta electronics",
    "delta electronics europe", "abb", "siemens", "siemens energy", "siemens gamesa",
    "general electric", "general electric ge", "ge vernova", "mitsubishi electric",
    "toshiba", "toshiba energy systems solutions", "panasonic", "panasonic heating cooling",
    "lg electronics", "lg energy solution", "byd", "rec group", "solaria", "silfab solar",
    "q cells", "hanwha q cells", "hanwha corporation", "hanwha energy corporation",
    "first solar", "jinko solar", "jinkosolar", "trina solar", "trinasolar",
    "canadian solar", "longi", "longi solar", "longi green energy", "ja solar",
    "enphase", "enphase energy", "solaredge", "solaredge technologies", "tesla",
    "vestas", "nordex", "nordex group", "goldwind", "mingyang", "schneider electric",
    "eaton", "sma", "sma solar", "sma solar technology ag", "fronius", "fronius international",
    "goodwe", "ginlong", "solis", "catl", "samsung sdi", "orsted", "iberdrola",
    "nextera energy", "nextera energy resources", "enel", "enel green power", "enel x",
    "engie", "edf renewables", "bp", "shell", "shell energy uk", "totalenergies",
    "equinor", "rwe", "rwe renewables", "acciona energy", "acciona energia",
    "mission solar energy", "risen energy", "talesun solar", "yingli solar",
    "sunpower corporation", "vikram solar", "renewsys india", "kyocera corporation",
    "emmvee photovoltaic power", "websol energy system", "solar frontier japan",
    "hyosung heavy industries", "jakson group", "adani green energy", "chint new energy",
    "tongwei solar", "waaree energies", "waaree energies ltd", "suzlon energy",
    "victron energy", "yaskawa electric", "varta ag", "tesvolt ag",
}

CHANNEL_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|dealer|channel partner|"
    r"value[- ]added distributor|\bvad\b|authorized dealer|stockist|importer|"
    r"system integrator|epc contractor|solar installer|installation company|"
    r"merchant wholesaler|electrical distributor|pv distributor)\b",
    re.I,
)
OEM_RE = re.compile(
    r"\b(manufacturer|manufactures|oem|module maker|inverter manufacturer|"
    r"turbine manufacturer|produces (?:solar|wind|pv|inverters|modules)|"
    r"leading manufacturer|global manufacturer|tier[- ]?1 manufacturer)\b",
    re.I,
)
UTILITY_RE = re.compile(
    r"\b(utility|independent power producer|\bipp\b|power producer|asset owner|"
    r"project developer|owns and operates|energy retailer|electricity supplier)\b",
    re.I,
)
RENEW_RE = re.compile(
    r"\b(solar|pv|photovoltaic|wind|renewable|inverter|energy storage|battery|bess|"
    r"clean energy|turbine|hydrogen|microgrid|heat pump|greentech)\b",
    re.I,
)
NAME_CHANNEL_RE = re.compile(
    r"\b(distributor|distributors|wholesale|wholesaler|dealer|dealers|reseller)\b",
    re.I,
)


def _norm(s: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def _is_blocked_oem(name: str) -> bool:
    key = _norm(name)
    if key in OEM_BLOCK:
        return True
    for m in OEM_BLOCK:
        if key == m or key.startswith(m + " "):
            return True
        if len(m.split()) == 1 and len(m) >= 5:
            if re.fullmatch(rf"{re.escape(m)}( solar| energy| power| technologies| technology)?", key):
                return True
    return False


def _score(blob: str) -> dict[str, int]:
    return {
        "channel": len(CHANNEL_RE.findall(blob)),
        "oem": len(OEM_RE.findall(blob)),
        "utility": len(UTILITY_RE.findall(blob)),
        "renew": len(RENEW_RE.findall(blob)),
    }


def _excel_blob(row: dict[str, str]) -> str:
    """Use specialty/categories/brands only — skip Summary/Distribution Type defaults."""
    return "\n".join(
        str(row.get(k) or "")
        for k in (
            "Specialty Focus",
            "Core Categories",
            "Key Brands Represented",
        )
    )


def _scrub_excel_for_score(text: str) -> str:
    """Remove pipeline boilerplate that falsely looks like channel evidence."""
    t = text or ""
    t = re.sub(r"(?i)channel participant[^.]*\.?", " ", t)
    t = re.sub(r"(?i)verify:\s*", " ", t)
    t = re.sub(r"(?i)\bnot publicly disclosed\b", " ", t)
    # Bare default "Distributor" alone is not evidence
    t = re.sub(r"(?i)^distributor$", " ", t.strip())
    return t


def _decide(name: str, scores: dict[str, int], *, excel_blob: str = "") -> tuple[str, str]:
    if _is_blocked_oem(name):
        return "drop_oem", "known_oem_blocklist"

    if NAME_CHANNEL_RE.search(name) and (
        RENEW_RE.search(name) or RENEW_RE.search(excel_blob) or scores["renew"] > 0
    ):
        return "keep", "name_channel_marker"

    ex = _score(f"{name}\n{_scrub_excel_for_score(excel_blob)}")
    web_ch, web_oem, web_util, web_renew = (
        scores["channel"],
        scores["oem"],
        scores["utility"],
        scores["renew"],
    )
    ch = web_ch + ex["channel"]
    oem = web_oem + ex["oem"]
    util = web_util + ex["utility"]
    renew = web_renew + ex["renew"] + (1 if RENEW_RE.search(name) else 0)

    if renew == 0:
        return "drop_unrelated", "no_renewable_signal"

    # Prefer WEB evidence for OEM/utility drops
    if web_oem >= 2 and web_ch == 0:
        return "drop_oem", f"web_oem={web_oem} web_channel=0"
    if web_oem >= 1 and web_ch == 0 and ex["channel"] == 0:
        return "drop_oem", f"oem_only web_oem={web_oem}"
    if web_util >= 2 and web_ch == 0:
        return "drop_utility", f"web_utility={web_util}"

    # Strong web channel
    if web_ch >= 1 and web_ch >= web_oem:
        return "keep", f"web_channel={web_ch} web_oem={web_oem} renew={renew}"

    # Name contains Solar/Wind/… AND web has no OEM/utility — likely installer/dealer
    if RENEW_RE.search(name) and web_oem == 0 and web_util == 0 and web_renew + (1 if RENEW_RE.search(name) else 0) >= 1:
        # Require at least one channel-ish specialty word OR name channel OR web renew hits besides name
        if ex["channel"] >= 1 or NAME_CHANNEL_RE.search(excel_blob) or web_renew >= 1:
            return "keep", "renewable_named_channelish"
        # Still keep pure "* Solar" names without OEM web signal (common for dealers)
        if re.search(r"(?i)\b(solar|wind|renewable)\b", name) and web_oem == 0:
            return "keep", "renewable_named_no_oem_web"

    # Specialty text explicitly channel + renewable
    if ex["channel"] >= 1 and renew >= 1 and web_oem == 0:
        return "keep", f"specialty_channel={ex['channel']}"

    if web_oem >= 1 and web_ch == 0:
        return "drop_oem", f"weak_oem oem={web_oem}"
    if util >= 2 and ch == 0:
        return "drop_utility", f"utility={util}"
    return "drop_unclear", f"web_ch={web_ch} web_oem={web_oem} renew={renew} util={util}"


def _load_rows() -> list[dict[str, str]]:
    wb = load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb.active
    raw = list(ws.iter_rows(values_only=True))
    wb.close()
    headers = [str(h or "").strip() for h in raw[0]]
    out: list[dict[str, str]] = []
    for r in raw[1:]:
        if not r or not str(r[0] or "").strip():
            continue
        d = {h: "" for h in SLIM_COLUMNS}
        for i, h in enumerate(headers):
            if h in d:
                d[h] = "" if r[i] is None else str(r[i])
        out.append(d)
    return out


async def _verify_one(
    name: str,
    sem: asyncio.Semaphore,
    urls: list[str],
    *,
    excel_blob: str,
) -> dict:
    async with sem:
        q = (
            f'"{name}" (solar OR renewable OR wind OR PV) '
            f'(distributor OR wholesaler OR dealer OR manufacturer OR OEM OR '
            f'"channel partner" OR installer)'
        )
        try:
            hits, _base = await searxng_search_any(q, urls, max_results=10)
        except Exception as exc:  # noqa: BLE001
            return {
                "company": name,
                "verdict": "drop_unclear",
                "reason": f"search_error:{type(exc).__name__}",
                "scores": {},
                "snippets_preview": str(exc)[:200],
                "hit_count": 0,
            }
        bits = [
            f"{h.title} | {h.snippet} | {h.link}"
            for h in (hits or [])
        ]
        snippets = "\n".join(bits)[:5000]
        scores = _score(f"{name}\n{snippets}")
        verdict, reason = _decide(name, scores, excel_blob=excel_blob)
        return {
            "company": name,
            "verdict": verdict,
            "reason": reason,
            "scores": scores,
            "snippets_preview": snippets[:400],
            "hit_count": len(hits or []),
        }


async def main() -> None:
    rows = _load_rows()
    print(f"loaded {len(rows)} rows", flush=True)

    pre_drop: list[dict] = []
    todo_rows: list[dict[str, str]] = []
    for r in rows:
        name = str(r.get("Company") or "").strip()
        if _is_blocked_oem(name):
            pre_drop.append(
                {"company": name, "verdict": "drop_oem", "reason": "known_oem_blocklist", "scores": {}}
            )
        else:
            todo_rows.append(r)
    print(f"OEM blocklist drop={len(pre_drop)}; web-verify={len(todo_rows)}", flush=True)

    # Fresh run (classifier changed) — ignore old checkpoint
    if (os.getenv("CHANNEL_VERIFY_FRESH") or "1").strip() in ("1", "true", "yes"):
        if CKPT.exists():
            CKPT.unlink()
            print("cleared old checkpoint (CHANNEL_VERIFY_FRESH)", flush=True)

    done: dict[str, dict] = {}
    pending = todo_rows
    print(f"pending web searches: {len(pending)}", flush=True)

    urls = searxng_urls((os.getenv("SEARXNG_BASE_URL") or "http://127.0.0.1:8080").strip())
    print(f"searxng urls: {urls}", flush=True)
    conc = int(os.getenv("CHANNEL_VERIFY_CONCURRENT") or "10")
    sem = asyncio.Semaphore(max(1, conc))
    results: list[dict] = []

    async def run_one(i: int, row: dict[str, str]) -> dict:
        name = str(row.get("Company") or "").strip()
        print(f"  [{i}/{len(pending)}] {name}", flush=True)
        return await _verify_one(name, sem, urls, excel_blob=_excel_blob(row))

    chunk = 30
    for start in range(0, len(pending), chunk):
        batch = pending[start : start + chunk]
        part = await asyncio.gather(*[run_one(start + j + 1, r) for j, r in enumerate(batch)])
        for item in part:
            done[item["company"]] = item
            results.append(item)
        CKPT.write_text(
            json.dumps({"results": list(done.values())}, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        keeps = sum(1 for x in done.values() if x.get("verdict") == "keep")
        print(f"  checkpoint {len(done)}/{len(todo_rows)} (keep so far {keeps})", flush=True)

    by = {x["company"]: x for x in results}
    kept_rows: list[dict[str, str]] = []
    dropped: list[dict] = list(pre_drop)
    counts: Counter[str] = Counter(x["verdict"] for x in pre_drop)

    for row in todo_rows:
        name = str(row.get("Company") or "").strip()
        v = by.get(name) or {"company": name, "verdict": "drop_unclear", "reason": "missing"}
        counts[v["verdict"]] += 1
        if v["verdict"] == "keep":
            kept_rows.append(row)
        else:
            dropped.append(v)

    print("\n=== COUNTS ===", flush=True)
    for k, n in counts.most_common():
        print(f"  {n:4d}  {k}", flush=True)
    print(f"KEEP {len(kept_rows)} / {len(rows)}", flush=True)

    REPORT.write_text(
        json.dumps(
            {
                "source": str(XLSX),
                "total_in": len(rows),
                "kept": len(kept_rows),
                "dropped": len(dropped),
                "counts": dict(counts),
                "kept_companies": [r.get("Company") for r in kept_rows],
                "dropped_detail": dropped,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    write_final_xlsx(
        OUT_XLSX,
        kept_rows,
        "Renewable Channel",
        {
            "filter": "searxng_web_verified_channel_only",
            "kept": len(kept_rows),
            "dropped": len(dropped),
        },
    )
    print(f"wrote {OUT_XLSX}", flush=True)
    print(f"wrote {REPORT}", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
