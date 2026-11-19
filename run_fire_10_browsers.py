#!/usr/bin/env python3
"""Launch discovery for the 25-market Fire Safety batch, 10 markets at a
time, running EACH wave of 10 to completion (250 verified, or fully
exhausted) before moving to the next 10 -- per the user's explicit
instruction, not a round-robin across all 25 at once.

Generalized from run_10_browsers.py (which drove 10 workers across 4
drone markets). This drives 10 CONCURRENT MARKETS at once, one worker
each. A market whose current state/city has run dry moves to the next
unswept one and keeps going, rather than stopping the way the first
version of the drone launcher did.

DISCOVER-THEN-VERIFY, INTERLEAVED: discovery alone can never reach a
VERIFIED target -- verified only moves when a separate verification pass
judges the discovered backlog (this project's standing lesson, learned
repeatedly on the drone batches: "6 of 26 unchanged" for hours because
nothing had judged the new companies). So each round here is: every
worker in the wave does ONE region, then the wave stops, exports, and
verifies together (each market still gets its own Chromium profile for
verification -- one profile per market, never shared with discovery,
since a shared profile produces about:blank tabs). Then the wave resumes
discovery on whichever markets are still short. This repeats until every
market in the wave hits target or exhausts its search space, and ONLY
THEN does the next wave of 10 begin.

Each fire market gets its OWN Chromium profile (data/fire_profile_<slug>)
with its OWN staged captcha-raptor copy.

Usage:
    python run_fire_10_browsers.py
    python run_fire_10_browsers.py --target 250 --wave 10
"""
from __future__ import annotations

import argparse
import re
import json
import os
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from india_states import INDIA_STATES  # noqa: E402
from india_cities import INDIA_CITIES  # noqa: E402
import supervisor_all as S  # noqa: E402
import market_batches  # noqa: E402
from fire_markets import MARKETS_FIRE  # noqa: E402
from food_markets import MARKETS_FOOD  # noqa: E402
from wound_markets import GLOBAL_REGIONS, MARKETS_WOUND  # noqa: E402
from water_markets import BRAZIL_REGIONS, MARKETS_BRAZIL_RO  # noqa: E402
from chemical_markets import GLOBAL_REGIONS as CHEM_REGIONS, MARKETS_CHEMICAL  # noqa: E402
from ingredient_markets import GLOBAL_REGIONS as INGR_REGIONS, MARKETS_INGREDIENT  # noqa: E402
from geo_markets import MARKETS_GEO, REGIONS_BY_SLUG as GEO_REGIONS  # noqa: E402
from h2_markets import MARKETS_H2  # noqa: E402
from dc_markets import DC_REGIONS, MARKETS_DC  # noqa: E402
from caps_markets import CAPS_REGIONS, MARKETS_CAPS  # noqa: E402
from mg_markets import MARKETS_MG, MG_REGIONS  # noqa: E402
from ent_markets import ENT_REGIONS, MARKETS_ENT  # noqa: E402
from spice_markets import MARKETS_SPICE  # noqa: E402
from soft_markets import MARKETS_SOFT, SOFT_REGIONS  # noqa: E402
from vip_markets import MARKETS_VIP, VIP_REGIONS  # noqa: E402
from cmp_markets import CMP_REGIONS, MARKETS_CMP  # noqa: E402
from buf_markets import BUF_REGIONS, MARKETS_BUF  # noqa: E402
from moz_markets import MOZ_REGIONS, MARKETS_MOZ  # noqa: E402
from jpf_markets import JPF_REGIONS, MARKETS_JPF  # noqa: E402
from esu_markets import ESU_REGIONS, MARKETS_ESU  # noqa: E402
from adh_markets import ADH_REGIONS, MARKETS_ADH  # noqa: E402
from iww_markets import IWW_REGIONS, MARKETS_IWW  # noqa: E402
from eadh_markets import EADH_REGIONS, MARKETS_EADH  # noqa: E402
from lug_markets import LUG_REGIONS, MARKETS_LUG  # noqa: E402
from cphi_markets import CPHI_REGIONS, MARKETS_CPHI  # noqa: E402
from eth_markets import ETH_REGIONS, MARKETS_ETH  # noqa: E402
from sorb_markets import SORB_REGIONS, MARKETS_SORB  # noqa: E402
from bie_markets import BIE_REGIONS, MARKETS_BIE  # noqa: E402
from alc_markets import ALC_REGIONS, MARKETS_ALC  # noqa: E402
from cphin_markets import CPHIN_REGIONS, MARKETS_CPHIN  # noqa: E402
from vendor_intel.pipeline.market_research.state import market_slug  # noqa: E402

_INDIA_REGIONS = tuple(INDIA_STATES) + tuple(INDIA_CITIES)

# --batch picks which batch this launcher drives. Same wave / rolling logic
# for all; the runner script, profile prefix, log folder and the regions
# discovery rotates through (Indian states+cities, or countries) differ.
_BATCHES = {
    "fire": ("run_fire_market.py", "fire_profile_", "logs_fire", MARKETS_FIRE, _INDIA_REGIONS),
    "food": ("run_food_market.py", "food_profile_", "logs_food", MARKETS_FOOD, _INDIA_REGIONS),
    "wound": ("run_wound_market.py", "wound_profile_", "logs_wound", MARKETS_WOUND, GLOBAL_REGIONS),
    "brazil": ("run_brazil_market.py", "brazil_profile_", "logs_brazil", MARKETS_BRAZIL_RO, BRAZIL_REGIONS),
    "chem": ("run_chemical_market.py", "chem_profile_", "logs_chem", MARKETS_CHEMICAL, CHEM_REGIONS),
    "ingr": ("run_ingredient_market.py", "ingr_profile_", "logs_ingr", MARKETS_INGREDIENT, INGR_REGIONS),
    # India-only since 2026-10-14 (user): Indian states then cities, like
    # fire/food. (GEO_REGIONS = the earlier worldwide + focus plan.)
    # States only, no city sweeps (user, 2026-10-14).
    "geo": ("run_geo_market.py", "geo_profile_", "logs_geo", MARKETS_GEO, tuple(INDIA_STATES)),
    # PEM & AEM electrolyzers, India-HQ only, states only (user, 2026-10-14).
    "h2": ("run_h2_market.py", "h2_profile_", "logs_h2", MARKETS_H2, tuple(INDIA_STATES)),
    # Global Data Center Market, worldwide country by country (user, 2026-10-15).
    "dc": ("run_dc_market.py", "dc_profile_", "logs_dc", MARKETS_DC, DC_REGIONS),
    # Global Empty Capsules Market, distributors, worldwide (user, 2026-10-15).
    "caps": ("run_caps_market.py", "caps_profile_", "logs_caps", MARKETS_CAPS, CAPS_REGIONS),
    # MENA Magnesium Glycinate, distributors, 16 MENA countries (user, 2026-10-18).
    "mg": ("run_mg_market.py", "mg_profile_", "logs_mg", MARKETS_MG, MG_REGIONS),
    # U.S. Entresto, pharma wholesale distributors, state by state (user, 2026-10-19).
    "ent": ("run_ent_market.py", "ent_profile_", "logs_ent", MARKETS_ENT, ENT_REGIONS),
    # India Spices And Seasonings, distributors / wholesalers, India states only (user, 2026-10-19).
    "spice": ("run_spice_market.py", "spice_profile_", "logs_spice", MARKETS_SPICE, tuple(INDIA_STATES)),
    # Global Soft Serve Ice Cream Machines, distributors/importers/dealers, worldwide (user, 2026-10-21).
    "soft": ("run_soft_market.py", "soft_profile_", "logs_soft", MARKETS_SOFT, SOFT_REGIONS),
    # NA + Europe Vacuum Insulation Panels, distributors only (user, 2026-10-25).
    "vip": ("run_vip_market.py", "vip_profile_", "logs_vip", MARKETS_VIP, VIP_REGIONS),
    # U.S. 503B Compounding Pharmacies Market: 503B outsourcing facilities, U.S. HQ,
    # state by state (user, 2026-10-26).
    "cmp": ("run_cmp_market.py", "cmp_profile_", "logs_cmp", MARKETS_CMP, CMP_REGIONS),
    # Global Biopharma Buffer Market: DISTRIBUTORS, global (user, 2026-10-28).
    "buf": ("run_buf_market.py", "buf_profile_", "logs_buf", MARKETS_BUF, BUF_REGIONS),
    # Mozambique Fire Insurance Market: insurers + brokers/agents, Mozambique HQ (user, 2026-10-28).
    "moz": ("run_moz_market.py", "moz_profile_", "logs_moz", MARKETS_MOZ, MOZ_REGIONS),
    # Japan Fashion E-commerce Market: online fashion sellers, Japan HQ (user, 2026-10-29).
    "jpf": ("run_jpf_market.py", "jpf_profile_", "logs_jpf", MARKETS_JPF, JPF_REGIONS),
    # Global Electrosurgical and Allied Surgical Medical Devices Market: distributors / dealers, global (user, 2026-10-30).
    "esu": ("run_esu_market.py", "esu_profile_", "logs_esu", MARKETS_ESU, ESU_REGIONS),
    # 120 Adhesives & Bonding / Technical Foam markets: distributors / converters / dealers, global (user, 2026-10-30).
    "adh": ("run_adh_market.py", "adh_profile_", "logs_adh", MARKETS_ADH, ADH_REGIONS),
    # India Industrial Water & Wastewater Market: solution providers, India HQ only, states then cities (user, 2026-11-02).
    "iww": ("run_iww_market.py", "iww_profile_", "logs_iww", MARKETS_IWW, IWW_REGIONS),
    # The same 120 adhesives & foam markets, EUROPE only, seeded with the global finals' Europe companies (user, 2026-11-03).
    "eadh": ("run_eadh_market.py", "eadh_profile_", "logs_eadh", MARKETS_EADH, EADH_REGIONS),
    # India Luggage Delivery Service Market: luggage delivery service providers, India HQ,
    # nationwide round naming the mandatory companies, then states, then cities (user, 2026-11-05).
    "lug": ("run_lug_market.py", "lug_profile_", "logs_lug", MARKETS_LUG, LUG_REGIONS),
    # 200 CPHI India 2026 pharma markets: distributors / wholesalers (CDMOs for 191-200), global,
    # one market at a time with 20 browsers, every step per market (user, 2026-11-12).
    "cphi": ("run_cphi_market.py", "cphi_profile_", "logs_cphi", MARKETS_CPHI, CPHI_REGIONS),
    # Global Ethanolamine + Global Ethylene Oxide Derivatives: chemical distributors, global, one market
    # at a time with 10 browsers, every step per market (user, 2026-11-15).
    "eth": ("run_eth_market.py", "eth_profile_", "logs_eth", MARKETS_ETH, ETH_REGIONS),
    # Potassium Sorbate / Sorbic Acid / Caffeine, each Global + India (India HQ), distributors +
    # wholesalers, one market at a time with 10 browsers (user, 2026-11-16).
    "sorb": ("run_sorb_market.py", "sorb_profile_", "logs_sorb", MARKETS_SORB, SORB_REGIONS),
    # Brazil Industrial Electrification Market: solution providers (distributors, dealers,
    # integrators, EPC/O&M, makers), Brazil-based only, cap 400, 10 browsers (user, 2026-11-18).
    "bie": ("run_bie_market.py", "bie_profile_", "logs_bie", MARKETS_BIE, BIE_REGIONS),
    # India Alcohol Market: India-HQ liquor importers / distributors / channel partners, focus gin,
    # vodka, whisky (client export expansion), cap 400, 10 browsers (user, 2026-11-19).
    "alc": ("run_alc_market.py", "alc_profile_", "logs_alc", MARKETS_ALC, ALC_REGIONS),
    # The 200 CPHI pharma markets, INDIA ONLY (India HQ), cap 200 each, one market at a time
    # with 10 browsers, low-yield stop (user, 2026-11-19).
    "cphin": ("run_cphin_market.py", "cphin_profile_", "logs_cphin", MARKETS_CPHIN, CPHIN_REGIONS),
}
RUNNER, PROFILE_PREFIX = "", ""
LOGS = ROOT
FIRE_SLUGS: tuple[str, ...] = ()
REGIONS: tuple[str, ...] | dict[str, tuple[str, ...]] = ()  # shared list, or per-market plan


BATCH_NAME = "fire"


def select_batch(name: str) -> None:
    global RUNNER, PROFILE_PREFIX, LOGS, FIRE_SLUGS, REGIONS, BATCH_NAME
    BATCH_NAME = name
    RUNNER, PROFILE_PREFIX, logs, markets, REGIONS = _BATCHES[name]
    LOGS = ROOT / logs
    LOGS.mkdir(exist_ok=True)
    FIRE_SLUGS = tuple(slug for slug, *_ in markets)


select_batch("fire")

BATCH_SIZE = 20
MAX_ROUNDS_PER_WAVE = 200  # safety ceiling, not expected to be hit
_EXTENSION_MASTER = ROOT / "data" / "ai_mode_chrome_profile_extensions"


def log(msg: str) -> None:
    print(msg, flush=True)


def _slugify_for_match(slug: str) -> str:
    """The regex-safe fragment of the --file path verify_all_drones.py
    passes to each verifier, e.g. gas_flame_detection ->
    'gas-flame-detection-market'. Matching on the file path (derived
    the same way verify_all_drones.py derives it, via market_slug) is
    reliable where matching on --market-name is not: market names like
    "Gas & Flame Detection Market - India" contain regex-special
    characters ('&') that silently broke the match, which is what let
    a verifier be mistaken for already-drained while still running.
    """
    market = market_batches.by_slug(slug)[1]
    return re.escape(market_slug(market))


def stage_extension(slug: str, profile_suffix: str | None = None) -> None:
    """Give this market's Chromium profile its own captcha-raptor copy.

    Fire markets have never run before, so there is no existing
    fire_profile_<slug>_extensions sibling to copy from. Stage directly
    from the confirmed-working master copy instead. A suffixed worker
    (fire_profile_<slug>_w2, for a second parallel worker on the same
    market) needs its OWN staged copy too -- run_fire_market.py never
    stages extensions itself, only this launcher does, and the profile
    directory name (which is what needs a matching _extensions sibling)
    includes the suffix.
    """
    name = f"{PROFILE_PREFIX}{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    dst = ROOT / "data" / f"{name}_extensions"
    if dst.is_dir():
        return
    if _EXTENSION_MASTER.is_dir():
        shutil.copytree(_EXTENSION_MASTER, dst)


def _coverage_file(slug: str) -> Path:
    return ROOT / "data" / "region_coverage" / f"{PROFILE_PREFIX}{slug}.json"


def finished_regions(slug: str) -> set[str]:
    """Regions whose search FINISHED, recorded by the launcher itself.

    Several browsers of one market write the same log file at once, and on
    Windows they overwrite each other's lines, so the "coverage:" lines in
    that log lose most of the regions (Brazil run: 3 of 15 survived, and
    round 2 re-searched four regions instead of moving on). This file has one
    writer (the market's own loop), so nothing is lost."""
    try:
        return set(json.loads(_coverage_file(slug).read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return set()


def mark_finished(slug: str, regions: list[str]) -> None:
    done = finished_regions(slug) | {r.split("(")[0].strip() for r in regions}
    f = _coverage_file(slug)
    f.parent.mkdir(parents=True, exist_ok=True)
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(sorted(done), ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, f)


def covered_regions(slug: str) -> set[str]:
    used = finished_regions(slug)
    log_path = LOGS / f"{slug}.log"
    if not log_path.is_file():
        return used
    try:
        text = log_path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return used
    return used | {m.split("(")[0].strip()
                   for m in re.findall(r"coverage: ([^\n]+)", text)}


def next_region(slug: str) -> str | None:
    found = next_regions(slug, 1)
    return found[0] if found else None


def next_regions(slug: str, n: int) -> list[str]:
    """Up to N distinct unswept regions for this market, for running
    multiple workers on the same market at once. Reading covered_regions
    once and picking N from it (rather than calling next_region() N
    times) matters here: next_region() only knows a region is taken once
    its own "coverage: ..." log line has been written, which happens
    AFTER that region's worker finishes -- two workers started together
    would both read the same empty log and pick the identical region."""
    used = covered_regions(slug)
    out: list[str] = []
    # A batch may give each market its own search plan ({slug: regions},
    # e.g. geo: worldwide + that market's focus countries) instead of one
    # shared country list.
    pool = REGIONS.get(slug, ()) if isinstance(REGIONS, dict) else REGIONS
    for region in pool:
        if len(out) >= n:
            break
        # "Karnataka (Bengaluru, Mysuru)" is logged and matched as "Karnataka".
        if region.split("(")[0].strip() not in used and region not in out:
            out.append(region)
    return out


def discover_one_region(slug: str, region: str, goal_padding: int,
                        profile_suffix: str | None = None) -> int:
    """ONE discovery call for this market's current region. Returns
    companies gained. Deliberately does not loop internally -- the wave
    controls rotation so it can interleave a verify pass between rounds
    instead of a worker holding its slot until its region list runs out.

    profile_suffix lets a SECOND (or Nth) worker cover this same market
    at once, each in its own isolated Chromium profile via
    run_fire_market.py's own --profile-suffix (never shared -- a shared
    profile produces about:blank tabs). Both workers write to the same
    on-disk state file for this market; that is safe because state.py's
    save() merges by dedupe key under a file lock rather than
    overwriting, which was built and stress-tested specifically for
    this kind of concurrent same-market write.
    """
    market = market_batches.by_slug(slug)[1]
    before = S.discovered(market)
    goal = before + goal_padding

    # Always the SAME log file regardless of profile_suffix: coverage_
    # regions() reads this file to know which regions are already swept,
    # and that has to stay one shared ledger across every worker on this
    # market or two suffixed workers would each think the other's region
    # was still untried.
    log_path = LOGS / f"{slug}.log"
    handle = log_path.open("a", encoding="utf-8", errors="replace")
    cmd = [
        sys.executable, "-u", str(ROOT / RUNNER),
        "--slug", slug,
        "--target", str(goal),
        "--batch-size", str(BATCH_SIZE),
        "--region", region,
    ]
    if profile_suffix:
        cmd += ["--profile-suffix", profile_suffix]
    try:
        subprocess.run(cmd, cwd=str(ROOT), stdout=handle,
                       stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                       timeout=1800)
    except subprocess.TimeoutExpired:
        pass
    except Exception as exc:  # noqa: BLE001
        handle.write(f"\n[launcher] worker error: {exc}\n")
    finally:
        handle.close()

    return S.discovered(market) - before


def _count_matching(pattern: str) -> int:
    ps = (
        "(Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | Where-Object { $_.CommandLine "
        f"-match '{pattern}' }} | Measure-Object).Count"
    )
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True,
                             timeout=60).stdout.strip()
        return int(out or 0)
    except (ValueError, subprocess.SubprocessError):
        return -1  # unknown -- caller should not assume drained


def stop_browsers_for(slugs: list[str]) -> None:
    """Kill any Chromium still open for these slugs, so verification's
    OWN launch of the same profiles never collides with a discovery
    process that has not exited yet. Polls until the kill is actually
    confirmed rather than a fixed sleep -- a Stop-Process -Force can
    still take longer than a few seconds to fully release a Chromium
    profile lock, and starting a verifier before that happens is what
    produces about:blank tabs sharing the profile with a live worker."""
    if not slugs:
        return
    pattern = "|".join(slugs)
    kill_ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | Where-Object { "
        f"$_.CommandLine -match '{Path(RUNNER).stem}.*--slug ({pattern})\\b' "
        "} | ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
        "-ErrorAction SilentlyContinue }"
    )
    match_pattern = f"{Path(RUNNER).stem}.*--slug ({pattern})\\b"
    for attempt in range(10):
        try:
            subprocess.run(["powershell", "-NoProfile", "-Command", kill_ps],
                           capture_output=True, timeout=60)
        except subprocess.SubprocessError:
            pass
        time.sleep(2)
        remaining = _count_matching(match_pattern)
        if remaining == 0:
            break
    else:
        log(f"  !! stop_browsers_for: {slugs} still alive after "
            f"{attempt + 1} kill attempts")
    time.sleep(2)  # let the OS finish releasing the profile lock


def export_and_verify(slugs: list[str]) -> None:
    """Export the wave's markets, then verify them -- each in its own
    profile, so this can run in parallel across the wave the same way
    discovery does.
    """
    log(f"  -- exporting + verifying {slugs}")
    subprocess.run([sys.executable, str(ROOT / "market_batches.py")],
                   cwd=str(ROOT), capture_output=True, timeout=900)

    cmd = [sys.executable, str(ROOT / "verify_all_drones.py")]
    for slug in slugs:
        cmd += ["--only", slug]
    subprocess.run(cmd, cwd=str(ROOT), capture_output=True, timeout=300)

    # verify_all_drones.py launches detached children (one Popen per
    # slug, each spawned by its own subprocess.Popen call, which is not
    # instant); wait for the verifiers scoped to THIS wave's slugs to
    # both APPEAR and then drain before resuming discovery on the same
    # profiles. Checking only for "0 running" without first confirming
    # they ever started was a false-positive trap: a slow-to-launch
    # child (or one whose command line didn't match the regex yet on
    # the very first poll) reads as "already drained", and the next
    # round's discovery then launches straight into a profile a
    # verifier is still using -- the about:blank collision.
    match_pattern = "verify_distributors.*(" + \
        "|".join(_slugify_for_match(s) for s in slugs) + ")"
    seen = False
    waited = 0
    while waited < 1800:
        remaining = _count_matching(match_pattern)
        if remaining > 0:
            seen = True
        elif seen:
            break
        elif waited >= 60:
            # never saw them appear after a full minute -- something
            # about the launch or the match pattern is wrong; stop
            # waiting on a check that may never turn positive.
            log(f"  !! export_and_verify: verifiers for {slugs} never "
                f"observed running after {waited}s -- proceeding anyway")
            break
        time.sleep(10)
        waited += 10
    log(f"  -- verification drained after {waited}s (seen={seen})")


def run_wave(slugs: list[str], target: int, workers_per_market: int = 1) -> None:
    """Run this wave of markets to completion: alternate one discovery
    round per market with a verify pass, until every market in the wave
    hits target or has no region left to search.

    workers_per_market > 1 runs that many CONCURRENT discovery workers
    on the SAME market, each on a different unswept region and its own
    isolated Chromium profile (via --profile-suffix w1, w2, ...). Useful
    for a small wave (e.g. wave 3's 5 markets) where running only 5
    browsers leaves capacity unused; this fills it by covering more
    ground per round instead of leaving 5 of the "10 browsers" idle.
    """
    for slug in slugs:
        stage_extension(slug)
        for i in range(1, workers_per_market):
            stage_extension(slug, profile_suffix=f"w{i + 1}")

    for round_num in range(1, MAX_ROUNDS_PER_WAVE + 1):
        market_of = market_batches.by_slug
        active = [s for s in slugs if S.verified(market_of(s)[1]) < target]
        active = [s for s in active if next_region(s) is not None]
        if not active:
            break

        log(f"  round {round_num}: discovering {active} "
            f"({workers_per_market} worker(s)/market)")
        threads = []
        for slug in active:
            regions = next_regions(slug, workers_per_market)
            for i, region in enumerate(regions):
                # worker 0 uses the plain (unsuffixed) profile -- the
                # one every single-worker round already used, with its
                # extension already staged by the plain stage_extension
                # call above. Only workers 1+ get "w2", "w3", ... to
                # match the suffixes stage_extension() staged for them.
                suffix = f"w{i + 1}" if i > 0 else None
                t = threading.Thread(
                    target=lambda s=slug, r=region, sf=suffix:
                        discover_one_region(s, r, 60, profile_suffix=sf),
                    daemon=True,
                )
                threads.append(t)
                t.start()
                time.sleep(2)  # stagger Chromium launches
        for t in threads:
            t.join()

        stop_browsers_for(active)
        export_and_verify(active)

    for slug in slugs:
        market = market_batches.by_slug(slug)[1]
        v = S.verified(market)
        reason = "target reached" if v >= target else "region search exhausted"
        log(f"  {slug:38} verified={v:>4}/{target}  ({reason})")


# -- rolling mode: N slots, a finished market hands its slot to the next ------

# A verifier still running after this is stuck. VERIFY_TIMEOUT_SEC raises it
# for throttled runs (eth 2026-11-15: two 40-min cutoffs left 324 kept
# companies unmerged and started a needless extra search round).
VERIFY_TIMEOUT = int(os.environ.get("VERIFY_TIMEOUT_SEC") or 2400)
WORKERS_PER_MARKET = 1         # browsers per market per round (set by main)
ZERO_ROUNDS_EXHAUSTED = 12     # consecutive regions with nothing new = exhausted
_print_lock = threading.Lock()


def rlog(msg: str) -> None:
    with _print_lock:
        log(f"[{time.strftime('%H:%M:%S')}] {msg}")


def _kill(ps_filter: str) -> None:
    ps = ("Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | "
          f"Where-Object {{ {ps_filter} }} | ForEach-Object {{ "
          "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, timeout=60)
    except subprocess.SubprocessError:
        pass


def reset_profile(slug: str, reason: str) -> None:
    """Close this market's browsers and archive its profile (cookies and
    all), so the next round starts on a brand-new identity. Archived, not
    deleted, like every other profile reset in this project."""
    base = f"{PROFILE_PREFIX}{slug}"
    # The market's own profile plus its parallel workers' (_w2, _w3, ...),
    # each with or without the _chromium suffix.
    own = re.compile(rf"^{re.escape(base)}(_w\d+)?(_chromium)?$")
    _kill(f"$_.Name -match '^(chrome|chromium)\\.exe$' -and "
          f"$_.CommandLine -match '{re.escape(base)}(_w\\d+)?(_chromium)?([\\s\"]|$)'")
    time.sleep(3)
    archive = ROOT / "data" / "_archived_profiles_rolling"
    archive.mkdir(exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    moved = 0
    for p in sorted((ROOT / "data").iterdir()):
        if p.is_dir() and own.match(p.name):
            try:
                p.rename(archive / f"{p.name}_{stamp}")
                moved += 1
            except OSError:
                shutil.rmtree(p, ignore_errors=True)
    rlog(f"  {slug}: profile reset ({reason}), {moved} folder(s) archived")
    _prune_archive(archive)


ARCHIVE_KEEP = 10  # newest archived profiles kept; older ones deleted


def _prune_archive(archive: Path) -> None:
    """Archived profiles are only kept for a look at what went wrong. Unbounded,
    they filled the C: drive (adh 2026-10-31: 11 GB, 'No space left on device'
    killed the batch), so only the newest ARCHIVE_KEEP stay (user OK'd deleting
    archived profiles)."""
    try:
        dirs = sorted((p for p in archive.iterdir() if p.is_dir()),
                      key=lambda p: p.stat().st_mtime, reverse=True)
    except OSError:
        return
    for old in dirs[ARCHIVE_KEEP:]:
        shutil.rmtree(old, ignore_errors=True)


def verify_one(slug: str) -> None:
    """Export and verify ONE market, waiting for its verifier to finish.
    A verifier still running after VERIFY_TIMEOUT is killed and its profile
    reset: it is stuck on a block page, and the next round must not share
    the profile with it."""
    subprocess.run([sys.executable, str(ROOT / "market_batches.py"), "--only", slug],
                   cwd=str(ROOT), capture_output=True, timeout=900)
    # PARALLEL_VERIFY=N: verify with N browsers at once (scripts/parallel_verify.py
    # splits the unjudged rows, merges the verdicts into this market's own
    # checkpoint and writes the same verified workbook). One browser takes
    # 35-45 min for a 300-company round; 10 take ~5-8 (user, 2026-10-19).
    # Default: as many verifiers as this market has search browsers.
    parts = int(os.environ.get("PARALLEL_VERIFY") or WORKERS_PER_MARKET or 0)
    if parts > 1:
        market = market_batches.by_slug(slug)[1]
        log_path = LOGS / f"{slug}_parallel_verify.log"
        # Up to 2 attempts: a run that ends WITHOUT its final "done" line was
        # cut short (spice round 2, 2026-10-20: killed mid-run, nothing
        # merged). parallel_verify resumes from the part checkpoints, so the
        # retry only judges what is left and then merges everything.
        for attempt in (1, 2):
            start = log_path.stat().st_size if log_path.exists() else 0
            try:
                with log_path.open("a", encoding="utf-8") as fh:
                    subprocess.run([sys.executable, "-u", str(ROOT / "scripts" / "parallel_verify.py"),
                                    "--market", market, "--parts", str(parts)],
                                   cwd=str(ROOT), stdout=fh, stderr=subprocess.STDOUT,
                                   timeout=VERIFY_TIMEOUT)
            except subprocess.TimeoutExpired:
                # Its checkpoints are kept; the next round's verify resumes them.
                rlog(f"  {slug}: parallel verify still running after {VERIFY_TIMEOUT // 60} min — "
                     "stopped, resumes next round")
                return
            with log_path.open("r", encoding="utf-8", errors="replace") as fh:
                fh.seek(start)
                finished = "] done" in fh.read()
            if finished:
                return
            rlog(f"  {slug}: parallel verify ended without merging (attempt {attempt}) — "
                 + ("retrying" if attempt == 1 else "will merge next round"))
        return
    pattern = "verify_distributors.*" + _slugify_for_match(slug)
    # Never two verifiers on one market (same checkpoint + profile): a
    # leftover from an earlier round is killed first (adh 2026-10-31: three
    # markets had an orphan 50-60 min verifier next to a new one).
    _kill(f"$_.Name -eq 'python.exe' -and $_.CommandLine -match '{pattern}'")
    try:
        subprocess.run([sys.executable, str(ROOT / "verify_all_drones.py"), "--only", slug],
                       cwd=str(ROOT), capture_output=True, timeout=300)
    except subprocess.TimeoutExpired:
        # Slow start on a busy PC (adh065 2026-10-31): the verifier is
        # usually up anyway -- go on and wait for it below, don't crash the
        # market loop into a retry.
        rlog(f"  {slug}: verifier launch slow (>300 s) -- waiting for it")
    seen, waited = False, 0
    while waited < VERIFY_TIMEOUT:
        n = _count_matching(pattern)
        if n > 0:
            seen = True
        elif n < 0:
            pass  # count timed out (busy PC): unknown, NOT drained -- keep waiting
        elif seen or waited >= 90:
            return
        time.sleep(15)
        waited += 15
    _kill(f"$_.Name -eq 'python.exe' -and $_.CommandLine -match '{pattern}'")
    reset_profile(slug, f"verifier stuck {VERIFY_TIMEOUT // 60} min")


# A region list may start with "Worldwide round N" entries (global searches,
# no country named) followed by countries. Global searches keep returning the
# same well-known names, so once a worldwide round adds fewer than this many
# new companies the market's remaining worldwide rounds are marked done and
# the next rounds take the countries it has not searched yet (user, adh
# 2026-10-31: "okay do").
WORLDWIDE_MIN_GAIN = 10

# Optional low-yield stop (env; 0 = off, the default for every older batch).
LOW_YIELD_MIN = int(os.environ.get("LOW_YIELD_MIN") or 0)
LOW_YIELD_ROUNDS = int(os.environ.get("LOW_YIELD_ROUNDS") or 2)


def _is_worldwide(region: str) -> bool:
    return region.lower().startswith("worldwide")


def switch_to_countries(slug: str, gained: int) -> None:
    pool = REGIONS.get(slug, ()) if isinstance(REGIONS, dict) else REGIONS
    rest = [r for r in pool if _is_worldwide(r)]
    if not any(not _is_worldwide(r) for r in pool):
        return  # worldwide-only plan: nothing to switch to
    mark_finished(slug, rest)
    rlog(f"  {slug}: worldwide round added {gained} (< {WORLDWIDE_MIN_GAIN}) "
         f"-> switching to countries not searched yet")


def market_loop(slug: str, target: int, start_delay: float) -> None:
    """Drive one market to target (or exhaustion) on its own: discover one
    region, verify, repeat. A round that finds nothing new resets the
    profile, since that is what a browser stuck on CAPTCHA / rate-limit
    pages looks like from outside."""
    time.sleep(start_delay)
    market = market_batches.by_slug(slug)[1]
    # Per-market cap: a market module may define TARGETS = {slug: cap}
    # (sorb, user 2026-11-16: India markets 200, Global markets 400).
    own = market_batches._module_attr("TARGETS", slug)
    if own:
        target = int(own)
    zero_rounds = 0
    low_rounds = 0
    rlog(f"START {slug} (verified {S.verified(market)}/{target}, "
         f"{WORKERS_PER_MARKET} browser(s))")
    while True:
        v = S.verified(market)
        if v >= target:
            rlog(f"DONE {slug}: target reached ({v})")
            return
        regions = next_regions(slug, WORKERS_PER_MARKET)
        if not regions:
            rlog(f"DONE {slug}: every region searched ({v} verified)")
            return
        before = S.discovered(market)
        threads = []
        lasted: dict[str, float] = {}   # region -> seconds its worker ran
        for i, region in enumerate(regions):
            # Worker 0 keeps the plain profile; others get _w2, _w3, ...
            suffix = f"w{i + 1}" if i else None
            stage_extension(slug, suffix)
            # The goal is market total + padding, and the total is SHARED by
            # every worker on this market: a flat +60 was reached within
            # minutes by 10-20 workers together, each region got ~3 companies
            # and was then marked searched (dc log: "Reached 193/60"). Scale
            # it so each region runs until it is genuinely exhausted.
            def _search(region=region, suffix=suffix, pad=60 * len(regions)):
                t0 = time.time()
                discover_one_region(slug, region, pad, profile_suffix=suffix)
                lasted[region] = time.time() - t0

            t = threading.Thread(target=_search, daemon=True)
            threads.append(t)
            t.start()
            time.sleep(3)  # stagger Chromium launches
        for t in threads:
            t.join()
        # A worker that died in under 90 s (browser failed to start) did not
        # search its region, so it is not marked and will be retried.
        mark_finished(slug, [r for r in regions if lasted.get(r, 0) >= 90])
        gained = S.discovered(market) - before
        stop_browsers_for([slug])
        verify_one(slug)
        v = S.verified(market)
        names = ", ".join(r.split("(")[0].strip() for r in regions)
        rlog(f"  {slug}: {names} +{gained} found, verified {v}/{target}")
        if all(_is_worldwide(r) for r in regions) and gained < WORLDWIDE_MIN_GAIN:
            switch_to_countries(slug, gained)
        # Low-yield stop (cphin, user 2026-11-19: 200 India markets): when
        # LOW_YIELD_MIN is set, a market whose rounds add fewer than that
        # many companies LOW_YIELD_ROUNDS times in a row is treated as
        # exhausted, instead of searching every remaining city for 1-2 more.
        if LOW_YIELD_MIN and gained < LOW_YIELD_MIN:
            low_rounds += 1
            if low_rounds >= LOW_YIELD_ROUNDS:
                rlog(f"DONE {slug}: {low_rounds} rounds in a row added < {LOW_YIELD_MIN} "
                     f"new companies ({v} verified)")
                return
        else:
            low_rounds = 0
        if gained > 0:
            zero_rounds = 0
            continue
        zero_rounds += 1
        reset_profile(slug, "round found nothing new")
        if zero_rounds >= ZERO_ROUNDS_EXHAUSTED:
            rlog(f"DONE {slug}: {zero_rounds} regions in a row found nothing new ({v} verified)")
            return


def finalize(batch: str, slugs: list[str], workers: int) -> Path:
    """Last pipeline step: active-status + ownership check on every verified
    company, via scripts/status_check_folder.py. Closed companies (with
    evidence) are removed, acquired ones renamed "ABC (acquired by XYZ)";
    output is the checked xlsx + a JSON per market and a report (no zip).

    The verified files are copied into output/final/<batch>/<batch>_verified
    under slug names first: the checker reads a folder, and the folder name
    keys its checkpoints, so each batch gets its own and re-running resumes.
    """
    final = ROOT / "output" / "final" / batch
    stage = final / f"{batch}_verified"
    stage.mkdir(parents=True, exist_ok=True)
    copied = 0
    for slug in slugs:
        market = market_batches.by_slug(slug)[1]
        src = ROOT / "output" / "market_research" / f"{market_slug(market)}_verified_distributors.xlsx"
        if src.is_file():
            shutil.copy2(src, stage / f"{slug}_verified_distributors.xlsx")
            copied += 1
    out = final / f"{batch} distributors (status checked)"
    rlog(f"FINALIZE {batch}: status + ownership check on {copied} market file(s) -> {out}")
    env = dict(os.environ, STATUS_WALL_WAITS=os.environ.get("STATUS_WALL_WAITS", "6"))
    # No --zip: the user wants the Excel and JSON files only (2026-10-13).
    subprocess.run([sys.executable, "-u", str(ROOT / "scripts" / "status_check_folder.py"),
                    "--input", str(stage), "--output", str(out),
                    "--workers", str(workers)], cwd=str(ROOT), env=env)
    report = out / "status_check_report.xlsx"
    # The checker builds output only once EVERY company is checked, and writes
    # the report last; a run cut short by Google blocks leaves no fresh report,
    # and re-running resumes it.
    if report.is_file() and report.stat().st_mtime > time.time() - 600:
        rlog(f"FINALIZE {batch}: done -> {out}")
        # Integrity audit of the delivered files (duplicates, JSON == Excel,
        # row-2 definition, region, flagged leaks, no zip): the checks that
        # were done by hand until 2026-10-21. Problems are logged, not fatal.
        audit = subprocess.run([sys.executable, str(ROOT / "scripts" / "audit_final.py"),
                                "--batch", batch], cwd=str(ROOT), capture_output=True,
                               text=True, encoding="utf-8", errors="replace")
        for line in (audit.stdout or "").splitlines():
            rlog(f"  {line}")
    else:
        rlog(f"FINALIZE {batch}: status check incomplete (blocked) — re-run to resume")
    return out


def run_rolling(slugs: list[str], target: int, slots: int) -> None:
    from concurrent.futures import ThreadPoolExecutor

    todo = [s for s in slugs if S.verified(market_batches.by_slug(s)[1]) < target]
    rlog(f"rolling: {len(todo)} market(s) to run, {slots} at a time, target {target}")
    with ThreadPoolExecutor(max_workers=slots) as pool:
        # Only the first `slots` markets start together; stagger their
        # browser launches. Later markets start the moment a slot frees up.
        for i, slug in enumerate(todo):
            pool.submit(_safe_market_loop, slug, target, 5.0 * i if i < slots else 0.0)


def _safe_market_loop(slug: str, target: int, start_delay: float) -> None:
    """A pool thread's exception is otherwise silent. Log it and retry the
    market (its saved work is kept) up to 3 times before giving the slot up."""
    import traceback

    for attempt in range(1, 4):
        try:
            market_loop(slug, target, start_delay if attempt == 1 else 30.0)
            if os.environ.get("FINALIZE_EACH") == "1":
                finalize_one(slug)
            return
        except Exception:  # noqa: BLE001
            rlog(f"ERROR {slug} (attempt {attempt}/3):\n{traceback.format_exc()}")
    rlog(f"GAVE UP {slug} after 3 errors")


STATUS_WORKERS = 20


def _finished_slugs(upto: str) -> list[str]:
    """Markets of this batch up to and including `upto` (they run in order,
    one at a time, under FINALIZE_EACH) whose verified workbook exists."""
    out = []
    for s in FIRE_SLUGS:
        market = market_batches.by_slug(s)[1]
        if (ROOT / "output" / "market_research" / f"{market_slug(market)}_verified_distributors.xlsx").is_file():
            out.append(s)
        if s == upto:
            break
    return out


def finalize_one(slug: str) -> None:
    """FINALIZE_EACH=1 (cphi, user 2026-11-12: "after one market is complete
    with all pipeline steps, run the next"): right after a market finishes
    discovery + verification, run the status + ownership check and the build
    for every market finished so far (only the new companies are asked;
    earlier answers are reused), then delete this market's browser profiles
    so 200 markets x 20 profiles never fill the disk -- the next market
    starts on fresh ones."""
    # A resumed launcher walks every market again; one whose final file is
    # already newer than its verified workbook needs no status pass (each
    # empty pass cost ~5 min of staggered workers -- 2026-11-16 resume spent
    # 30 min re-finalizing markets 1 and 5-8 with no browser open).
    market = market_batches.by_slug(slug)[1]
    verified = ROOT / "output" / "market_research" / f"{market_slug(market)}_verified_distributors.xlsx"
    final = (ROOT / "output" / "final" / BATCH_NAME / f"{BATCH_NAME} distributors (status checked)"
             / f"{slug}_verified_distributors.xlsx")
    if final.is_file() and verified.is_file() and final.stat().st_mtime >= verified.stat().st_mtime:
        rlog(f"FINALIZE-EACH {slug}: final file already up to date - skipped")
        return
    rlog(f"FINALIZE-EACH {slug}: status check + build")
    finalize(BATCH_NAME, _finished_slugs(slug), STATUS_WORKERS)
    stop_browsers_for([slug])
    base = f"{PROFILE_PREFIX}{slug}"
    own = re.compile(rf"^{re.escape(base)}(_w\d+)?(_chromium|_extensions)?$")
    freed = 0
    for p in (ROOT / "data").iterdir():
        if p.is_dir() and own.match(p.name):
            shutil.rmtree(p, ignore_errors=True)
            freed += 1
    rlog(f"FINALIZE-EACH {slug}: done, {freed} profile folder(s) deleted")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--wave", type=int, default=10,
                    help="how many markets run per wave")
    ap.add_argument("--skip-waves", type=int, default=0,
                    help="number of leading waves to skip (already done)")
    ap.add_argument("--workers-per-market", type=int, default=1,
                    help="concurrent discovery workers per market, for "
                    "using spare browser capacity on a small wave")
    ap.add_argument("--batch", choices=sorted(_BATCHES), default="fire",
                    help="which India batch to run (default fire)")
    ap.add_argument("--rolling", action="store_true",
                    help="keep --wave markets running at once; a finished "
                    "market's slot starts the next one immediately")
    ap.add_argument("--finalize-only", action="store_true",
                    help="skip discovery; run only the final status + "
                    "ownership check and packaging")
    ap.add_argument("--status-workers", type=int, default=10,
                    help="browsers for the final status + ownership check")
    args = ap.parse_args()
    select_batch(args.batch)
    global STATUS_WORKERS
    STATUS_WORKERS = args.status_workers

    by_slug = market_batches.by_slug
    if not args.finalize_only:
        # Refuse to launch a batch whose discovery prompt cannot fit Google's
        # URL: it would be cut from the end, answer in prose and find nothing
        # (vip 2026-10-25, 0 companies in 10 minutes). Measured, not guessed.
        pre = subprocess.run([sys.executable, str(ROOT / "scripts" / "preflight_prompt.py"),
                              "--batch", args.batch], cwd=str(ROOT), capture_output=True,
                             text=True, encoding="utf-8", errors="replace")
        for line in (pre.stdout or "").splitlines():
            rlog(line)
        if pre.returncode != 0:
            rlog(f"PREFLIGHT FAILED for {args.batch}: discovery prompt too long for the "
                 "Google URL -- shorten the scope; not launching.")
            return
    if args.finalize_only:
        finalize(args.batch, list(FIRE_SLUGS), args.status_workers)
        return
    if args.rolling:
        global WORKERS_PER_MARKET
        WORKERS_PER_MARKET = max(1, args.workers_per_market)
        run_rolling(list(FIRE_SLUGS), args.target, args.wave)
        log("")
        log(f"all {len(FIRE_SLUGS)} {args.batch} markets processed:")
        for slug in FIRE_SLUGS:
            log(f"  {slug:38} verified={S.verified(by_slug(slug)[1]):>4}")
        finalize(args.batch, list(FIRE_SLUGS), args.status_workers)
        return
    remaining = list(FIRE_SLUGS)
    if args.skip_waves:
        skip_count = args.skip_waves * args.wave
        skipped, remaining = remaining[:skip_count], remaining[skip_count:]
        log(f"skipping {len(skipped)} already-done market(s): {skipped}")

    while remaining:
        wave = remaining[: args.wave]
        remaining = remaining[args.wave:]
        log(f"===== wave: {wave} =====")
        run_wave(wave, args.target, workers_per_market=args.workers_per_market)

    log("")
    log(f"all {len(FIRE_SLUGS)} {args.batch} markets processed:")
    for slug in FIRE_SLUGS:
        market = by_slug(slug)[1]
        log(f"  {slug:38} verified={S.verified(market):>4}")


if __name__ == "__main__":
    main()
