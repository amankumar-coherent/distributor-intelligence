"""Per-market persisted state: master company list + dedup + batch numbering.

Backs the "give me more 20" continuation — each market gets one JSON state
file under data/market_research/, keyed by a slug of the MARKET string. State
tracks every company already returned (by dedupe key, so aliases/renamed
companies collapse the same way the rest of the codebase already handles
duplicates) plus the last row number assigned, so a new batch continues
numbering and excludes everything already seen.

Reuses vendor_intel.utils.domains.company_dedupe_key rather than
reimplementing name-normalization dedup logic.
"""
from __future__ import annotations

import json
import re
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from vendor_intel.utils.domains import company_dedupe_key

STATE_DIR = Path(__file__).resolve().parents[4] / "data" / "market_research"


@contextmanager
def _file_lock(path: Path, timeout: float = 60.0):
    """A simple cross-process mutex via exclusive file creation.

    Needed because save() now does read-merge-write: without this, two
    processes racing the "read the current file, merge, write" sequence
    could both read the same pre-merge state and one write would still
    clobber the other's. os.O_CREAT | os.O_EXCL is atomic on both
    Windows and POSIX -- exactly one process can create the lock file at
    a time, so this closes that race without adding a dependency (the
    machine already has multiple concurrent discovery workers, so this
    needed to work without portalocker/filelock being installed).
    """
    import os

    lock_path = path.with_suffix(path.suffix + ".lock")
    deadline = time.monotonic() + timeout
    fd = None
    while True:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_RDWR)
            break
        except (FileExistsError, PermissionError):
            # Windows raises PermissionError (WinError 5 / errno 13) for
            # O_CREAT|O_EXCL against an existing file, not FileExistsError
            # -- verified live: catching only FileExistsError left this
            # uncaught, which killed the worker outright instead of
            # waiting for the lock. POSIX raises FileExistsError; both are
            # "someone else holds the lock", so both mean "keep waiting".
            if time.monotonic() > deadline:
                # A stale lock from a crashed process must not wedge every
                # future save forever -- break it and proceed rather than
                # hang the whole discovery run. Reset the deadline after
                # breaking it, so a genuinely busy lock (another live
                # worker, not a crashed one) gets a full fresh timeout to
                # acquire instead of spinning tight forever once the
                # first deadline has passed.
                try:
                    lock_path.unlink()
                except OSError:
                    pass
                deadline = time.monotonic() + timeout
                continue
            time.sleep(0.05)
    try:
        yield
    finally:
        if fd is not None:
            os.close(fd)
        try:
            lock_path.unlink()
        except OSError:
            pass


def market_slug(market: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", market.strip().lower()).strip("-")
    return slug or "market"


def state_path(market: str) -> Path:
    return STATE_DIR / f"{market_slug(market)}.json"



def website_identity(row: dict) -> str:
    """Identity key from website + country, or "" when unavailable.

    Catches the same company listed under name variants, which
    company_dedupe_key cannot: it normalises punctuation and legal
    suffixes, but "PowernSun" and "Power & Sun" are simply different
    strings. A shared primary domain is much stronger evidence.

    Country is part of the key on purpose. A multinational runs one
    domain across distinct national distributors -- Krannich Solar has
    five, each with its own geography and contacts -- and collapsing
    those would destroy real companies. Same domain AND same country is
    the combination that means "one company written twice".
    """
    website = str(row.get("Website") or "").strip().lower()
    if not website:
        return ""
    domain = re.sub(r"^https?://", "", website)
    domain = re.sub(r"^www\.", "", domain).split("/")[0].split("?")[0].strip()
    if not domain:
        return ""

    hq = str(row.get("Headquarters") or "")
    country = hq.split(",")[-1] if "," in hq else hq
    return f"{domain}|{normalise_country(country)}"


# Country spellings that must collapse to one key. Measured in the live
# data: "UAE" (12 rows) and "United Arab Emirates" (65), "USA" (5) and
# "United States" (26), "UK" (2) and "United Kingdom" (13). Without this
# a same-domain duplicate written with the other spelling escapes the
# check entirely, which is the failure it exists to prevent.
_COUNTRY_ALIASES = {
    "uae": "unitedarabemirates",
    "unitedarabemirate": "unitedarabemirates",
    "usa": "unitedstates",
    "us": "unitedstates",
    "unitedstatesofamerica": "unitedstates",
    "america": "unitedstates",
    "uk": "unitedkingdom",
    "greatbritain": "unitedkingdom",
    "britain": "unitedkingdom",
    "england": "unitedkingdom",
    "scotland": "unitedkingdom",
    "wales": "unitedkingdom",
    "rok": "southkorea",
    # Bare "Korea" means the South in every commercial context; a
    # distributor writing "Seoul, Korea" must not be rejected by a gate
    # looking for "South Korea". North Korea is spelled out and maps to
    # its own key, so this cannot silently merge the two.
    "korea": "southkorea",
    "southkorea": "southkorea",
    "coreadelsur": "southkorea",
    "northkorea": "northkorea",
    "koreanorth": "northkorea",
    "democraticpeoplesrepublicofkorea": "northkorea",
    # Remaining APAC spellings seen in real headquarters strings.
    "hongkongsar": "hongkong",
    "hongkongsarchina": "hongkong",
    "chinahongkong": "hongkong",
    "macausar": "macau",
    "roc": "taiwan",
    "taiwanroc": "taiwan",
    "chinesetaipei": "taiwan",
    "burma": "myanmar",
    "republicoftheunionofmyanmar": "myanmar",
    "vietnamsocialistrepublic": "vietnam",
    "socialistrepublicofvietnam": "vietnam",
    "vietnam": "vietnam",
    "vietnam ": "vietnam",
    "bharat": "india",
    "republicofindia": "india",
    "ceylon": "srilanka",
    "republicofsingapore": "singapore",
    "kingdomofthailand": "thailand",
    "siam": "thailand",
    "philippinesrepublic": "philippines",
    "republicofthephilippines": "philippines",
    "aotearoa": "newzealand",
    "nz": "newzealand",
    "aus": "australia",
    "commonwealthofaustralia": "australia",
    "koreasouth": "southkorea",
    "republicofkorea": "southkorea",
    "prc": "china",
    "peoplesrepublicofchina": "china",
    "rsa": "southafrica",
    "holland": "netherlands",
    "thenetherlands": "netherlands",
    "czechia": "czechrepublic",
    "turkiye": "turkey",
    "tuerkiye": "turkey",
    "ksa": "saudiarabia",
    "kingdomofsaudiarabia": "saudiarabia",
    "vietnamsocialistrepublic": "vietnam",
    "ivorycoast": "cotedivoire",
}


def normalise_country(value: str) -> str:
    """Comparison key for a country name, collapsing common aliases."""
    key = re.sub(r"[^a-z]", "", str(value or "").lower())
    return _COUNTRY_ALIASES.get(key, key)


@dataclass
class MarketResearchState:
    market: str
    last_index: int = 0
    seen_keys: set[str] = field(default_factory=set)
    companies: list[dict[str, Any]] = field(default_factory=list)
    # Lazily derived from `companies` (see the website_keys property), so
    # it is never persisted and is always correct for a state file
    # written before this dedupe existed.
    _website_keys: set[str] | None = field(default=None, repr=False, compare=False)

    @classmethod
    def load(cls, market: str) -> "MarketResearchState":
        path = state_path(market)
        if not path.exists():
            return cls(market=market)
        # A concurrent worker's save() writes atomically (temp file +
        # os.replace), so a straight read here should never observe a
        # partial write -- but retry briefly anyway rather than crashing
        # the whole discovery loop on a transient race, since losing a
        # worker to an unhandled exception is far worse than a 50ms wait.
        last_exc: Exception | None = None
        for _attempt in range(5):
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                break
            except (json.JSONDecodeError, OSError) as exc:
                last_exc = exc
                time.sleep(0.05)
        else:
            raise RuntimeError(
                f"could not read state file {path} after 5 attempts"
            ) from last_exc
        return cls(
            market=data.get("market", market),
            last_index=int(data.get("last_index", 0)),
            seen_keys=set(data.get("seen_keys", [])),
            companies=list(data.get("companies", [])),
        )

    def save(self) -> None:
        """Persist this state, MERGING with whatever is on disk.

        A run_until_target loop calls .load() once, then holds `state`
        in memory across many batches, saving after each one. If a
        second worker is running the SAME market concurrently (multiple
        browsers split across Indian states/cities), that worker's
        earlier save() is invisible to this process's in-memory copy --
        an overwrite here would silently discard every company the other
        worker found in between. Companies are merged in by dedupe key
        before writing, under a short-lived file lock so two concurrent
        saves cannot both read the same pre-merge file and still race.
        """
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        path = state_path(self.market)

        with _file_lock(path):
            on_disk = self.companies
            on_disk_keys = set(self.seen_keys)
            on_disk_last_index = self.last_index
            if path.exists():
                try:
                    disk_data = json.loads(path.read_text(encoding="utf-8"))
                except (json.JSONDecodeError, OSError):
                    disk_data = None
                if disk_data:
                    disk_companies = disk_data.get("companies") or []
                    mine_keys = {
                        company_dedupe_key(c.get("Company", ""))
                        for c in self.companies if c.get("Company")
                    }
                    # Companies on disk that this process does NOT already
                    # have in memory -- almost always another worker's
                    # finds saved while this process was mid-batch.
                    merged = list(self.companies)
                    for c in disk_companies:
                        key = company_dedupe_key(c.get("Company", ""))
                        if key and key not in mine_keys:
                            merged.append(c)
                            mine_keys.add(key)
                    on_disk = merged
                    on_disk_keys = self.seen_keys | set(
                        disk_data.get("seen_keys") or [])
                    on_disk_last_index = max(
                        self.last_index, int(disk_data.get("last_index", 0)))

            payload = {
                "market": self.market,
                "last_index": on_disk_last_index,
                "seen_keys": sorted(on_disk_keys),
                "companies": on_disk,
            }
            import os
            tmp_path = path.with_suffix(path.suffix + f".tmp{os.getpid()}")
            tmp_path.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            # os.replace() is atomic, but on Windows it can raise a
            # transient PermissionError (WinError 5) if the destination
            # is momentarily held open by another process's read or by
            # antivirus scanning -- verified live under real concurrent
            # load. This is not a real conflict (the lock already
            # serializes writers); retry briefly rather than letting one
            # transient OS-level hiccup kill the whole discovery worker.
            for _attempt in range(10):
                try:
                    os.replace(str(tmp_path), str(path))
                    break
                except PermissionError:
                    time.sleep(0.05)
            else:
                os.replace(str(tmp_path), str(path))  # let it raise for real

        # Reflect the merge back into this process's own in-memory state,
        # so the NEXT save (and the running total this process logs)
        # accounts for what other workers have already found too.
        self.companies = on_disk
        self.seen_keys = on_disk_keys
        self.last_index = on_disk_last_index
        self._website_keys = None

    def is_seen(self, company_name: str) -> bool:
        return company_dedupe_key(company_name) in self.seen_keys

    def is_seen_row(self, row: dict) -> bool:
        """True when this row duplicates one already collected, by NAME or
        by website+country.

        The website check exists because name matching alone let the same
        company in up to seven times ("PowernSun", "Power & Sun",
        "Powernsun Group", ...). See website_identity for why country is
        part of the key.
        """
        if self.is_seen(str(row.get("Company") or "")):
            return True
        identity = website_identity(row)
        return bool(identity) and identity in self.website_keys

    @property
    def website_keys(self) -> set[str]:
        """website+country keys for everything collected so far.

        Derived rather than stored, so it is correct for state files
        written before this check existed -- of which there are several
        already on disk.
        """
        if self._website_keys is None:
            self._website_keys = {
                key for key in (website_identity(c) for c in self.companies) if key
            }
        return self._website_keys

    def excluded_names(self) -> list[str]:
        """Company names already returned, for the discovery-prompt exclusion list."""
        return [c.get("Company", "") for c in self.companies if c.get("Company")]

    def add_batch(self, rows: list[dict[str, str]]) -> None:
        for row in rows:
            name = row.get("Company", "")
            if not name:
                continue
            key = company_dedupe_key(name)
            if key in self.seen_keys:
                continue
            # Same company under a different name: 13-19% of collected
            # rows before this check.
            identity = website_identity(row)
            if identity and identity in self.website_keys:
                continue
            self.seen_keys.add(key)
            if identity:
                self.website_keys.add(identity)
            self.companies.append(row)
            idx = row.get("#")
            if idx and str(idx).isdigit():
                self.last_index = max(self.last_index, int(idx))
