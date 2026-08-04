"""Find Phase 1 / Phase 2 artifacts so a stopped market can skip early phases."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from vendor_intel.config import _project_root


def _norm_market_slug(market_slug: str) -> str:
    s = (market_slug or "").strip().lower()
    return re.sub(r"_market_global$", "_global", s)


def _clean_query(q: str) -> str:
    return re.sub(r"\s+", " ", (q or "").split("--")[0].strip()).lower()


def find_phase1_plan(
    *,
    query: str,
    market_slug: str | None = None,
) -> Path | None:
    """Return path to an existing phase1_plan_*.json for this market, if any."""
    plans_dir = _project_root() / "output" / "phase1"
    if not plans_dir.is_dir():
        return None

    candidates: list[Path] = []
    if market_slug:
        stem = _norm_market_slug(market_slug)
        exact = plans_dir / f"phase1_plan_{stem}.json"
        if exact.is_file():
            return exact
        # Truncation / punctuation variants
        for f in plans_dir.glob("phase1_plan_*.json"):
            fstem = f.stem.replace("phase1_plan_", "", 1)
            if fstem == stem or stem.startswith(fstem) or fstem.startswith(stem[:40]):
                candidates.append(f)

    q = _clean_query(query)
    if q:
        for f in plans_dir.glob("phase1_plan_*.json"):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
            except Exception:
                continue
            pq = _clean_query(str(data.get("query") or ""))
            if not pq:
                continue
            if pq == q or pq.startswith(q) or q.startswith(pq.replace(" market", "")):
                return f
            # keep fuzzy stem candidates for later
            if f not in candidates and market_slug:
                fstem = f.stem.replace("phase1_plan_", "", 1)
                stem = _norm_market_slug(market_slug or "")
                if stem and (fstem in stem or stem in fstem):
                    candidates.append(f)

    if candidates:
        return max(candidates, key=lambda p: p.stat().st_mtime)
    return None


def find_phase2_companies(
    *,
    query: str,
    market_slug: str | None = None,
    phase1_plan: Path | None = None,
) -> Path | None:
    """Return path to phase2_companies_*.json for this market, if any."""
    p2_dir = _project_root() / "output" / "phase2"
    if not p2_dir.is_dir():
        return None

    stems: list[str] = []
    if market_slug:
        stems.append(_norm_market_slug(market_slug))
    if phase1_plan is not None:
        stems.append(phase1_plan.stem.replace("phase1_plan_", "", 1))

    for stem in stems:
        if not stem:
            continue
        exact = p2_dir / f"phase2_companies_{stem}.json"
        if exact.is_file():
            return exact

    q = _clean_query(query)
    best: Path | None = None
    best_mtime = -1.0
    for f in p2_dir.glob("phase2_companies_*.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        pq = _clean_query(str(data.get("query") or ""))
        if q and pq and (pq == q or pq.startswith(q) or q.startswith(pq)):
            mt = f.stat().st_mtime
            if mt > best_mtime:
                best, best_mtime = f, mt
            continue
        fstem = f.stem.replace("phase2_companies_", "", 1)
        for stem in stems:
            if stem and (fstem == stem or stem.startswith(fstem) or fstem.startswith(stem[:40])):
                mt = f.stat().st_mtime
                if mt > best_mtime:
                    best, best_mtime = f, mt
                break
    return best


def load_phase2_companies(path: Path) -> list[dict[str, str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    companies = data.get("companies") or []
    out: list[dict[str, str]] = []
    for c in companies:
        if not isinstance(c, dict):
            continue
        name = str(c.get("name") or "").strip()
        domain = str(c.get("domain") or "").strip()
        if name and domain:
            row = {str(k): str(v) if v is not None else "" for k, v in c.items()}
            out.append(row)
    return out


def phase3_enrich_stem(*, query: str, market_slug: str | None = None) -> str:
    """Stable filename stem for Phase 3 enrich checkpoints."""
    if market_slug:
        return _norm_market_slug(market_slug)
    stem = re.sub(r"[^a-z0-9]+", "_", _clean_query(query)).strip("_")
    return (stem or "market")[:90]


def phase3_enrich_path(*, query: str, market_slug: str | None = None) -> Path:
    """Path: output/phase3/phase3_enrich_<stem>.json"""
    d = _project_root() / "output" / "phase3"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"phase3_enrich_{phase3_enrich_stem(query=query, market_slug=market_slug)}.json"


def load_phase3_checkpoint(path: Path | None) -> dict[str, Any] | None:
    """Load mid-enrich checkpoint (domain-fixed company list + enriched crawl map)."""
    if path is None or not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    return data


def save_phase3_checkpoint(
    path: Path,
    *,
    query: str = "",
    market_slug: str | None = None,
    companies: list[dict[str, Any]] | None = None,
    enriched: dict[str, Any] | None = None,
    merge: bool = True,
) -> None:
    """Persist Phase 3 progress. merge=True keeps prior enriched keys not in this write."""
    from datetime import datetime, timezone

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prev = load_phase3_checkpoint(path) if merge else None
    out: dict[str, Any] = {
        "phase": 3,
        "query": query or (prev or {}).get("query") or "",
        "market_slug": market_slug or (prev or {}).get("market_slug") or "",
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    if companies is not None:
        out["companies"] = companies
    elif prev and isinstance(prev.get("companies"), list):
        out["companies"] = prev["companies"]

    prev_enr = dict((prev or {}).get("enriched") or {}) if merge else {}
    if enriched is not None:
        prev_enr.update(enriched)
    out["enriched"] = prev_enr
    out["enriched_count"] = len(prev_enr)
    if isinstance(out.get("companies"), list):
        out["company_count"] = len(out["companies"])

    path.write_text(json.dumps(out, indent=2, ensure_ascii=False, default=str), encoding="utf-8")


def enriched_lookup_keys(company: dict[str, Any]) -> list[str]:
    """Keys used to match a company against the enriched checkpoint map."""
    keys: list[str] = []
    name = str(company.get("name") or "").strip()
    dom = str(company.get("domain") or "").strip().lower()
    if name:
        keys.append(name)
    if dom:
        keys.append(dom)
        # also bare host without www.
        if dom.startswith("www."):
            keys.append(dom[4:])
    return keys


def company_already_enriched(company: dict[str, Any], enriched: dict[str, Any]) -> bool:
    if not enriched:
        return False
    for k in enriched_lookup_keys(company):
        if k in enriched:
            return True
    return False


def phase4_classify_path(*, query: str, market_slug: str | None = None) -> Path:
    """Path: output/phase4/phase4_classify_<stem>.json"""
    d = _project_root() / "output" / "phase4"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"phase4_classify_{phase3_enrich_stem(query=query, market_slug=market_slug)}.json"


def load_phase4_checkpoint(path: Path | None) -> dict[str, Any] | None:
    if path is None or not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def save_phase4_checkpoint(
    path: Path,
    *,
    query: str = "",
    market_slug: str | None = None,
    classified: dict[str, Any] | None = None,
    merge: bool = True,
) -> None:
    """Persist Phase 4 classify progress. ``classified`` is keyed by name/domain."""
    from datetime import datetime, timezone

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prev = load_phase4_checkpoint(path) if merge else None
    prev_cls = dict((prev or {}).get("classified") or {}) if merge else {}
    if classified is not None:
        prev_cls.update(classified)
    # Count unique companies (prefer domain keys / avoid double-count name+domain)
    seen_ids: set[str] = set()
    for k, v in prev_cls.items():
        if not isinstance(v, dict):
            continue
        ident = str(v.get("domain") or v.get("website") or v.get("company") or k).strip().lower()
        seen_ids.add(ident or k)
    out: dict[str, Any] = {
        "phase": 4,
        "query": query or (prev or {}).get("query") or "",
        "market_slug": market_slug or (prev or {}).get("market_slug") or "",
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "classified": prev_cls,
        "classified_count": len(seen_ids),
    }
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False, default=str), encoding="utf-8")


def company_already_classified(company: dict[str, Any], classified: dict[str, Any]) -> bool:
    if not classified:
        return False
    for k in enriched_lookup_keys(company):
        if k in classified and isinstance(classified[k], dict):
            return True
    return False


def lookup_classified(
    company: dict[str, Any], classified: dict[str, Any]
) -> dict[str, Any] | None:
    for k in enriched_lookup_keys(company):
        v = classified.get(k)
        if isinstance(v, dict):
            return v
    return None


def checkpoint_summary(
    *,
    query: str,
    market_slug: str | None = None,
) -> dict[str, Any]:
    p1 = find_phase1_plan(query=query, market_slug=market_slug)
    p2 = find_phase2_companies(query=query, market_slug=market_slug, phase1_plan=p1)
    p3 = phase3_enrich_path(query=query, market_slug=market_slug)
    p3_data = load_phase3_checkpoint(p3)
    p4 = phase4_classify_path(query=query, market_slug=market_slug)
    p4_data = load_phase4_checkpoint(p4)
    return {
        "phase1_plan": str(p1) if p1 else None,
        "phase2_companies": str(p2) if p2 else None,
        "phase3_enrich": str(p3) if p3_data else None,
        "phase4_classify": str(p4) if p4_data else None,
        "can_skip_phase1": bool(p1),
        "can_skip_phase2": bool(p1 and p2),
        "can_resume_enrich": bool(p3_data and (p3_data.get("enriched") or p3_data.get("companies"))),
        "can_resume_classify": bool(p4_data and p4_data.get("classified")),
        "enrich_done_count": int((p3_data or {}).get("enriched_count") or 0),
        "classify_done_count": int((p4_data or {}).get("classified_count") or 0),
    }
