"""DDL + helpers for dedicated vendor_intel_exports / vendor_intel_companies tables."""

from __future__ import annotations

from typing import Any

VENDOR_INTEL_DDL = """
CREATE TABLE IF NOT EXISTS vendor_intel_exports (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  market_name TEXT NOT NULL UNIQUE,
  query TEXT,
  geography TEXT,
  export_file TEXT,
  pipeline_profile TEXT,
  scope JSONB,
  query_context JSONB,
  company_count INT NOT NULL DEFAULT 0,
  completed_at TIMESTAMPTZ,
  pushed_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_vie_market_name ON vendor_intel_exports (market_name);
CREATE INDEX IF NOT EXISTS idx_vie_geography ON vendor_intel_exports (geography);
CREATE INDEX IF NOT EXISTS idx_vie_query ON vendor_intel_exports (query);
CREATE TABLE IF NOT EXISTS vendor_intel_companies (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  export_id UUID NOT NULL REFERENCES vendor_intel_exports(id) ON DELETE CASCADE,
  sort_order INT NOT NULL DEFAULT 0,
  company_name TEXT NOT NULL,
  brand TEXT,
  country TEXT,
  role TEXT,
  website TEXT,
  key_products TEXT,
  summary TEXT,
  company_function TEXT,
  parent TEXT,
  confidence REAL,
  quality_score REAL,
  export_data JSONB NOT NULL,
  pushed_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (export_id, company_name)
);
CREATE INDEX IF NOT EXISTS idx_vic_export_id ON vendor_intel_companies (export_id);
CREATE INDEX IF NOT EXISTS idx_vic_company_name ON vendor_intel_companies (company_name);
CREATE INDEX IF NOT EXISTS idx_vic_role ON vendor_intel_companies (role);
"""


def ensure_tables(cur) -> None:
    cur.execute(VENDOR_INTEL_DDL)


def _crawl(row: dict[str, Any]) -> dict[str, Any]:
    c = row.get("_crawl")
    return c if isinstance(c, dict) else {}


def _str(v: Any) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def flatten_company_row(row: dict[str, Any]) -> dict[str, Any]:
    c = _crawl(row)
    return {
        "company_name": _str(row.get("brand") or row.get("company")) or "",
        "brand": _str(row.get("brand")),
        "country": _str(c.get("headquarters")) or _str(c.get("contact_address")),
        "role": _str(row.get("role")),
        "website": _str(row.get("website") or row.get("domain")),
        "key_products": _str(row.get("key_products")),
        "summary": _str(row.get("company_summary") or row.get("summary")),
        "company_function": _str(row.get("company_function")),
        "parent": _str(row.get("parent") or c.get("ownership_parent")),
        "confidence": row.get("confidence") if isinstance(row.get("confidence"), (int, float)) else None,
        "quality_score": row.get("quality_score") if isinstance(row.get("quality_score"), (int, float)) else None,
        "export_data": row,
    }


_PLACEHOLDERS = frozenset(
    {
        "",
        "not publicly disclosed",
        "not specified",
        "unknown",
        "n/a",
        "na",
        "-",
        "none",
        "tbd",
        "null",
    }
)


def _is_empty_value(value: Any) -> bool:
    if value is None:
        return True
    return str(value).strip().lower() in _PLACEHOLDERS


def _coalesce_str(new: Any, old: Any) -> str | None:
    n = _str(new)
    if n:
        return n
    return _str(old)


def _coalesce_num(new: Any, old: Any) -> float | int | None:
    if isinstance(new, (int, float)) and not isinstance(new, bool):
        return new
    if isinstance(old, (int, float)) and not isinstance(old, bool):
        return old
    return None


def merge_export_payload(existing: Any, incoming: dict[str, Any]) -> dict[str, Any]:
    """Merge gap-fill / re-push updates — fill empty fields only, keep good data."""
    if not isinstance(existing, dict):
        return dict(incoming)
    out = dict(existing)
    for key, val in incoming.items():
        if key == "_crawl" and isinstance(val, dict):
            old_crawl = out.get("_crawl") if isinstance(out.get("_crawl"), dict) else {}
            merged = dict(old_crawl)
            for ck, cv in val.items():
                if ck == "crawl_text" and isinstance(cv, str) and len(cv) > 12_000:
                    cv = cv[:12_000] + "…"
                if _is_empty_value(merged.get(ck)) and not _is_empty_value(cv):
                    merged[ck] = cv
            out["_crawl"] = merged
        elif key in (
            "market_query",
            "market_name",
            "pipeline_profile",
            "export_scope",
            "query_context",
            "completed_at",
        ):
            out[key] = val
        elif _is_empty_value(out.get(key)) and not _is_empty_value(val):
            out[key] = val
    return out


def merge_company_flat(existing: dict[str, Any] | None, flat: dict[str, Any]) -> dict[str, Any]:
    """Merge scalar vendor_intel_companies columns (gap-fill safe)."""
    if not existing:
        return flat
    return {
        **flat,
        "brand": _coalesce_str(flat.get("brand"), existing.get("brand")),
        "country": _coalesce_str(flat.get("country"), existing.get("country")),
        "role": _coalesce_str(flat.get("role"), existing.get("role")),
        "website": _coalesce_str(flat.get("website"), existing.get("website")),
        "key_products": _coalesce_str(flat.get("key_products"), existing.get("key_products")),
        "summary": _coalesce_str(flat.get("summary"), existing.get("summary")),
        "company_function": _coalesce_str(flat.get("company_function"), existing.get("company_function")),
        "parent": _coalesce_str(flat.get("parent"), existing.get("parent")),
        "confidence": _coalesce_num(flat.get("confidence"), existing.get("confidence")),
        "quality_score": _coalesce_num(flat.get("quality_score"), existing.get("quality_score")),
    }
