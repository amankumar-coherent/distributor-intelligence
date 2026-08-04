"""Automated multi-source gap-fill for missing firmographics."""

from vendor_intel.enrichment.gap_fill.batch import gap_fill_batch, gap_fill_batch_async, gap_fill_status
from vendor_intel.enrichment.gap_fill.gaps import missing_fields, row_has_gaps
from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_market, gap_fill_market_from_slug

__all__ = [
    "gap_fill_batch",
    "gap_fill_batch_async",
    "gap_fill_status",
    "gap_fill_market",
    "gap_fill_market_from_slug",
    "missing_fields",
    "row_has_gaps",
]
