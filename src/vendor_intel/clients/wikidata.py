"""Wikidata client — free SPARQL + Search API (no API key)."""
from vendor_intel.placeholders.wikidata import (
    is_enabled,
    lookup_company,
    lookup_parent_org,
    search_entity_qid,
)

__all__ = [
    "is_enabled",
    "lookup_company",
    "lookup_parent_org",
    "search_entity_qid",
]
