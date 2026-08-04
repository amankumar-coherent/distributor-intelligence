"""Discovery search query quality — block listicle-style queries before DDG."""
from __future__ import annotations

import re

_LISTICLE_QUERY_RE = re.compile(
    r"\b(?:top\s+\d+|best\s+\d+|\d+\s+(?:best|top)|list\s+of|companies\s+list|"
    r"full\s+list|largest\s+\d+|leading\s+\d+|ranking\s+of|wiki\s+)\b|"
    r"\b(?:market\s+leaders?|industry\s+players?|company\s+names)\b|"
    r"\b(?:top|best)\s+(?:\w+\s+){0,5}(?:companies|company|firms?|manufacturers?|vendors?)\b|"
    r"\b(?:companies|company|firms?)\s+list\b|"
    r"\blist\s+of\s+(?:companies|firms?|manufacturers?)\b",
    re.I,
)

_LISTICLE_QUERY_SUBS: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\bcompanies\s+list\b", re.I), "manufacturers official site"),
    (re.compile(r"\blist\s+of\b", re.I), ""),
    (re.compile(r"\b(?:top|best)\s+", re.I), ""),
    (re.compile(r"\btop\s+\d+\b", re.I), ""),
    (re.compile(r"\bbest\s+\d+\b", re.I), ""),
    (re.compile(r"\bleading\s+\d+\b", re.I), ""),
    (re.compile(r"\blargest\s+\d+\b", re.I), ""),
)


def is_listicle_discovery_query(text: str) -> bool:
    """Queries that mostly return directories and listicles, not company sites."""
    t = (text or "").strip()
    if len(t) < 6:
        return True
    return bool(_LISTICLE_QUERY_RE.search(t))


def sanitize_discovery_query(text: str) -> str:
    """Best-effort rewrite; returns empty if still listicle-heavy."""
    t = re.sub(r"\s+", " ", (text or "").strip())
    for pat, repl in _LISTICLE_QUERY_SUBS:
        t = pat.sub(repl, t)
    t = re.sub(r"\s+", " ", t).strip()
    if is_listicle_discovery_query(t):
        return ""
    return t
