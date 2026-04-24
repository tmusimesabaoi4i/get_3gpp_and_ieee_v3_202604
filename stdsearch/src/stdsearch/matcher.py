"""Text normalization and SINGLE / AND matching."""
from __future__ import annotations

import re
import unicodedata
from typing import List, Sequence

from .models import MATCH_AND, SearchHit, SearchRule, SearchUnit


# Map a variety of hyphen / dash / minus characters to a plain ASCII '-'.
_HYPHEN_CHARS = (
    "\u2010"  # HYPHEN
    "\u2011"  # NON-BREAKING HYPHEN
    "\u2012"  # FIGURE DASH
    "\u2013"  # EN DASH
    "\u2014"  # EM DASH
    "\u2015"  # HORIZONTAL BAR
    "\u2212"  # MINUS SIGN
    "\uFF0D"  # FULLWIDTH HYPHEN-MINUS
)
_HYPHEN_TABLE = {ord(ch): "-" for ch in _HYPHEN_CHARS}

_SPACE_RE = re.compile(r"\s+")


def normalize(text: str) -> str:
    """Fold text to a canonical form for comparison.

    - NFKC unicode normalization (handles full-width / half-width).
    - Map unicode dashes / minuses / full-width hyphen-minus to ASCII '-'.
    - Case-fold.
    - Collapse runs of whitespace (including newlines) to a single space.
    """
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", text)
    s = s.translate(_HYPHEN_TABLE)
    s = s.casefold()
    s = _SPACE_RE.sub(" ", s).strip()
    return s


def contains_all(haystack_norm: str, terms_norm: Sequence[str]) -> bool:
    return all(term and term in haystack_norm for term in terms_norm)


def contains_any(haystack_norm: str, terms_norm: Sequence[str]) -> bool:
    return any(term and term in haystack_norm for term in terms_norm)


def match_rule(rule: SearchRule, units: List[SearchUnit]) -> List[SearchHit]:
    terms_norm = [normalize(t) for t in rule.terms if t]
    if not terms_norm:
        return []
    is_and = rule.match_type == MATCH_AND
    hits: List[SearchHit] = []
    for i, unit in enumerate(units):
        hay = normalize(unit.text)
        if not hay:
            continue
        ok = contains_all(hay, terms_norm) if is_and else contains_any(hay, terms_norm)
        if not ok:
            continue
        before = units[i - 1].text if i > 0 and units[i - 1].doc is unit.doc else ""
        after = units[i + 1].text if i < len(units) - 1 and units[i + 1].doc is unit.doc else ""
        hits.append(
            SearchHit(
                rule=rule,
                doc=unit.doc,
                location_type=unit.location_type,
                location_no=unit.location_no,
                heading_context=unit.heading_context,
                matched_text=unit.text,
                before_text=before,
                after_text=after,
                html_anchor=unit.html_anchor,
            )
        )
    return hits
