"""Data classes for stdsearch."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


# ---------- MatchType ----------
MATCH_SINGLE = "SINGLE"
MATCH_AND = "AND"
VALID_MATCH_TYPES = {MATCH_SINGLE, MATCH_AND}


# ---------- Scope ----------
SCOPE_SENTENCE = "sentence"
SCOPE_PARAGRAPH = "paragraph"
SCOPE_BLOCK = "block"
SCOPE_SECTION = "section"
VALID_SCOPES = {SCOPE_SENTENCE, SCOPE_PARAGRAPH, SCOPE_BLOCK, SCOPE_SECTION}


# ---------- doc_type ----------
DOC_WORD = "word"
DOC_PPTX = "pptx"
DOC_GENERIC = "generic"


@dataclass
class SearchRule:
    """A single row in Sheet1 (>=5) selecting what to search."""

    row_no: int
    use: bool
    rule_name: str
    match_type: str                 # SINGLE / AND
    terms: List[str]                # cleaned (trimmed) search terms
    scope: str                      # sentence / paragraph / block / section
    notes: str = ""


@dataclass
class SearchProject:
    """Top of the input Excel."""

    project_name: str
    job_folder: Path
    output_folder: Optional[Path]   # None → JobFolder/search/<date>_<project>
    rules: List[SearchRule] = field(default_factory=list)


@dataclass
class TargetDocument:
    """One individual HTML file discovered via manifest.json."""

    seq: int
    row_no: int
    title: str
    source_type: str                 # 3gpp / ieee / ""
    doc_type: str = DOC_GENERIC      # determined by html_parser
    source_file: str = ""
    pdf_path: str = ""
    html_path: Path = field(default_factory=lambda: Path())
    size_bytes: int = 0
    status: str = ""


@dataclass
class SearchUnit:
    """One 'search unit' (sentence, paragraph, block, section) extracted from a doc."""

    doc: TargetDocument
    location_type: str               # slide / paragraph / section / sentence ...
    location_no: str                 # "4" for slide 4, "block_0003", ...
    heading_context: str
    text: str
    html_anchor: str                 # "" or "#slide-004" / "#block-0003"


@dataclass
class SearchHit:
    """A hit produced by a rule against a search unit."""

    rule: SearchRule
    doc: TargetDocument
    location_type: str
    location_no: str
    heading_context: str
    matched_text: str
    before_text: str
    after_text: str
    html_anchor: str
