"""Parse individual HTML pages produced by stdharvest.

Outputs a list of 'blocks' (paragraph-like structural units), each with enough
metadata for the splitter to build sentence/paragraph/block/section scopes.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from bs4 import BeautifulSoup, Tag

from .models import DOC_GENERIC, DOC_PPTX, DOC_WORD, TargetDocument

logger = logging.getLogger(__name__)


# Tags we consider as "blocks" inside Word HTML body.
WORD_BLOCK_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "tr"}
WORD_HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6"}


@dataclass
class ParsedBlock:
    """A single structural block inside a document."""

    block_type: str              # "heading", "paragraph", "list-item", "table-row", "slide-text"
    text: str
    heading_context: str = ""    # nearest heading stack, " > " joined
    anchor: str = ""             # "#slide-003" or "#block-0007" or ""
    location_type: str = ""      # "slide" / "paragraph" / "list-item" / "heading" / "table-row"
    location_no: str = ""        # "3", "block_0007", ...
    slide_no: Optional[int] = None


@dataclass
class ParsedDocument:
    doc: TargetDocument
    blocks: List[ParsedBlock] = field(default_factory=list)


def _tag_text(tag: Tag) -> str:
    """Extract readable text from a tag (joined with single spaces)."""
    if tag is None:
        return ""
    raw = tag.get_text(separator=" ", strip=True)
    # Collapse runs of whitespace.
    return " ".join(raw.split())


def _detect_doc_type(soup: BeautifulSoup) -> str:
    if soup.find("section", class_="slide") is not None:
        return DOC_PPTX
    if soup.find("div", class_="card docx") is not None:
        return DOC_WORD
    # `<div class="card docx">` may be rendered with a second space separator
    if soup.select_one("div.card.docx") is not None:
        return DOC_WORD
    return DOC_GENERIC


def _pptx_blocks(soup: BeautifulSoup) -> List[ParsedBlock]:
    blocks: List[ParsedBlock] = []
    for section in soup.select("section.slide"):
        slide_id = section.get("id") or ""
        slide_no: Optional[int] = None
        if slide_id.startswith("slide-"):
            try:
                slide_no = int(slide_id.replace("slide-", ""))
            except ValueError:
                slide_no = None

        h3 = section.find("h3")
        heading = _tag_text(h3) if h3 else (f"Slide {slide_no}" if slide_no else "Slide")

        text_div = section.select_one("div.slide-text")
        text = _tag_text(text_div) if text_div else ""
        anchor = f"#{slide_id}" if slide_id else ""

        blocks.append(
            ParsedBlock(
                block_type="slide-text",
                text=text,
                heading_context=heading,
                anchor=anchor,
                location_type="slide",
                location_no=str(slide_no) if slide_no else "",
                slide_no=slide_no,
            )
        )
    return blocks


def _word_blocks(soup: BeautifulSoup) -> List[ParsedBlock]:
    body = soup.select_one("div.card.docx") or soup.find("div", class_="card docx")
    if body is None:
        return []
    blocks: List[ParsedBlock] = []
    heading_stack: List[str] = []
    counter = 0

    def _push(block_type: str, tag: Tag, location_type: str) -> None:
        nonlocal counter
        text = _tag_text(tag)
        if not text:
            return
        counter += 1
        anchor = ""
        if tag.get("id"):
            anchor = f"#{tag.get('id')}"
        location_no = tag.get("id") or f"block_{counter:04d}"
        blocks.append(
            ParsedBlock(
                block_type=block_type,
                text=text,
                heading_context=" > ".join(heading_stack) if heading_stack else "",
                anchor=anchor,
                location_type=location_type,
                location_no=location_no,
            )
        )

    for tag in body.find_all(True, recursive=True):
        name = (tag.name or "").lower()
        if name not in WORD_BLOCK_TAGS:
            continue
        # Skip tr inside thead/tbody/table but ensure we only count each row once.
        if name in WORD_HEADING_TAGS:
            level = int(name[1])
            heading_stack[:] = heading_stack[: level - 1]
            heading_stack.append(_tag_text(tag))
            _push("heading", tag, "heading")
        elif name == "p":
            _push("paragraph", tag, "paragraph")
        elif name == "li":
            _push("list-item", tag, "list-item")
        elif name == "tr":
            _push("table-row", tag, "table-row")
    return blocks


def _generic_blocks(soup: BeautifulSoup) -> List[ParsedBlock]:
    """Fallback: use any <p> and heading within <body>."""
    body = soup.body or soup
    blocks: List[ParsedBlock] = []
    counter = 0
    heading_stack: List[str] = []
    for tag in body.find_all(["h1", "h2", "h3", "h4", "h5", "p", "li"]):
        name = tag.name.lower()
        text = _tag_text(tag)
        if not text:
            continue
        counter += 1
        if name in WORD_HEADING_TAGS:
            level = int(name[1])
            heading_stack[:] = heading_stack[: level - 1]
            heading_stack.append(text)
            block_type = "heading"
            loc_type = "heading"
        elif name == "li":
            block_type = "list-item"
            loc_type = "list-item"
        else:
            block_type = "paragraph"
            loc_type = "paragraph"
        anchor = f"#{tag.get('id')}" if tag.get("id") else ""
        blocks.append(
            ParsedBlock(
                block_type=block_type,
                text=text,
                heading_context=" > ".join(heading_stack),
                anchor=anchor,
                location_type=loc_type,
                location_no=tag.get("id") or f"block_{counter:04d}",
            )
        )
    return blocks


def parse_html(doc: TargetDocument) -> ParsedDocument:
    if not doc.html_path.exists():
        logger.warning("HTML missing on disk: %s", doc.html_path)
        return ParsedDocument(doc=doc, blocks=[])
    try:
        raw = doc.html_path.read_text(encoding="utf-8", errors="ignore")
    except OSError as exc:
        logger.error("Cannot read HTML %s: %s", doc.html_path, exc)
        return ParsedDocument(doc=doc, blocks=[])

    soup = BeautifulSoup(raw, "lxml")
    doc_type = _detect_doc_type(soup)
    doc.doc_type = doc_type

    if doc_type == DOC_PPTX:
        blocks = _pptx_blocks(soup)
    elif doc_type == DOC_WORD:
        blocks = _word_blocks(soup)
    else:
        blocks = _generic_blocks(soup)
    return ParsedDocument(doc=doc, blocks=blocks)
