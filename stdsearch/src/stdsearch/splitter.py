"""Convert ParsedBlock lists into SearchUnit lists according to Scope."""
from __future__ import annotations

import re
from typing import List

from .html_parser import ParsedBlock, ParsedDocument
from .models import (
    DOC_PPTX,
    DOC_WORD,
    SCOPE_BLOCK,
    SCOPE_PARAGRAPH,
    SCOPE_SECTION,
    SCOPE_SENTENCE,
    SearchUnit,
    TargetDocument,
)


# Sentence boundary: '.', '!', '?', 。 ！ ？, newline. Keep trailing punctuation.
_SENT_SPLIT = re.compile(r"(?<=[\.!\?。！？])\s+|[\r\n]+")


def _split_sentences(text: str) -> List[str]:
    if not text:
        return []
    parts = _SENT_SPLIT.split(text)
    return [p.strip() for p in parts if p and p.strip()]


def _split_paragraph_lines(text: str) -> List[str]:
    if not text:
        return []
    parts = re.split(r"[\r\n]+", text)
    return [p.strip() for p in parts if p and p.strip()]


def _unit(
    doc: TargetDocument,
    block: ParsedBlock,
    text: str,
    override_type: str = "",
    override_no: str = "",
) -> SearchUnit:
    return SearchUnit(
        doc=doc,
        location_type=override_type or block.location_type,
        location_no=override_no or block.location_no,
        heading_context=block.heading_context,
        text=text,
        html_anchor=block.anchor,
    )


# ---------- PPTX splitting ----------
def _split_pptx(doc: TargetDocument, blocks: List[ParsedBlock], scope: str) -> List[SearchUnit]:
    units: List[SearchUnit] = []
    for blk in blocks:
        if not blk.text:
            continue
        if scope == SCOPE_SENTENCE:
            for i, sent in enumerate(_split_sentences(blk.text), start=1):
                units.append(
                    _unit(doc, blk, sent,
                          override_type="sentence",
                          override_no=f"{blk.location_no}.s{i:02d}" if blk.location_no else f"s{i:02d}")
                )
        elif scope == SCOPE_PARAGRAPH:
            for i, line in enumerate(_split_paragraph_lines(blk.text), start=1):
                units.append(
                    _unit(doc, blk, line,
                          override_type="paragraph",
                          override_no=f"{blk.location_no}.p{i:02d}" if blk.location_no else f"p{i:02d}")
                )
        else:  # block or section → entire slide
            units.append(
                _unit(doc, blk, blk.text, override_type="slide", override_no=blk.location_no)
            )
    return units


# ---------- Word splitting ----------
def _split_word(doc: TargetDocument, blocks: List[ParsedBlock], scope: str) -> List[SearchUnit]:
    if scope == SCOPE_SECTION:
        return _word_sections(doc, blocks)

    units: List[SearchUnit] = []
    for blk in blocks:
        if not blk.text:
            continue
        if scope == SCOPE_SENTENCE:
            for i, sent in enumerate(_split_sentences(blk.text), start=1):
                units.append(
                    _unit(doc, blk, sent,
                          override_type="sentence",
                          override_no=f"{blk.location_no}.s{i:02d}")
                )
        elif scope == SCOPE_PARAGRAPH:
            # Each block (p / li / tr / heading) is already a paragraph-like row.
            units.append(_unit(doc, blk, blk.text))
        else:  # block
            units.append(_unit(doc, blk, blk.text, override_type="block"))
    return units


def _word_sections(doc: TargetDocument, blocks: List[ParsedBlock]) -> List[SearchUnit]:
    """Section = from one heading to the next heading of same or higher level."""
    units: List[SearchUnit] = []
    current_heading_block: ParsedBlock | None = None
    buffer: List[str] = []
    section_idx = 0

    def _flush() -> None:
        nonlocal section_idx
        if not buffer:
            return
        section_idx += 1
        text = "\n".join(buffer).strip()
        if not text:
            return
        head = current_heading_block
        if head is None:
            anchor = ""
            heading = ""
            location_no = f"section_{section_idx:04d}"
        else:
            anchor = head.anchor
            heading = head.heading_context or head.text
            location_no = head.location_no or f"section_{section_idx:04d}"
        units.append(
            SearchUnit(
                doc=doc,
                location_type="section",
                location_no=location_no,
                heading_context=heading,
                text=text,
                html_anchor=anchor,
            )
        )

    for blk in blocks:
        if blk.block_type == "heading":
            _flush()
            buffer = [blk.text]
            current_heading_block = blk
        else:
            if blk.text:
                buffer.append(blk.text)
    _flush()
    return units


def split_document(parsed: ParsedDocument, scope: str) -> List[SearchUnit]:
    doc = parsed.doc
    if scope not in (SCOPE_SENTENCE, SCOPE_PARAGRAPH, SCOPE_BLOCK, SCOPE_SECTION):
        scope = SCOPE_PARAGRAPH
    if doc.doc_type == DOC_PPTX:
        return _split_pptx(doc, parsed.blocks, scope)
    if doc.doc_type == DOC_WORD:
        return _split_word(doc, parsed.blocks, scope)
    # generic
    return _split_word(doc, parsed.blocks, scope)
