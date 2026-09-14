"""Structural extraction for `get_structure`.

Builds a paragraph index with resolved heading levels, a heading-derived
table of contents, every table rendered as rows of cell text, and a
lightweight footnote-anchor index. Heading-level resolution walks
`word/styles.xml`'s `w:outlineLvl`/`w:basedOn` chain, falling back to
`docx_mcp.document`'s `HeadingN` digit-suffix heuristic - see
[specs/get_structure.md §3](specs/get_structure.md#3-output-schema) for the
exact resolution order and [ADR-0003](../../docs/adr/0003-phase-2-module-layout.md)
for why this reuses `document.py`'s paragraph-text helper instead of
re-implementing paragraph traversal.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from docx_mcp.document import (
    DOCUMENT_PART,
    NSMAP,
    WORD_NS,
    get_body,
    heading_level_from_style_prefix,
    paragraph_plain_text,
    paragraph_style_id,
    render_paragraph,
)
from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    parse_xml,
    read_optional_part,
    read_part,
    validate_and_open,
)

STYLES_PART = "word/styles.xml"

_W_STYLE_ID = f"{{{WORD_NS}}}styleId"
_W_ID = f"{{{WORD_NS}}}id"
_W_VAL = f"{{{WORD_NS}}}val"
_W_FOOTNOTE_REFERENCE = f"{{{WORD_NS}}}footnoteReference"

_BODY_TEXT_OUTLINE_VAL = 9
_MAX_BASED_ON_HOPS = 64


@dataclass(frozen=True)
class ParagraphInfo:
    """One top-level body paragraph (see spec §3, `paragraphs`)."""

    paragraph_index: int
    text: str
    style_id: str | None
    heading_level: int | None


@dataclass(frozen=True)
class TocEntry:
    """One heading, filtered from `ParagraphInfo` (see spec §3, `toc`)."""

    paragraph_index: int
    level: int
    text: str


@dataclass(frozen=True)
class TableInfo:
    """One top-level table, rendered as rows of cell text (see spec §3, `tables`)."""

    table_index: int
    rows: tuple[tuple[str, ...], ...]


@dataclass(frozen=True)
class FootnoteAnchor:
    """One footnote reference's anchor location (see spec §3, `footnotes`)."""

    id: str
    paragraph_index: int


@dataclass(frozen=True)
class DocumentStructure:
    """The full result of `get_document_structure` (see specs/get_structure.md §3)."""

    paragraphs: tuple[ParagraphInfo, ...]
    toc: tuple[TocEntry, ...]
    tables: tuple[TableInfo, ...]
    footnotes: tuple[FootnoteAnchor, ...]


def _parse_outline_val(raw_val: str | None) -> int | None:
    """Parse a raw `w:outlineLvl/@w:val` into an int in `0..9`, or `None` if
    the attribute is absent, not an integer, or outside OOXML's valid `0..9`
    range for this attribute. Deliberately distinct from `None` meaning
    "value 9, Body Text" - see `_level_from_outline_val`."""
    if raw_val is None:
        return None
    try:
        n = int(raw_val)
    except ValueError:
        return None
    return n if 0 <= n <= _BODY_TEXT_OUTLINE_VAL else None


def _level_from_outline_val(n: int) -> int | None:
    """Map a validated `0..9` outline value to a 1-based heading level.

    `0`-`8` map to levels `1`-`9`; `9` ("Body Text") maps to `None`.
    """
    return None if n == _BODY_TEXT_OUTLINE_VAL else n + 1


def _direct_outline_level(paragraph: etree._Element) -> tuple[bool, int | None]:
    """Return `(has_direct_outline_lvl, resolved_level)`.

    The first element tells the caller whether a direct `w:outlineLvl` was
    present *with a valid `0..9` value* - which is authoritative and
    short-circuits the rest of the resolution order (spec §3, step 1),
    regardless of whether it itself resolved to a heading level or to `None`
    (value `9`). An `w:outlineLvl` element that is present but carries a
    missing, non-numeric, or out-of-range value is treated the same as no
    element at all (`has_direct=False`) - a malformed direct override must
    not produce a nonsensical terminal `heading_level` (e.g. `0` from
    `val="-1"`, or `101` from `val="100"`) or block falling through to
    style-based resolution.
    """
    element = paragraph.find("w:pPr/w:outlineLvl", namespaces=NSMAP)
    if element is None:
        return False, None
    n = _parse_outline_val(element.get(_W_VAL))
    if n is None:
        return False, None
    return True, _level_from_outline_val(n)


def _build_styles_by_id(styles_root: etree._Element) -> dict[str, etree._Element]:
    styles_by_id: dict[str, etree._Element] = {}
    for style in styles_root.findall("w:style", namespaces=NSMAP):
        style_id = style.get(_W_STYLE_ID)
        if style_id is not None:
            styles_by_id[style_id] = style
    return styles_by_id


def _style_chain_outline_level(
    style_id: str, styles_by_id: dict[str, etree._Element]
) -> int | None:
    """Walk `style_id`'s `w:basedOn` chain (nearest ancestor first) looking for
    the first *validly-valued* declared `w:outlineLvl`, up to
    `_MAX_BASED_ON_HOPS` hops - a guard against a malformed circular
    `w:basedOn` chain (spec §7). A style whose own `w:outlineLvl` value is
    missing, non-numeric, or outside `0..9` is treated as not declaring one -
    the walk continues to its `w:basedOn` ancestor rather than stopping on a
    nonsensical value, same as `_direct_outline_level`'s handling.
    """
    seen: set[str] = set()
    current: str | None = style_id
    hops = 0
    while current is not None and current not in seen and hops < _MAX_BASED_ON_HOPS:
        seen.add(current)
        style = styles_by_id.get(current)
        if style is None:
            return None
        outline_element = style.find("w:pPr/w:outlineLvl", namespaces=NSMAP)
        if outline_element is not None:
            n = _parse_outline_val(outline_element.get(_W_VAL))
            if n is not None:
                return _level_from_outline_val(n)
        based_on = style.find("w:basedOn", namespaces=NSMAP)
        current = based_on.get(_W_VAL) if based_on is not None else None
        hops += 1
    return None


def _resolve_heading_level(
    paragraph: etree._Element, styles_by_id: dict[str, etree._Element]
) -> int | None:
    """Resolve a paragraph's heading level per specs/get_structure.md §3."""
    has_direct, direct_level = _direct_outline_level(paragraph)
    if has_direct:
        return direct_level

    style_id = paragraph_style_id(paragraph)
    if style_id is not None and style_id in styles_by_id:
        style_level = _style_chain_outline_level(style_id, styles_by_id)
        if style_level is not None:
            return style_level

    return heading_level_from_style_prefix(style_id)


def _cell_text(cell: etree._Element) -> str:
    """A table cell's text, rendered with the same heading/footnote markers as
    `read_document` (spec §3) - `paragraph_plain_text` is deliberately not used
    here, unlike for the top-level `paragraphs` list."""
    paragraphs = cell.findall("w:p", namespaces=NSMAP)
    return "\n".join(render_paragraph(p) for p in paragraphs)


def _render_table(table: etree._Element) -> tuple[tuple[str, ...], ...]:
    rows = []
    for row in table.findall("w:tr", namespaces=NSMAP):
        cells = row.findall("w:tc", namespaces=NSMAP)
        rows.append(tuple(_cell_text(cell) for cell in cells))
    return tuple(rows)


def get_document_structure(
    docx_path: Path, *, max_size_bytes: int = MAX_DOCX_SIZE_BYTES
) -> DocumentStructure:
    """Build a `.docx`'s structural outline: paragraphs, TOC, tables, footnote anchors.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots (see
            `docx_mcp.security.resolve_safe_path`); this function does not
            perform any sandboxing itself.
        max_size_bytes: Reject the file if it is larger than this, before it
            is opened as a ZIP archive.

    Returns:
        A `DocumentStructure` per
        [specs/get_structure.md §3](specs/get_structure.md#3-output-schema).

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, is missing or has a
            malformed `word/document.xml`, or has a `word/styles.xml` that is
            present but malformed.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        document_xml = read_part(archive, DOCUMENT_PART)
        styles_xml = read_optional_part(archive, STYLES_PART)

    document_root = parse_xml(document_xml, part_name=DOCUMENT_PART)
    body = get_body(document_root)

    styles_by_id: dict[str, etree._Element] = {}
    if styles_xml is not None:
        styles_root = parse_xml(styles_xml, part_name=STYLES_PART)
        styles_by_id = _build_styles_by_id(styles_root)

    paragraphs: list[ParagraphInfo] = []
    toc: list[TocEntry] = []
    footnotes: list[FootnoteAnchor] = []
    for index, paragraph in enumerate(body.findall("w:p", namespaces=NSMAP)):
        text = paragraph_plain_text(paragraph)
        style_id = paragraph_style_id(paragraph)
        heading_level = _resolve_heading_level(paragraph, styles_by_id)
        paragraphs.append(
            ParagraphInfo(
                paragraph_index=index, text=text, style_id=style_id, heading_level=heading_level
            )
        )
        if heading_level is not None:
            toc.append(TocEntry(paragraph_index=index, level=heading_level, text=text))
        for footnote_ref in paragraph.iter(_W_FOOTNOTE_REFERENCE):
            footnote_id = footnote_ref.get(_W_ID)
            if footnote_id is not None:
                footnotes.append(FootnoteAnchor(id=footnote_id, paragraph_index=index))

    tables = tuple(
        TableInfo(table_index=table_index, rows=_render_table(table))
        for table_index, table in enumerate(body.findall("w:tbl", namespaces=NSMAP))
    )

    return DocumentStructure(
        paragraphs=tuple(paragraphs), toc=tuple(toc), tables=tables, footnotes=tuple(footnotes)
    )
