"""WordprocessingML paragraph rendering: `extract_text` for `read_document`.

Opens a `.docx` as a ZIP archive (via `docx_mcp.ooxml`) and parses
`word/document.xml` with a hardened `lxml` parser (see
[ADR-0001](../../docs/adr/0001-ooxml-library-and-module-layout.md) and
[docs/security-model.md §4](../../docs/security-model.md#4-input-validation)),
rendering the body as one line of text per paragraph with Markdown-style
heading markers and inline footnote-anchor markers. See
[specs/read_document.md](specs/read_document.md) for the exact output format
and documented limitations.

This is also the one place `w:p`/`w:t`/`w:tab`/`w:footnoteReference`
traversal happens: `paragraph_plain_text` (the marker-free half of
`_render_paragraph`'s logic) and `paragraph_style_id` are reused by
`docx_mcp.structure` and `docx_mcp.metadata` rather than re-implemented, per
[ADR-0003](../../docs/adr/0003-phase-2-module-layout.md). `find_footnote_anchors`
extends that ownership to anchor-finding: it is the one place a footnote
anchor's `(id, paragraph_index)` is computed, reused by `docx_mcp.structure`
(the footnote-anchor index) and `docx_mcp.footnotes` (content resolution), per
[ADR-0004](../../docs/adr/0004-phase-3-footnote-module-and-shared-anchor-resolution.md) -
so the two tools cannot disagree on which paragraph anchors which footnote id.
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    InvalidDocumentError,
    parse_xml,
    read_part,
    validate_and_open,
)

__all__ = [
    "MAX_DOCX_SIZE_BYTES",
    "InvalidDocumentError",
    "NSMAP",
    "WORD_NS",
    "extract_text",
    "find_footnote_anchors",
    "get_body",
    "heading_level_from_style_prefix",
    "paragraph_plain_text",
    "paragraph_style_id",
    "render_paragraph",
]

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NS}

_W_T = f"{{{WORD_NS}}}t"
_W_TAB = f"{{{WORD_NS}}}tab"
_W_FOOTNOTE_REFERENCE = f"{{{WORD_NS}}}footnoteReference"
_W_ID = f"{{{WORD_NS}}}id"
_W_VAL = f"{{{WORD_NS}}}val"

_HEADING_STYLE_PREFIX = "Heading"
DOCUMENT_PART = "word/document.xml"


def heading_level_from_style_prefix(style_id: str | None) -> int | None:
    """Return the heading level implied by a `pStyle` id's `HeadingN` prefix, or `None`.

    Only recognizes Word's built-in heading style ids (`Heading1`..`Heading9`)
    by their id text; this is the Phase 1 heuristic, kept as `get_structure`'s
    last-resort fallback (see
    [specs/get_structure.md §3](specs/get_structure.md#3-output-schema), step 3)
    and as `read_document`'s only heading detection (see the "Heading
    detection" limitation in
    [specs/read_document.md](specs/read_document.md#5-limitations-non-goals)).
    """
    if not style_id or not style_id.startswith(_HEADING_STYLE_PREFIX):
        return None
    suffix = style_id[len(_HEADING_STYLE_PREFIX) :]
    return int(suffix) if suffix.isdigit() else None


def paragraph_style_id(paragraph: etree._Element) -> str | None:
    """Return a paragraph's raw `w:pPr/w:pStyle/@w:val`, or `None` if it has none."""
    style = paragraph.find("w:pPr/w:pStyle", namespaces=NSMAP)
    if style is None:
        return None
    value = style.get(_W_VAL)
    return str(value) if value is not None else None


def _paragraph_text(paragraph: etree._Element, *, include_footnote_markers: bool) -> str:
    """Walk every descendant in document order, collecting visible text and,
    if requested, inline footnote-anchor markers - the shared traversal behind
    both `paragraph_plain_text` and `_render_paragraph`.

    Collects `w:t` (plain text; deliberately excludes `w:delText`, so
    tracked-change deletions are never included, while tracked-change
    insertions - ordinary `w:t` inside `w:ins` - are), `w:tab` (rendered as a
    literal tab), and, when `include_footnote_markers` is true,
    `w:footnoteReference` (rendered as `[^N]` at its exact position).
    """
    parts: list[str] = []
    for node in paragraph.iter():
        tag = node.tag
        if tag == _W_T:
            parts.append(node.text or "")
        elif tag == _W_TAB:
            parts.append("\t")
        elif tag == _W_FOOTNOTE_REFERENCE and include_footnote_markers:
            footnote_id = node.get(_W_ID)
            parts.append(f"[^{footnote_id}]")
    return "".join(parts)


def paragraph_plain_text(paragraph: etree._Element) -> str:
    """A paragraph's text with no heading `'# '` prefix and no inline `'[^N]'`
    footnote marker - those are `read_document`'s own rendering convention,
    not structural data. Shared by `docx_mcp.structure` (paragraph/table-cell
    text) and `docx_mcp.metadata` (word count), per
    [ADR-0003](../../docs/adr/0003-phase-2-module-layout.md).
    """
    return _paragraph_text(paragraph, include_footnote_markers=False)


def render_paragraph(paragraph: etree._Element) -> str:
    """Render one `w:p` element as a single line of text with inline markers.

    See [specs/read_document.md §3](specs/read_document.md#3-output-schema).
    Reused by `docx_mcp.structure` for table-cell text, which - per
    [specs/get_structure.md §3](specs/get_structure.md#3-output-schema) - uses
    this same marker-inclusive rendering, since a cell can validly contain a
    heading-styled or footnote-anchored paragraph.
    """
    text = _paragraph_text(paragraph, include_footnote_markers=True)
    level = heading_level_from_style_prefix(paragraph_style_id(paragraph))
    if level is not None:
        return f"{'#' * level} {text}"
    return text


def get_body(document_root: etree._Element, *, part_name: str = DOCUMENT_PART) -> etree._Element:
    """Return `word/document.xml`'s `<w:body>`, or raise if it is missing.

    Shared by `extract_text`, `docx_mcp.structure`, and `docx_mcp.metadata`,
    so "a `.docx` without a `<w:body>` is invalid" has exactly one check.
    """
    body = document_root.find("w:body", namespaces=NSMAP)
    if body is None:
        raise InvalidDocumentError(f"not a valid .docx file: no <w:body> in {part_name}")
    return body


def find_footnote_anchors(body: etree._Element) -> tuple[tuple[str, int], ...]:
    """Every `w:footnoteReference` found while walking `body`'s top-level
    paragraphs, as `(id, paragraph_index)` pairs in document order.

    Anchor-driven and presentation-agnostic: this returns plain tuples, not a
    tool-specific dataclass, so `docx_mcp.structure` (which pairs each anchor
    with its own `FootnoteAnchor` output type) and `docx_mcp.footnotes` (which
    pairs each anchor with resolved content) can each build their own typed
    result from the same underlying data without depending on each other's
    output schema - see
    [ADR-0004](../../docs/adr/0004-phase-3-footnote-module-and-shared-anchor-resolution.md).
    Only top-level `w:body/w:p` paragraphs are walked - a footnote reference
    inside a table cell is not found here, matching `get_structure`'s existing
    `paragraph_index` scope.
    """
    anchors: list[tuple[str, int]] = []
    for index, paragraph in enumerate(body.findall("w:p", namespaces=NSMAP)):
        for footnote_ref in paragraph.iter(_W_FOOTNOTE_REFERENCE):
            footnote_id = footnote_ref.get(_W_ID)
            if footnote_id is not None:
                anchors.append((footnote_id, index))
    return tuple(anchors)


def extract_text(docx_path: Path, *, max_size_bytes: int = MAX_DOCX_SIZE_BYTES) -> str:
    """Extract the full body text of a `.docx`, with heading and footnote markers inline.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots (see
            `docx_mcp.security.resolve_safe_path`); this function does not
            perform any sandboxing itself.
        max_size_bytes: Reject the file if it is larger than this, before it
            is opened as a ZIP archive. Defaults to `MAX_DOCX_SIZE_BYTES`;
            overridable for tests.

    Returns:
        The document's paragraphs joined by `"\\n"`, in document order, per
        [specs/read_document.md §3](specs/read_document.md#3-output-schema).

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, is missing
            `word/document.xml`, or that part is not well-formed XML.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)

    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    body = get_body(root)

    paragraphs = body.findall("w:p", namespaces=NSMAP)
    return "\n".join(render_paragraph(p) for p in paragraphs)
