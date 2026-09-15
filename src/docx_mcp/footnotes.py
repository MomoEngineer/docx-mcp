"""Footnote content resolution for `get_footnotes`.

Combines `document.py`'s shared, anchor-driven `find_footnote_anchors` (also
used by `docx_mcp.structure`'s footnote-anchor index, per
[ADR-0004](../../docs/adr/0004-phase-3-footnote-module-and-shared-anchor-resolution.md))
with content resolved from `word/footnotes.xml`, a part `docx_mcp.structure`
deliberately never opens (see
[specs/get_structure.md §3](specs/get_structure.md#3-output-schema)). See
[specs/get_footnotes.md](specs/get_footnotes.md) for the exact output format,
population rules, and error/degradation behavior.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from docx_mcp.document import (
    DOCUMENT_PART,
    NSMAP,
    WORD_NS,
    find_footnote_anchors,
    get_body,
    paragraph_plain_text,
)
from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    parse_xml,
    read_optional_part,
    read_part,
    validate_and_open,
)

FOOTNOTES_PART = "word/footnotes.xml"

_W_ID = f"{{{WORD_NS}}}id"


@dataclass(frozen=True)
class FootnoteEntry:
    """One footnote reference, resolved to its content (see specs/get_footnotes.md §3)."""

    id: str
    paragraph_index: int
    content: str | None


def render_footnote_content(footnote: etree._Element) -> str:
    """Render one declared `<w:footnote>` element's content as `get_footnotes` does.

    Its `w:p` children are rendered with `paragraph_plain_text` (marker-free),
    joined by `"\\n"` for a multi-paragraph footnote - deliberately not the
    marker-inclusive `render_paragraph` `get_structure` uses for table-cell
    text, since footnote content is a terminal leaf never itself navigated
    into (spec §3). Exported so `docx_mcp.footnote_edit` can compute
    `edit_footnote`'s `previous_content` from the exact same rendering
    `get_footnotes` uses, rather than a second, independently-consistent
    copy of it - the same reasoning
    [ADR-0004](../../docs/adr/0004-phase-3-footnote-module-and-shared-anchor-resolution.md)
    already established for anchor-finding (see also ADR-0007).
    """
    paragraphs = footnote.findall("w:p", namespaces=NSMAP)
    return "\n".join(paragraph_plain_text(p) for p in paragraphs)


def _content_by_id(footnotes_root: etree._Element) -> dict[str, str]:
    """Map every declared `w:footnote/@w:id` to its rendered content.

    Word's own boilerplate separator/continuationSeparator footnotes are
    harmlessly included here too (they simply never match an anchor id, since
    Word never emits a body `w:footnoteReference` for them - spec §5).
    """
    content: dict[str, str] = {}
    for footnote in footnotes_root.findall("w:footnote", namespaces=NSMAP):
        footnote_id = footnote.get(_W_ID)
        if footnote_id is None:
            continue
        content[footnote_id] = render_footnote_content(footnote)
    return content


def get_document_footnotes(
    docx_path: Path, *, max_size_bytes: int = MAX_DOCX_SIZE_BYTES
) -> tuple[FootnoteEntry, ...]:
    """List every footnote a `.docx`'s body references, with anchor and content.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots (see
            `docx_mcp.security.resolve_safe_path`); this function does not
            perform any sandboxing itself.
        max_size_bytes: Reject the file if it is larger than this, before it
            is opened as a ZIP archive.

    Returns:
        A `FootnoteEntry` per `w:footnoteReference` found in the body, in
        document order, per
        [specs/get_footnotes.md §3](specs/get_footnotes.md#3-output-schema).
        `content` is `None` when `word/footnotes.xml` is absent or does not
        declare a matching id - never an error by itself.

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, is missing or has a
            malformed `word/document.xml`, or has a `word/footnotes.xml` that
            is present but malformed.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        document_xml = read_part(archive, DOCUMENT_PART)
        footnotes_xml = read_optional_part(archive, FOOTNOTES_PART)

    document_root = parse_xml(document_xml, part_name=DOCUMENT_PART)
    body = get_body(document_root)
    anchors = find_footnote_anchors(body)

    content_by_id: dict[str, str] = {}
    if footnotes_xml is not None:
        footnotes_root = parse_xml(footnotes_xml, part_name=FOOTNOTES_PART)
        content_by_id = _content_by_id(footnotes_root)

    return tuple(
        FootnoteEntry(
            id=footnote_id,
            paragraph_index=paragraph_index,
            content=content_by_id.get(footnote_id),
        )
        for footnote_id, paragraph_index in anchors
    )
