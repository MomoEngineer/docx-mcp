"""Core document properties and word count for `get_metadata`.

Reads `docProps/core.xml` (title/author/created/modified) and computes a
word count by tokenizing `document.py`'s marker-free paragraph and
table-cell text - never `read_document`'s own `'# '`/`'[^N]'` rendering
markers. See [specs/get_metadata.md §3](specs/get_metadata.md#3-output-schema)
for exact field semantics, and §7 for why a malformed `docProps/core.xml`
degrades to `None` fields instead of raising, unlike a malformed
`word/document.xml`.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from docx_mcp.document import DOCUMENT_PART, NSMAP, get_body, paragraph_plain_text
from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    InvalidDocumentError,
    parse_xml,
    read_optional_part,
    read_part,
    validate_and_open,
)

CORE_PART = "docProps/core.xml"

DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
_CORE_NSMAP = {"dc": DC_NS, "dcterms": DCTERMS_NS}


@dataclass(frozen=True)
class DocumentMetadata:
    """The full result of `get_document_metadata` (see specs/get_metadata.md §3)."""

    title: str | None
    author: str | None
    created: str | None
    modified: str | None
    word_count: int


def _text_or_none(core_root: etree._Element, xpath: str) -> str | None:
    """An element's text, or `None` if the element is absent or its text is empty."""
    element = core_root.find(xpath, namespaces=_CORE_NSMAP)
    if element is None or not element.text:
        return None
    return element.text


def _word_count(body: etree._Element) -> int:
    """Whitespace-token count across every top-level paragraph and table cell.

    Paragraphs/cells are joined with `"\\n"` before tokenizing (so a word
    never spans two of them), but runs *within* one paragraph/cell are
    concatenated with no inserted separator - a word Word split across two
    `w:r`/`w:t` runs is still one word, not two (spec §3).
    """
    texts = [paragraph_plain_text(p) for p in body.findall("w:p", namespaces=NSMAP)]
    for table in body.findall("w:tbl", namespaces=NSMAP):
        for row in table.findall("w:tr", namespaces=NSMAP):
            for cell in row.findall("w:tc", namespaces=NSMAP):
                cell_paragraphs = cell.findall("w:p", namespaces=NSMAP)
                texts.append("\n".join(paragraph_plain_text(p) for p in cell_paragraphs))
    return len("\n".join(texts).split())


def get_document_metadata(
    docx_path: Path, *, max_size_bytes: int = MAX_DOCX_SIZE_BYTES
) -> DocumentMetadata:
    """Read a `.docx`'s core properties and compute its word count.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots (see
            `docx_mcp.security.resolve_safe_path`); this function does not
            perform any sandboxing itself.
        max_size_bytes: Reject the file if it is larger than this, before it
            is opened as a ZIP archive.

    Returns:
        A `DocumentMetadata` per [specs/get_metadata.md §3](specs/get_metadata.md#3-output-schema).
        `title`/`author`/`created`/`modified` are `None` when `docProps/core.xml`
        is missing, malformed, or the individual element is absent/empty;
        `word_count` never depends on `docProps/core.xml`.

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, or is missing or
            has a malformed `word/document.xml`.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        document_xml = read_part(archive, DOCUMENT_PART)
        core_xml = read_optional_part(archive, CORE_PART)

    document_root = parse_xml(document_xml, part_name=DOCUMENT_PART)
    body = get_body(document_root)
    word_count = _word_count(body)

    title: str | None = None
    author: str | None = None
    created: str | None = None
    modified: str | None = None

    if core_xml is not None:
        try:
            core_root: etree._Element | None = parse_xml(core_xml, part_name=CORE_PART)
        except InvalidDocumentError:
            core_root = None
        if core_root is not None:
            title = _text_or_none(core_root, "dc:title")
            author = _text_or_none(core_root, "dc:creator")
            created = _text_or_none(core_root, "dcterms:created")
            modified = _text_or_none(core_root, "dcterms:modified")

    return DocumentMetadata(
        title=title, author=author, created=created, modified=modified, word_count=word_count
    )
