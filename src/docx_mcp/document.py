"""OOXML text extraction for `read_document`.

Opens a `.docx` as a ZIP archive and parses `word/document.xml` with a
hardened `lxml` parser (see [ADR-0001](../../docs/adr/0001-ooxml-library-and-module-layout.md)
and [docs/security-model.md §4](../../docs/security-model.md#4-input-validation)),
rendering the body as one line of text per paragraph with Markdown-style
heading markers and inline footnote-anchor markers. See
[specs/read_document.md](specs/read_document.md) for the exact output format
and documented limitations.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NSMAP = {"w": WORD_NS}

_W_T = f"{{{WORD_NS}}}t"
_W_TAB = f"{{{WORD_NS}}}tab"
_W_FOOTNOTE_REFERENCE = f"{{{WORD_NS}}}footnoteReference"
_W_ID = f"{{{WORD_NS}}}id"
_W_VAL = f"{{{WORD_NS}}}val"

MAX_DOCX_SIZE_BYTES = 50 * 1024 * 1024
_HEADING_STYLE_PREFIX = "Heading"
_DOCUMENT_PART = "word/document.xml"


class InvalidDocumentError(Exception):
    """Raised when a file is not a readable, valid `.docx` document."""


def _hardened_parser() -> etree.XMLParser:
    """Build an `lxml` parser that never resolves entities or touches the network.

    `resolve_entities=False` and `load_dtd=False` together mean a crafted
    external entity (XXE) in a `.docx`'s XML is never substituted and never
    read; `no_network=True` blocks any network-based entity/DTD fetch as
    defense in depth; `huge_tree=False` (the default) keeps libxml2's built-in
    entity-expansion ("billion laughs") limits active. See
    [ADR-0001](../../docs/adr/0001-ooxml-library-and-module-layout.md).
    """
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        huge_tree=False,
        dtd_validation=False,
        load_dtd=False,
    )


def _heading_level(style_id: str | None) -> int | None:
    """Return the heading level for a `pStyle` id, or `None` if it isn't a heading.

    Only recognizes Word's built-in heading style ids (`Heading1`..`Heading9`);
    see the "Heading detection" limitation in
    [specs/read_document.md](specs/read_document.md#5-limitations-non-goals).
    """
    if not style_id or not style_id.startswith(_HEADING_STYLE_PREFIX):
        return None
    suffix = style_id[len(_HEADING_STYLE_PREFIX) :]
    return int(suffix) if suffix.isdigit() else None


def _paragraph_style_id(paragraph: etree._Element) -> str | None:
    style = paragraph.find("w:pPr/w:pStyle", namespaces=_NSMAP)
    if style is None:
        return None
    value = style.get(_W_VAL)
    return str(value) if value is not None else None


def _render_paragraph(paragraph: etree._Element) -> str:
    """Render one `w:p` element as a single line of text with inline markers.

    Walks every descendant in document order and collects only the nodes that
    carry visible text or an anchor marker: `w:t` (plain text; deliberately
    excludes `w:delText`, so tracked-change deletions are never included, while
    tracked-change insertions - ordinary `w:t` inside `w:ins` - are), `w:tab`
    (rendered as a literal tab), and `w:footnoteReference` (rendered as
    `[^N]` at its exact position). See
    [specs/read_document.md §3](specs/read_document.md#3-output-schema).
    """
    parts: list[str] = []
    for node in paragraph.iter():
        tag = node.tag
        if tag == _W_T:
            parts.append(node.text or "")
        elif tag == _W_TAB:
            parts.append("\t")
        elif tag == _W_FOOTNOTE_REFERENCE:
            footnote_id = node.get(_W_ID)
            parts.append(f"[^{footnote_id}]")

    text = "".join(parts)
    level = _heading_level(_paragraph_style_id(paragraph))
    if level is not None:
        return f"{'#' * level} {text}"
    return text


def _parse_document_xml(raw_xml: bytes) -> etree._Element:
    try:
        return etree.fromstring(raw_xml, parser=_hardened_parser())
    except etree.XMLSyntaxError as exc:
        message = f"not a valid .docx file: malformed XML in {_DOCUMENT_PART}: {exc}"
        raise InvalidDocumentError(message) from exc


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
    if not docx_path.is_file():
        raise InvalidDocumentError(f"file not found: {docx_path}")

    size = docx_path.stat().st_size
    if size > max_size_bytes:
        raise InvalidDocumentError(
            f"file exceeds the maximum allowed size ({size} > {max_size_bytes} bytes)"
        )

    try:
        with zipfile.ZipFile(docx_path) as archive:
            try:
                raw_xml = archive.read(_DOCUMENT_PART)
            except KeyError as exc:
                raise InvalidDocumentError(
                    f"not a valid .docx file: missing {_DOCUMENT_PART}"
                ) from exc
    except zipfile.BadZipFile as exc:
        raise InvalidDocumentError("not a valid .docx file: not a ZIP archive") from exc

    root = _parse_document_xml(raw_xml)
    body = root.find("w:body", namespaces=_NSMAP)
    if body is None:
        raise InvalidDocumentError(f"not a valid .docx file: no <w:body> in {_DOCUMENT_PART}")

    paragraphs = body.findall("w:p", namespaces=_NSMAP)
    return "\n".join(_render_paragraph(p) for p in paragraphs)
