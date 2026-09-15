"""`add_footnote`/`edit_footnote`: extending and updating `word/footnotes.xml`.

Owns both tools' logic (a deliberate, phase-scoped exception to
one-module-per-tool - see
[ADR-0007](../../docs/adr/0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md),
the same kind of exception `text_edit.py`/`paragraph_edit.py` already
established): both share one "build a footnote's paragraph content from a
`content` string" builder (`_build_footnote_paragraphs`) and one "look up a
declared footnote by id" resolution.

`add_footnote` always appends the new `w:footnoteReference` at the end of
its target paragraph's text (no character-offset placement - see
[specs/add_footnote.md §5](specs/add_footnote.md#5-limitations-non-goals)),
allocates the new footnote's id itself (never caller-supplied), and - when
the document has no footnotes yet - creates `word/footnotes.xml` from
scratch (with Word's two boilerplate separator/continuationSeparator
footnotes) together with the `[Content_Types].xml` `Override` and
`word/_rels/document.xml.rels` `Relationship` that wire it into the OPC
package, all in one atomic step via `docx_mcp.ooxml.atomic_write_parts`.

`edit_footnote` rebuilds the target footnote's content entirely fresh from
its `content` argument - `content.split("\\n")` becomes one `w:p` per line,
the exact inverse of `docx_mcp.footnotes.render_footnote_content`'s
`"\\n".join(...)` reading convention, so a round-trip through `get_footnotes`
reproduces `content` verbatim. It never touches `word/document.xml` (the
anchor and id are unaffected), so it reuses plain `atomic_write_part`. It
refuses to edit Word's reserved boilerplate ids (`"-1"`/`"0"`), which
`get_footnotes` never exposes to a caller in the first place.
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from docx_mcp.document import DOCUMENT_PART, NSMAP, WORD_NS, get_body
from docx_mcp.footnotes import FOOTNOTES_PART, render_footnote_content
from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    InvalidDocumentError,
    atomic_write_part,
    atomic_write_parts,
    parse_xml,
    read_optional_part,
    read_part,
    validate_and_open,
)

__all__ = [
    "FootnoteEditError",
    "add_footnote_to_document",
    "edit_footnote_in_document",
]

_W_FOOTNOTES = f"{{{WORD_NS}}}footnotes"
_W_FOOTNOTE = f"{{{WORD_NS}}}footnote"
_W_P = f"{{{WORD_NS}}}p"
_W_R = f"{{{WORD_NS}}}r"
_W_T = f"{{{WORD_NS}}}t"
_W_RPR = f"{{{WORD_NS}}}rPr"
_W_RSTYLE = f"{{{WORD_NS}}}rStyle"
_W_FOOTNOTE_REF = f"{{{WORD_NS}}}footnoteRef"
_W_FOOTNOTE_REFERENCE = f"{{{WORD_NS}}}footnoteReference"
_W_VERTALIGN = f"{{{WORD_NS}}}vertAlign"
_W_SEPARATOR = f"{{{WORD_NS}}}separator"
_W_CONTINUATION_SEPARATOR = f"{{{WORD_NS}}}continuationSeparator"
_W_ID = f"{{{WORD_NS}}}id"
_W_TYPE = f"{{{WORD_NS}}}type"
_W_VAL = f"{{{WORD_NS}}}val"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

CONTENT_TYPES_PART = "[Content_Types].xml"
DOCUMENT_RELS_PART = "word/_rels/document.xml.rels"

_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_FOOTNOTES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)
_FOOTNOTES_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)
_FOOTNOTE_REFERENCE_STYLE = "FootnoteReference"

_RESERVED_FOOTNOTE_IDS = frozenset({"-1", "0"})
"""Word's boilerplate separator (`-1`) / continuationSeparator (`0`) ids -
never referenced by a body `w:footnoteReference`, so `get_footnotes` never
exposes them to a caller (see
[specs/get_footnotes.md §5](specs/get_footnotes.md#5-limitations-non-goals)).
`edit_footnote` refuses to target them (see ADR-0007 §Decision.7)."""


class FootnoteEditError(Exception):
    """Raised for an `add_footnote`/`edit_footnote`-specific error.

    See [specs/add_footnote.md §7](specs/add_footnote.md#7-error-behavior)
    and [specs/edit_footnote.md §7](specs/edit_footnote.md#7-error-behavior)
    for the full error table.
    """


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _build_footnote_paragraphs(content: str) -> list[etree._Element]:
    """Build one `w:p` per `"\\n"`-separated line of `content`.

    The first paragraph is prefixed with the standard auto-number reference
    run (`w:r/w:rPr/w:rStyle` + `w:footnoteRef`) every real footnote's first
    paragraph carries. This is the exact inverse of
    `docx_mcp.footnotes.render_footnote_content`'s `"\\n".join(...)` reading,
    so a round-trip through `get_footnotes` reproduces `content` verbatim.
    """
    paragraphs: list[etree._Element] = []
    for index, line in enumerate(content.split("\n")):
        paragraph = etree.Element(_W_P)
        if index == 0:
            ref_run = etree.SubElement(paragraph, _W_R)
            ref_run_props = etree.SubElement(ref_run, _W_RPR)
            etree.SubElement(ref_run_props, _W_RSTYLE).set(_W_VAL, _FOOTNOTE_REFERENCE_STYLE)
            etree.SubElement(ref_run, _W_FOOTNOTE_REF)
        if line:
            text_run = etree.SubElement(paragraph, _W_R)
            text_elem = etree.SubElement(text_run, _W_T)
            text_elem.text = line
            text_elem.set(_XML_SPACE, "preserve")
        paragraphs.append(paragraph)
    return paragraphs


def _build_footnote_element(footnote_id: str, content: str) -> etree._Element:
    footnote = etree.Element(_W_FOOTNOTE)
    footnote.set(_W_ID, footnote_id)
    for paragraph in _build_footnote_paragraphs(content):
        footnote.append(paragraph)
    return footnote


def _boilerplate_footnote(footnote_id: str, kind: str) -> etree._Element:
    """One of Word's two reserved separator/continuationSeparator footnotes,
    matching the shape Word itself writes - see
    `tests/fixtures/generate_fixtures.py`'s `_boilerplate_footnote`, which
    this mirrors as production code (see ADR-0007 §Decision.5)."""
    footnote = etree.Element(_W_FOOTNOTE)
    footnote.set(_W_TYPE, kind)
    footnote.set(_W_ID, footnote_id)
    paragraph = etree.SubElement(footnote, _W_P)
    run = etree.SubElement(paragraph, _W_R)
    run_props = etree.SubElement(run, _W_RPR)
    etree.SubElement(run_props, _W_VERTALIGN).set(_W_VAL, "superscript")
    if kind == "separator":
        etree.SubElement(run, _W_SEPARATOR)
    else:
        etree.SubElement(run, _W_CONTINUATION_SEPARATOR)
    return footnote


def _build_fresh_footnotes_root(new_footnote: etree._Element) -> etree._Element:
    root = etree.Element(_W_FOOTNOTES, nsmap=NSMAP)
    root.append(_boilerplate_footnote("-1", "separator"))
    root.append(_boilerplate_footnote("0", "continuationSeparator"))
    root.append(new_footnote)
    return root


def _find_footnote(footnotes_root: etree._Element, footnote_id: str) -> etree._Element | None:
    for footnote in footnotes_root.findall("w:footnote", namespaces=NSMAP):
        if footnote.get(_W_ID) == footnote_id:
            return footnote
    return None


def _allocate_footnote_id(footnotes_root: etree._Element | None) -> str:
    """One more than the largest integer-parseable, non-negative declared
    id, or `"1"` if there is none - defensively re-checked against every
    existing id's exact spelling (see ADR-0007 §Decision.4)."""
    existing_ids: set[str] = set()
    max_numeric = 0
    if footnotes_root is not None:
        for footnote in footnotes_root.findall("w:footnote", namespaces=NSMAP):
            footnote_id = footnote.get(_W_ID)
            if footnote_id is None:
                continue
            existing_ids.add(footnote_id)
            try:
                value = int(footnote_id)
            except ValueError:
                continue
            if value >= 1:
                max_numeric = max(max_numeric, value)

    candidate = max_numeric + 1
    while str(candidate) in existing_ids:
        candidate += 1
    return str(candidate)


def _append_footnote_reference(paragraph: etree._Element, footnote_id: str) -> None:
    """Append the new `w:footnoteReference` run as the paragraph's last
    child - the "always at the paragraph's end" anchor rule (see
    [specs/add_footnote.md §5](specs/add_footnote.md#5-limitations-non-goals))."""
    run = etree.SubElement(paragraph, _W_R)
    run_props = etree.SubElement(run, _W_RPR)
    etree.SubElement(run_props, _W_RSTYLE).set(_W_VAL, _FOOTNOTE_REFERENCE_STYLE)
    reference = etree.SubElement(run, _W_FOOTNOTE_REFERENCE)
    reference.set(_W_ID, footnote_id)


def _content_types_declares_footnotes(content_types_root: etree._Element) -> bool:
    return any(
        override.get("PartName") == "/word/footnotes.xml"
        for override in content_types_root.findall(f"{{{_CONTENT_TYPES_NS}}}Override")
    )


def _add_footnotes_content_type(content_types_xml: bytes) -> bytes:
    root = parse_xml(content_types_xml, part_name=CONTENT_TYPES_PART)
    if not _content_types_declares_footnotes(root):
        override = etree.SubElement(root, f"{{{_CONTENT_TYPES_NS}}}Override")
        override.set("PartName", "/word/footnotes.xml")
        override.set("ContentType", _FOOTNOTES_CONTENT_TYPE)
    return _serialize(root)


def _rels_declares_footnotes(rels_root: etree._Element) -> bool:
    return any(
        relationship.get("Type") == _FOOTNOTES_RELATIONSHIP_TYPE
        for relationship in rels_root.findall(f"{{{_RELS_NS}}}Relationship")
    )


def _next_relationship_id(rels_root: etree._Element) -> str:
    existing_ids = {
        relationship.get("Id") for relationship in rels_root.findall(f"{{{_RELS_NS}}}Relationship")
    }
    max_numeric = 0
    for relationship_id in existing_ids:
        if relationship_id and relationship_id.startswith("rId") and relationship_id[3:].isdigit():
            max_numeric = max(max_numeric, int(relationship_id[3:]))
    candidate = max_numeric + 1
    while f"rId{candidate}" in existing_ids:
        candidate += 1
    return f"rId{candidate}"


def _add_footnotes_relationship(rels_xml: bytes | None) -> bytes:
    """`rels_xml` is `None` when `word/_rels/document.xml.rels` doesn't
    exist at all - valid OPC for a source part with zero relationships; a
    fresh one is created, per ADR-0007 §Decision.5."""
    root = (
        parse_xml(rels_xml, part_name=DOCUMENT_RELS_PART)
        if rels_xml is not None
        else parse_xml(
            f'<Relationships xmlns="{_RELS_NS}"></Relationships>'.encode(),
            part_name=DOCUMENT_RELS_PART,
        )
    )
    if not _rels_declares_footnotes(root):
        relationship = etree.SubElement(root, f"{{{_RELS_NS}}}Relationship")
        relationship.set("Id", _next_relationship_id(root))
        relationship.set("Type", _FOOTNOTES_RELATIONSHIP_TYPE)
        relationship.set("Target", "footnotes.xml")
    return _serialize(root)


def add_footnote_to_document(
    docx_path: Path,
    paragraph_index: int,
    content: str,
    *,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> str:
    """Attach a new footnote to a paragraph in a `.docx`'s body.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots; this function
            does not perform any sandboxing itself.
        paragraph_index: 0-based index, among top-level body paragraphs, of
            the paragraph the new footnote reference is appended to.
        content: The new footnote's text content. May be empty. A `"\\n"`
            inside it starts a new paragraph within the footnote.
        max_size_bytes: Reject the file if it exceeds this, before it is opened.

    Returns:
        The new footnote's `w:id`, as a string - always server-allocated
        (see `_allocate_footnote_id`), never derived from `content`.

    Raises:
        FootnoteEditError: If `paragraph_index` does not reference an
            existing paragraph.
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, is missing/has a
            malformed `word/document.xml`, has a malformed
            `word/footnotes.xml`, or - only when `word/footnotes.xml` is
            absent - is also missing `[Content_Types].xml`.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        document_xml = read_part(archive, DOCUMENT_PART)
        footnotes_xml = read_optional_part(archive, FOOTNOTES_PART)
        content_types_xml = read_optional_part(archive, CONTENT_TYPES_PART)
        rels_xml = read_optional_part(archive, DOCUMENT_RELS_PART)

    document_root = parse_xml(document_xml, part_name=DOCUMENT_PART)
    body = get_body(document_root)
    paragraphs = body.findall("w:p", namespaces=NSMAP)
    if not (0 <= paragraph_index < len(paragraphs)):
        raise FootnoteEditError("paragraph_index references a paragraph that does not exist")

    footnotes_root = (
        parse_xml(footnotes_xml, part_name=FOOTNOTES_PART) if footnotes_xml is not None else None
    )
    new_id = _allocate_footnote_id(footnotes_root)
    new_footnote = _build_footnote_element(new_id, content)

    parts_to_write: dict[str, bytes] = {}
    if footnotes_root is not None:
        footnotes_root.append(new_footnote)
        parts_to_write[FOOTNOTES_PART] = _serialize(footnotes_root)
    else:
        if content_types_xml is None:
            raise InvalidDocumentError(f"not a valid .docx file: missing {CONTENT_TYPES_PART}")
        parts_to_write[FOOTNOTES_PART] = _serialize(_build_fresh_footnotes_root(new_footnote))
        parts_to_write[CONTENT_TYPES_PART] = _add_footnotes_content_type(content_types_xml)
        parts_to_write[DOCUMENT_RELS_PART] = _add_footnotes_relationship(rels_xml)

    _append_footnote_reference(paragraphs[paragraph_index], new_id)
    parts_to_write[DOCUMENT_PART] = _serialize(document_root)

    atomic_write_parts(docx_path, parts_to_write, max_size_bytes=max_size_bytes)
    return new_id


def edit_footnote_in_document(
    docx_path: Path,
    footnote_id: str,
    content: str,
    *,
    expected_content: str | None = None,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> str:
    """Replace an existing footnote's content in a `.docx`'s `word/footnotes.xml`.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots; this function
            does not perform any sandboxing itself.
        footnote_id: The `w:id` of the footnote to edit. Must not be `"-1"`
            or `"0"` (Word's reserved boilerplate ids).
        content: The footnote's new text content. May be empty. A `"\\n"`
            inside it starts a new paragraph within the footnote, same
            convention as `add_footnote_to_document`.
        expected_content: If given, must equal the footnote's current
            content (as `docx_mcp.footnotes.render_footnote_content` would
            report it) exactly, or the call fails as stale. `None` (default)
            skips this check.
        max_size_bytes: Reject the file if it exceeds this, before it is opened.

    Returns:
        The footnote's content exactly as it was immediately before this
        call overwrote it.

    Raises:
        FootnoteEditError: If `footnote_id` is `"-1"`/`"0"`,
            `word/footnotes.xml` is absent, `footnote_id` is not declared in
            it, or `expected_content` is given and does not match.
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, is missing/has a
            malformed `word/document.xml`, or has a malformed
            `word/footnotes.xml`.
    """
    if footnote_id in _RESERVED_FOOTNOTE_IDS:
        raise FootnoteEditError(f"cannot edit Word's reserved boilerplate footnote {footnote_id}")

    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        read_part(archive, DOCUMENT_PART)  # baseline: this must be a valid .docx at all
        footnotes_xml = read_optional_part(archive, FOOTNOTES_PART)

    if footnotes_xml is None:
        raise FootnoteEditError(f"footnote {footnote_id} not found: {FOOTNOTES_PART} is absent")

    footnotes_root = parse_xml(footnotes_xml, part_name=FOOTNOTES_PART)
    footnote = _find_footnote(footnotes_root, footnote_id)
    if footnote is None:
        raise FootnoteEditError(f"footnote {footnote_id} not found")

    previous_content = render_footnote_content(footnote)
    if expected_content is not None and previous_content != expected_content:
        raise FootnoteEditError(
            "stale footnote_id: the document no longer contains expected_content there "
            "- re-run get_footnotes"
        )

    for child in list(footnote):
        footnote.remove(child)
    for paragraph in _build_footnote_paragraphs(content):
        footnote.append(paragraph)

    atomic_write_part(
        docx_path, FOOTNOTES_PART, _serialize(footnotes_root), max_size_bytes=max_size_bytes
    )
    return previous_content
