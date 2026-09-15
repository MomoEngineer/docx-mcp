"""`insert_paragraph`/`delete_paragraph`: whole-paragraph structure edits.

Owns both tools' logic (a deliberate, phase-scoped exception to
one-module-per-tool - see
[ADR-0006](../../docs/adr/0006-phase-5-paragraph-edit-module-layout.md)):
both address the same top-level `w:body/w:p` list by the same
`paragraph_index` convention `find_text`/`replace_text`/`get_structure`
already use, and share one "resolve and validate an index against the
current paragraph list" helper.

The style-inheritance rule for `insert_paragraph` (see
[ADR-0006 §Decision.3-4](../../docs/adr/0006-phase-5-paragraph-edit-module-layout.md)
for the full rationale) is fixed here:

- `heading_level` given (1-9): the new paragraph's `w:pPr` is built fresh,
  containing only `<w:pStyle w:val="HeadingN"/>` - independent of whatever
  the anchor paragraph's own formatting is.
- `heading_level` omitted, anchor exists and is *not* recognized as a
  heading (via `document.py`'s built-in-`HeadingN`-id heuristic only, the
  same one `read_document` already uses - deliberately not `structure.py`'s
  fuller `styles.xml`/`basedOn`-chain resolution, so this write path never
  needs to open `word/styles.xml`): the anchor's entire `w:pPr` is deep-copied
  onto the new paragraph verbatim, **except** any `w:sectPr` child, which is
  always stripped from the copy (see ADR-0006 §Decision.8): `w:sectPr` marks
  a section boundary, not stylistic formatting, and copying it verbatim
  would duplicate the anchor's section properties onto the new paragraph
  too, silently creating an unintended extra section break. If stripping
  `w:sectPr` leaves the copied `w:pPr` with no children at all, no `w:pPr`
  is attached (equivalent to the anchor never having had one beyond its
  section marker).
- `heading_level` omitted, and the anchor either does not exist (empty
  document) or *is* recognized as a heading: the new paragraph gets no
  `w:pPr` at all - specifically so inserting plain body text right after a
  heading does not silently produce a second heading.

`delete_paragraph` removes the `w:p` element as a whole, including any
footnote reference it contains - the corresponding `word/footnotes.xml`
entry is deliberately left orphaned rather than cleaned up (deferred to
Phase 6, see ADR-0006 Decision 5) - and refuses to delete a paragraph whose
`w:pPr` carries a `w:sectPr` (required section properties would otherwise be
lost, see ADR-0006 Decision 6).
"""

from __future__ import annotations

import copy
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
)
from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    atomic_write_part,
    parse_xml,
    read_part,
    validate_and_open,
)

__all__ = [
    "ParagraphEditError",
    "delete_paragraph_from_document",
    "insert_paragraph_in_document",
]

_W_P = f"{{{WORD_NS}}}p"
_W_R = f"{{{WORD_NS}}}r"
_W_T = f"{{{WORD_NS}}}t"
_W_PPR = f"{{{WORD_NS}}}pPr"
_W_PSTYLE = f"{{{WORD_NS}}}pStyle"
_W_VAL = f"{{{WORD_NS}}}val"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_MIN_HEADING_LEVEL = 1
_MAX_HEADING_LEVEL = 9


class ParagraphEditError(Exception):
    """Raised for an `insert_paragraph`/`delete_paragraph`-specific error.

    See [specs/insert_paragraph.md §7](specs/insert_paragraph.md#7-error-behavior)
    and [specs/delete_paragraph.md §7](specs/delete_paragraph.md#7-error-behavior)
    for the full error table.
    """


def _load_body(docx_path: Path, max_size_bytes: int) -> tuple[etree._Element, etree._Element]:
    """Open, validate, and parse `docx_path`, returning `(document_root, body)`."""
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)
    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    return root, get_body(root)


def _write_document(docx_path: Path, root: etree._Element, max_size_bytes: int) -> None:
    new_document_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    atomic_write_part(docx_path, DOCUMENT_PART, new_document_xml, max_size_bytes=max_size_bytes)


def _is_heading(paragraph: etree._Element) -> bool:
    """Whether `paragraph` is recognized as a heading, per the built-in-id
    heuristic only (see module docstring)."""
    return heading_level_from_style_prefix(paragraph_style_id(paragraph)) is not None


def _require_valid_index(index: int, paragraph_count: int, *, param_name: str) -> None:
    """Shared by `insert_paragraph_in_document` and
    `delete_paragraph_from_document` (see module docstring): both resolve
    and validate their index parameter against the current top-level
    paragraph count through this one check, so the two tools cannot drift
    into disagreeing on what counts as a valid `paragraph_index`."""
    if not (0 <= index < paragraph_count):
        raise ParagraphEditError(f"{param_name} references a paragraph that does not exist")


def _build_new_paragraph(
    text: str, *, anchor: etree._Element | None, heading_level: int | None
) -> etree._Element:
    new_paragraph = etree.Element(_W_P)

    if heading_level is not None:
        ppr = etree.SubElement(new_paragraph, _W_PPR)
        style = etree.SubElement(ppr, _W_PSTYLE)
        style.set(_W_VAL, f"Heading{heading_level}")
    elif anchor is not None and not _is_heading(anchor):
        anchor_ppr = anchor.find("w:pPr", namespaces=NSMAP)
        if anchor_ppr is not None:
            copied_ppr = copy.deepcopy(anchor_ppr)
            # w:sectPr marks a section boundary, not stylistic formatting -
            # copying it onto the new paragraph would duplicate the anchor's
            # section properties and silently create an unintended extra
            # section break (see ADR-0006 §Decision.8).
            for section_properties in copied_ppr.findall("w:sectPr", namespaces=NSMAP):
                copied_ppr.remove(section_properties)
            if len(copied_ppr) > 0:
                new_paragraph.append(copied_ppr)

    if text:
        run = etree.SubElement(new_paragraph, _W_R)
        run_text = etree.SubElement(run, _W_T)
        run_text.text = text
        run_text.set(_XML_SPACE, "preserve")

    return new_paragraph


def insert_paragraph_in_document(
    docx_path: Path,
    text: str,
    *,
    after_paragraph_index: int | None = None,
    heading_level: int | None = None,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> int:
    """Insert a new paragraph into a `.docx`'s body, inheriting its context's style.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots; this function
            does not perform any sandboxing itself.
        text: The new paragraph's plain text content. May be empty (a valid,
            run-less paragraph).
        after_paragraph_index: 0-based index of the paragraph to insert
            immediately after. `None` (default) inserts at the very start of
            the body.
        heading_level: `1`-`9` to make the new paragraph a heading at that
            level; `None` (default) for plain body text, whose style is
            derived from context (see module docstring).
        max_size_bytes: Reject the file if it exceeds this, before it is opened.

    Returns:
        The 0-based `paragraph_index` the new paragraph now occupies.

    Raises:
        ParagraphEditError: If `after_paragraph_index` does not reference an
            existing paragraph, or `heading_level` is not in `1..9`.
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, or is missing/has a
            malformed `word/document.xml`.
    """
    if heading_level is not None and not (
        _MIN_HEADING_LEVEL <= heading_level <= _MAX_HEADING_LEVEL
    ):
        raise ParagraphEditError(
            f"heading_level must be between {_MIN_HEADING_LEVEL} and {_MAX_HEADING_LEVEL}"
        )

    root, body = _load_body(docx_path, max_size_bytes)
    paragraphs = body.findall("w:p", namespaces=NSMAP)

    if after_paragraph_index is not None:
        _require_valid_index(
            after_paragraph_index, len(paragraphs), param_name="after_paragraph_index"
        )

    anchor = paragraphs[after_paragraph_index] if after_paragraph_index is not None else None
    if anchor is None and paragraphs:
        anchor = paragraphs[0]

    new_paragraph = _build_new_paragraph(text, anchor=anchor, heading_level=heading_level)

    if after_paragraph_index is not None:
        paragraphs[after_paragraph_index].addnext(new_paragraph)
        new_index = after_paragraph_index + 1
    elif paragraphs:
        paragraphs[0].addprevious(new_paragraph)
        new_index = 0
    else:
        body_sectpr = body.find("w:sectPr", namespaces=NSMAP)
        if body_sectpr is not None:
            body_sectpr.addprevious(new_paragraph)
        else:
            body.append(new_paragraph)
        new_index = 0

    _write_document(docx_path, root, max_size_bytes)
    return new_index


def delete_paragraph_from_document(
    docx_path: Path,
    paragraph_index: int,
    *,
    expected_text: str | None = None,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> str:
    """Remove a paragraph from a `.docx`'s body by its `paragraph_index`.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots; this function
            does not perform any sandboxing itself.
        paragraph_index: 0-based index of the paragraph to remove.
        expected_text: If given, must equal the target paragraph's current
            plain text exactly, or the call fails as stale. `None` (default)
            skips this check.
        max_size_bytes: Reject the file if it exceeds this, before it is opened.

    Returns:
        The removed paragraph's plain text (as it was immediately before
        deletion).

    Raises:
        ParagraphEditError: If `paragraph_index` does not reference an
            existing paragraph, `expected_text` is given and does not match,
            or the target paragraph's `w:pPr` contains a `w:sectPr`.
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, or is missing/has a
            malformed `word/document.xml`.
    """
    root, body = _load_body(docx_path, max_size_bytes)
    paragraphs = body.findall("w:p", namespaces=NSMAP)

    _require_valid_index(paragraph_index, len(paragraphs), param_name="paragraph_index")

    target = paragraphs[paragraph_index]
    deleted_text = paragraph_plain_text(target)

    if expected_text is not None and deleted_text != expected_text:
        raise ParagraphEditError(
            "stale paragraph_index: the document no longer contains expected_text there "
            "- re-run get_structure/read_document"
        )

    if target.find("w:pPr/w:sectPr", namespaces=NSMAP) is not None:
        raise ParagraphEditError(
            "cannot delete a paragraph that carries the document's section properties "
            "(w:sectPr) - this would remove required section properties"
        )

    body.remove(target)

    _write_document(docx_path, root, max_size_bytes)
    return deleted_text
