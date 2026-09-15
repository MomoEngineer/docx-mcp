"""Independent-library verification for `insert_paragraph`/`delete_paragraph`.

Every other test in this project verifies a write tool's output using this
project's *own* parsing code (`document.py`/`paragraph_edit.py`'s helpers) -
a blind spot in that code could, in principle, produce a document that looks
correct to our own assertions while being subtly wrong. This module instead
re-opens the edited document with `python-docx`, a completely independent
OOXML library already a dev dependency for fixture generation (see
`tests/fixtures/generate_fixtures.py`), and checks the result through its
own object model (`Document`, `Paragraph.style.name`, `paragraph_format`).
This is deliberately not part of `docs/testing.md §3`'s "cross-tool" category
(no two *docx-mcp* tools are being compared here) - it is an independent-
library sanity check, added during a dedicated post-implementation
verification pass rather than the original spec-first test design.

Only real, full-OPC-package fixtures (`paragraph_edits.docx`) are used here -
the minimal synthetic `word/document.xml`-only archives used elsewhere in
`tests/test_paragraph_edit.py` for edge cases deliberately omit
`[Content_Types].xml`/`_rels` and are not valid OPC packages `python-docx`
(or Word) can open at all; that is an intentional leniency difference
between this project's hardened-but-minimal parser and a full OPC reader,
not a bug.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document

from docx_mcp.paragraph_edit import delete_paragraph_from_document, insert_paragraph_in_document


def test_edited_document_reopens_correctly_in_python_docx(paragraph_edits_docx: Path) -> None:
    insert_paragraph_in_document(
        paragraph_edits_docx, "New Section", after_paragraph_index=1, heading_level=2
    )
    insert_paragraph_in_document(paragraph_edits_docx, "Follow-up text.", after_paragraph_index=2)
    delete_paragraph_from_document(paragraph_edits_docx, 0)
    insert_paragraph_in_document(
        paragraph_edits_docx, "Brand New Title", after_paragraph_index=None, heading_level=1
    )

    document = Document(str(paragraph_edits_docx))

    assert [(p.text, p.style.name) for p in document.paragraphs] == [
        ("Brand New Title", "Heading 1"),
        ("First body paragraph.", "Normal"),
        ("New Section", "Heading 2"),
        ("Follow-up text.", "Normal"),
        ("Details", "Heading 2"),
        ("Second body paragraph.", "Normal"),
        ("Last paragraph.", "Normal"),
    ]


def test_inherited_and_anchor_is_heading_formatting_is_visible_to_python_docx(
    paragraph_edits_docx: Path,
) -> None:
    """Independent confirmation, via python-docx's own `paragraph_format`
    view, of both style-inheritance branches: a plain-text insertion after
    the centered/indented "First body paragraph." must show that same
    formatting; a plain-text insertion after a heading must show neither."""
    insert_paragraph_in_document(
        paragraph_edits_docx, "New Section", after_paragraph_index=1, heading_level=2
    )
    insert_paragraph_in_document(paragraph_edits_docx, "Follow-up text.", after_paragraph_index=2)

    document = Document(str(paragraph_edits_docx))
    by_text = {p.text: p for p in document.paragraphs}

    anchor_format = by_text["First body paragraph."].paragraph_format
    assert anchor_format.alignment is not None
    assert anchor_format.left_indent is not None

    follow_up_format = by_text["Follow-up text."].paragraph_format
    assert follow_up_format.alignment is None
    assert follow_up_format.left_indent is None


def test_all_heading_levels_resolve_to_the_correct_style_name_in_python_docx(
    paragraph_edits_docx: Path,
) -> None:
    """All nine built-in `HeadingN` style ids - not just the two
    (`Heading1`/`Heading2`) already present in the fixture - must resolve
    to Word's own "Heading N" display style when opened independently."""
    anchor_index = 3
    for level in range(1, 10):
        anchor_index = insert_paragraph_in_document(
            paragraph_edits_docx,
            f"Heading level {level}",
            after_paragraph_index=anchor_index,
            heading_level=level,
        )

    document = Document(str(paragraph_edits_docx))
    by_text = {p.text: p.style.name for p in document.paragraphs}

    for level in range(1, 10):
        assert by_text[f"Heading level {level}"] == f"Heading {level}"


def test_deleting_down_to_a_single_paragraph_still_reopens_correctly(
    paragraph_edits_docx: Path,
) -> None:
    for _ in range(4):
        delete_paragraph_from_document(paragraph_edits_docx, 0)

    document = Document(str(paragraph_edits_docx))

    assert [p.text for p in document.paragraphs] == ["Last paragraph."]
