"""Cross-tool consistency checks for `insert_paragraph`/`delete_paragraph`.

Per [docs/testing.md §3](../../docs/testing.md#3-integration-tests-cross-tool):
this is the direct proof of Roadmap Phase 5's requirement that a paragraph
edit is correctly visible (or, for a deletion, correctly absent) to every
read tool afterward, not just to `paragraph_edit.py`'s own internal checks.
"""

from __future__ import annotations

from pathlib import Path

from docx_mcp.document import extract_text
from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.paragraph_edit import delete_paragraph_from_document, insert_paragraph_in_document
from docx_mcp.structure import get_document_structure
from docx_mcp.text_edit import find_text_matches


def test_inserted_paragraph_is_visible_to_read_document_and_get_structure(
    paragraph_edits_docx: Path,
) -> None:
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Freshly inserted sentence.", after_paragraph_index=1
    )

    text_lines = extract_text(paragraph_edits_docx).split("\n")
    assert text_lines[new_index] == "Freshly inserted sentence."

    structure = get_document_structure(paragraph_edits_docx)
    assert structure.paragraphs[new_index].text == "Freshly inserted sentence."
    assert structure.paragraphs[new_index].heading_level is None


def test_inserted_heading_is_visible_with_hash_marker_and_in_the_toc(
    paragraph_edits_docx: Path,
) -> None:
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "New Subsection", after_paragraph_index=1, heading_level=2
    )

    text_lines = extract_text(paragraph_edits_docx).split("\n")
    assert text_lines[new_index] == "## New Subsection"

    structure = get_document_structure(paragraph_edits_docx)
    assert structure.paragraphs[new_index].heading_level == 2
    assert any(
        entry.paragraph_index == new_index and entry.text == "New Subsection"
        for entry in structure.toc
    )


def test_inserted_paragraph_text_is_findable_via_find_text(paragraph_edits_docx: Path) -> None:
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "A uniquely searchable phrase.", after_paragraph_index=0
    )

    matches = find_text_matches(paragraph_edits_docx, "uniquely searchable")
    assert len(matches) == 1
    assert matches[0].location.paragraph_index == new_index


def test_deleted_paragraph_disappears_from_read_document_and_get_structure(
    paragraph_edits_docx: Path,
) -> None:
    before_structure = get_document_structure(paragraph_edits_docx)
    before_count = len(before_structure.paragraphs)

    delete_paragraph_from_document(paragraph_edits_docx, 1)

    text_lines = extract_text(paragraph_edits_docx).split("\n")
    assert "First body paragraph." not in text_lines

    after_structure = get_document_structure(paragraph_edits_docx)
    assert len(after_structure.paragraphs) == before_count - 1
    assert all(p.text != "First body paragraph." for p in after_structure.paragraphs)


def test_deleting_a_footnote_anchoring_paragraph_removes_it_from_get_footnotes(
    footnotes_docx: Path,
) -> None:
    """`footnotes.docx` paragraph 2 anchors footnote ids 1 and 2; paragraph 4
    anchors id 3 - see `tests/fixtures/generate_fixtures.py`. Deleting
    paragraph 2 must make `get_footnotes` stop reporting ids 1/2, while id 3
    (anchored in an unrelated, unaffected paragraph) keeps resolving
    correctly - the direct proof that `get_footnotes`, not just
    `paragraph_edit.py`'s own tests, agrees the anchors are gone."""
    before = get_document_footnotes(footnotes_docx)
    assert {f.id for f in before} == {"1", "2", "3"}

    delete_paragraph_from_document(footnotes_docx, 2)

    after = get_document_footnotes(footnotes_docx)
    assert {f.id for f in after} == {"3"}
    assert after[0].content == " Third footnote."
