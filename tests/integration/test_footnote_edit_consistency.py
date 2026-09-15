"""Cross-tool consistency checks for `add_footnote`/`edit_footnote`.

Per [docs/testing.md §3](../../docs/testing.md#3-integration-tests-cross-tool):
this is the direct proof of Roadmap Phase 6's requirement that a newly added
footnote is discoverable and correctly anchored via `get_footnotes`, and that
an edited footnote's content updates without breaking its anchor or id - not
just that `footnote_edit.py`'s own unit tests agree with themselves.
"""

from __future__ import annotations

from pathlib import Path

from docx_mcp.document import extract_text
from docx_mcp.footnote_edit import add_footnote_to_document, edit_footnote_in_document
from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.structure import get_document_structure


def test_added_footnote_resolves_consistently_across_read_document_get_structure_get_footnotes(
    structured_docx: Path,
) -> None:
    new_id = add_footnote_to_document(structured_docx, 1, "A brand new footnote.")

    text_lines = extract_text(structured_docx).split("\n")
    assert f"[^{new_id}]" in text_lines[1]

    structure = get_document_structure(structured_docx)
    assert any(f.id == new_id and f.paragraph_index == 1 for f in structure.footnotes)

    footnotes = get_document_footnotes(structured_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.paragraph_index == 1
    assert added.content == "A brand new footnote."


def test_added_footnote_on_a_document_that_already_has_footnotes_does_not_disturb_existing_ones(
    footnotes_docx: Path,
) -> None:
    before = get_document_footnotes(footnotes_docx)

    new_id = add_footnote_to_document(footnotes_docx, 5, "Fourth footnote.")

    after = get_document_footnotes(footnotes_docx)
    assert after[: len(before)] == before
    added = next(f for f in after if f.id == new_id)
    assert added.paragraph_index == 5
    assert added.content == "Fourth footnote."

    structure = get_document_structure(footnotes_docx)
    assert any(f.id == new_id and f.paragraph_index == 5 for f in structure.footnotes)


def test_edited_footnote_content_resolves_consistently_without_moving_its_anchor(
    footnotes_docx: Path,
) -> None:
    before_structure = get_document_structure(footnotes_docx)
    before_anchor = next(f for f in before_structure.footnotes if f.id == "1")

    edit_footnote_in_document(footnotes_docx, "1", "Rewritten footnote content.")

    after_structure = get_document_structure(footnotes_docx)
    after_anchor = next(f for f in after_structure.footnotes if f.id == "1")
    assert after_anchor.paragraph_index == before_anchor.paragraph_index

    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == "Rewritten footnote content."

    text_lines = extract_text(footnotes_docx).split("\n")
    assert "[^1]" in text_lines[after_anchor.paragraph_index]


def test_multiline_added_footnote_content_is_consistent_with_get_footnotes_join_convention(
    structured_docx: Path,
) -> None:
    new_id = add_footnote_to_document(structured_docx, 0, "First line.\nSecond line.")

    footnotes = get_document_footnotes(structured_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.content == "First line.\nSecond line."


def test_no_regression_in_existing_footnotes_after_add_and_edit(footnotes_docx: Path) -> None:
    """`footnotes.docx`'s existing ids 1-3 (Phase 3's own fixture, already
    proven consistent by `tests/integration/test_footnotes_consistency.py`)
    must still resolve correctly after an unrelated add and edit."""
    add_footnote_to_document(footnotes_docx, 1, "Unrelated new footnote.")
    edit_footnote_in_document(footnotes_docx, "3", "Rewritten third footnote.")

    footnotes = get_document_footnotes(footnotes_docx)
    assert next(f for f in footnotes if f.id == "1").content == " First footnote."
    assert next(f for f in footnotes if f.id == "2").content == " Second footnote."
    assert next(f for f in footnotes if f.id == "3").content == "Rewritten third footnote."
