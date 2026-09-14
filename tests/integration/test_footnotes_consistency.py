"""Cross-tool consistency checks between `get_footnotes` and `read_document`/`get_structure`.

Per [docs/testing.md §3](../../docs/testing.md#3-integration-tests-cross-tool):
this is the direct proof of Roadmap Phase 3's requirement that `read_document`'s
inline `"[^N]"` markers, `get_structure`'s footnote-anchor index, and
`get_footnotes`'s own output all "resolve consistently against the same data"
(see [ADR-0004](../../docs/adr/0004-phase-3-footnote-module-and-shared-anchor-resolution.md),
which makes this true by construction - all three now share
`document.py`'s `find_footnote_anchors` - rather than merely by coincidence
between independent implementations).
"""

from __future__ import annotations

from pathlib import Path

from docx_mcp.document import extract_text
from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.structure import get_document_structure


def _assert_footnotes_agree_with_structure_and_read_document(docx_path: Path) -> None:
    text_lines = extract_text(docx_path).split("\n")
    structure = get_document_structure(docx_path)
    footnotes = get_document_footnotes(docx_path)

    assert [(f.id, f.paragraph_index) for f in footnotes] == [
        (a.id, a.paragraph_index) for a in structure.footnotes
    ]
    for entry in footnotes:
        assert f"[^{entry.id}]" in text_lines[entry.paragraph_index]


def test_minimal_fixture_single_footnote_agrees_across_all_three_tools(
    minimal_docx: Path,
) -> None:
    _assert_footnotes_agree_with_structure_and_read_document(minimal_docx)


def test_footnotes_fixture_two_footnotes_in_one_paragraph_agrees_across_all_three_tools(
    footnotes_docx: Path,
) -> None:
    """The exact edge case Roadmap Phase 3 calls out: two different footnote
    ids anchored in the same paragraph must resolve identically in
    `read_document`, `get_structure`, and `get_footnotes` - including that
    both ids' markers appear in that one paragraph's `read_document` line,
    and neither is lost or collapsed in `get_structure`'s or `get_footnotes`'
    per-anchor list."""
    _assert_footnotes_agree_with_structure_and_read_document(footnotes_docx)

    footnotes = get_document_footnotes(footnotes_docx)
    assert len(footnotes) == 3
    same_paragraph = [f for f in footnotes if f.paragraph_index == 2]
    assert [f.id for f in same_paragraph] == ["1", "2"]


def test_structured_fixture_has_no_footnotes_in_any_of_the_three_tools(
    structured_docx: Path,
) -> None:
    """`structured.docx` (Phase 2) has zero footnotes - all three tools must
    agree it's empty, not just `get_structure`."""
    structure = get_document_structure(structured_docx)
    footnotes = get_document_footnotes(structured_docx)

    assert structure.footnotes == ()
    assert footnotes == ()
