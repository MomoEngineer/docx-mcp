"""Cross-tool consistency checks between `find_text` and `replace_text`.

Per [docs/testing.md §3](../../docs/testing.md#3-integration-tests-cross-tool):
this is the direct proof of Roadmap Phase 4's requirement that `find_text`
returns "stable, addressable locations that `replace_text` can act on
unambiguously" - and the direct proof of the stale-location safeguard
[ADR-0005](../../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md)
adds to close the race/hallucination window between the two calls.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import extract_text
from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.structure import get_document_structure
from docx_mcp.text_edit import TextEditError, find_text_matches, replace_text_in_document


def test_find_text_location_is_valid_replace_text_input(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "receive")
    assert len(matches) == 1

    count = replace_text_in_document(
        run_split_docx, "receive", "REPLIED", location=matches[0].location
    )

    assert count == 1
    assert find_text_matches(run_split_docx, "REPLIED")


def test_reusing_a_location_after_an_intervening_edit_is_rejected_as_stale(
    run_split_docx: Path,
) -> None:
    location = find_text_matches(run_split_docx, "receive")[0].location

    replace_text_in_document(run_split_docx, "receive", "REPLIED", location=location)

    with pytest.raises(TextEditError, match="stale"):
        replace_text_in_document(run_split_docx, "receive", "AGAIN", location=location)


def test_read_document_get_structure_get_footnotes_stay_consistent_after_a_write(
    minimal_docx: Path,
) -> None:
    """Extends the Phase 3 three-way consistency check
    (`tests/integration/test_footnotes_consistency.py`) across a write: after
    editing text next to `minimal.docx`'s footnote anchor, all three read
    tools must still agree with each other, not just with themselves."""
    replace_text_in_document(minimal_docx, "footnote reference", "annotation marker")

    text_lines = extract_text(minimal_docx).split("\n")
    structure = get_document_structure(minimal_docx)
    footnotes = get_document_footnotes(minimal_docx)

    assert [(f.id, f.paragraph_index) for f in footnotes] == [
        (a.id, a.paragraph_index) for a in structure.footnotes
    ]
    for entry in footnotes:
        assert f"[^{entry.id}]" in text_lines[entry.paragraph_index]
    assert "annotation marker" in text_lines[2]
