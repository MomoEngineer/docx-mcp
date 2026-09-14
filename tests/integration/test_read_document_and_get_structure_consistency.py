"""Cross-tool consistency checks between `read_document` and `get_structure`.

Per [docs/testing.md §3](../../docs/testing.md#3-integration-tests-cross-tool):
integration tests for chains/interactions between tools that build on each
other or expose the same underlying data two ways. `read_document`'s inline
`"# "`/`"[^N]"` markers and `get_structure`'s `heading_level`/`footnotes`
index are two independently-computed views over the same `word/document.xml`
- these tests prove they agree everywhere the specs promise they must, and
pin the one place they are documented to legitimately *disagree* (custom
heading styles - see
[specs/read_document.md §5](../../src/docx_mcp/specs/read_document.md#5-limitations-non-goals)
and
[specs/get_structure.md §5](../../src/docx_mcp/specs/get_structure.md#5-limitations-non-goals))
so that divergence stays an intentional, tested contract rather than a
silent accident either tool's implementation could drift into.
"""

from __future__ import annotations

from pathlib import Path

from docx_mcp.document import extract_text
from docx_mcp.metadata import get_document_metadata
from docx_mcp.structure import get_document_structure


def test_non_heading_paragraph_text_is_identical_in_both_tools(minimal_docx: Path) -> None:
    text_lines = extract_text(minimal_docx).split("\n")
    structure = get_document_structure(minimal_docx)

    # Paragraph 1 ("This is the first paragraph...") carries no heading/footnote
    # marker in either tool, so the two independently-assembled strings must
    # be byte-identical - the strongest form of this consistency check.
    assert (
        text_lines[1]
        == structure.paragraphs[1].text
        == ("This is the first paragraph of the document.")
    )


def test_footnote_anchor_paragraph_matches_the_inline_marker_position(
    minimal_docx: Path,
) -> None:
    """`get_structure`'s footnote-anchor index must point at the same
    paragraph where `read_document` places the `"[^N]"` marker inline."""
    text_lines = extract_text(minimal_docx).split("\n")
    structure = get_document_structure(minimal_docx)

    assert len(structure.footnotes) == 1
    anchor = structure.footnotes[0]
    assert f"[^{anchor.id}]" in text_lines[anchor.paragraph_index]


def test_builtin_heading_style_agrees_between_both_tools(structured_docx: Path) -> None:
    """A built-in `HeadingN` style is recognized by both tools' independent
    algorithms - read_document's digit-suffix heuristic and get_structure's
    styles.xml/outlineLvl resolution - and must agree on level and text."""
    text_lines = extract_text(structured_docx).split("\n")
    structure = get_document_structure(structured_docx)

    introduction = structure.paragraphs[0]
    assert introduction.heading_level == 1
    assert text_lines[0] == f"# {introduction.text}" == "# Introduction"


def test_custom_heading_style_is_a_documented_intentional_divergence(
    structured_docx: Path,
) -> None:
    """`structured.docx`'s paragraph 3 uses a custom style ("MySectionHeading")
    based on Heading2. get_structure resolves it correctly via styles.xml
    (heading_level=2); read_document's simpler HeadingN-prefix heuristic does
    not recognize it as a heading at all and renders it as plain text. Both
    are individually correct per their own spec (see specs/read_document.md
    §5 and specs/get_structure.md §5) - this test pins that the divergence
    is real and exists exactly where documented, not somewhere else."""
    text_lines = extract_text(structured_docx).split("\n")
    structure = get_document_structure(structured_docx)

    custom_heading = structure.paragraphs[3]
    assert custom_heading.style_id == "MySectionHeading"
    assert custom_heading.heading_level == 2

    # read_document does NOT prefix this paragraph with "##" - it is
    # indistinguishable from a plain paragraph in read_document's output.
    assert text_lines[3] == custom_heading.text == "Custom styled heading text"


def test_word_count_is_consistent_with_read_document_text_minus_markers(
    minimal_docx: Path,
) -> None:
    """A cheap independent cross-check: word_count (metadata.py) must equal a
    whitespace-token count of get_structure's marker-free paragraph texts -
    both are meant to describe the same underlying content."""
    structure = get_document_structure(minimal_docx)
    metadata = get_document_metadata(minimal_docx)

    expected = len("\n".join(p.text for p in structure.paragraphs).split())
    assert metadata.word_count == expected
