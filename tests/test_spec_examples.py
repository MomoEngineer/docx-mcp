"""Verifies that the concrete "Examples" section documented in every
`src/docx_mcp/specs/<tool>.md` (and mirrored in the tool's docstring in
`server.py`) actually produces the response it claims.

Added in Phase 7 ([Roadmap.md](../Roadmap.md), "Security Hardening") alongside
the "review of every tool description against the smell categories"
(purpose, parameters, limitations, examples - see
[CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), principle 5):
an example that is merely written down can silently drift from the
implementation as the code evolves; this file makes that drift a failing
test instead of a documentation bug nobody notices. Each test below mirrors
one tool's "## 9. Examples" section (and, in prose, its docstring's
"Example:" block) verbatim - if either is changed without updating the
other, or without updating this test, the mismatch is caught here.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.server import mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


@pytest.mark.anyio
async def test_read_document_example_matches_spec(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/read_document.md §9 / server.py docstring."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("read_document", {"path": str(minimal_docx)})

    assert result.structured_content == {
        "text": (
            "# Introduction\n"
            "This is the first paragraph of the document.\n"
            "This paragraph has a footnote reference.[^1]\n"
            "## Background\n"
            "\n"
            "Final paragraph."
        )
    }


@pytest.mark.anyio
async def test_get_structure_example_matches_spec(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/get_structure.md §9 / server.py docstring - primary example (minimal.docx)."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_structure", {"path": str(minimal_docx)})

    assert result.structured_content == {
        "paragraphs": [
            {
                "paragraph_index": 0,
                "text": "Introduction",
                "style_id": "Heading1",
                "heading_level": 1,
            },
            {
                "paragraph_index": 1,
                "text": "This is the first paragraph of the document.",
                "style_id": None,
                "heading_level": None,
            },
            {
                "paragraph_index": 2,
                "text": "This paragraph has a footnote reference.",
                "style_id": None,
                "heading_level": None,
            },
            {
                "paragraph_index": 3,
                "text": "Background",
                "style_id": "Heading2",
                "heading_level": 2,
            },
            {"paragraph_index": 4, "text": "", "style_id": None, "heading_level": None},
            {
                "paragraph_index": 5,
                "text": "Final paragraph.",
                "style_id": None,
                "heading_level": None,
            },
        ],
        "toc": [
            {"paragraph_index": 0, "level": 1, "text": "Introduction"},
            {"paragraph_index": 3, "level": 2, "text": "Background"},
        ],
        "tables": [],
        "footnotes": [{"id": "1", "paragraph_index": 2}],
        "total_paragraphs": 6,
    }


@pytest.mark.anyio
async def test_read_document_ranged_example_matches_spec(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/read_document.md §9 - scoped-range example (ADR-0008)."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "read_document",
            {"path": str(minimal_docx), "start_paragraph": 3, "end_paragraph": 99},
        )

    assert result.structured_content == {"text": "## Background\n\nFinal paragraph."}


@pytest.mark.anyio
async def test_get_structure_ranged_example_matches_spec(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/get_structure.md §9 - scoped-range example (ADR-0008)."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "get_structure",
            {"path": str(minimal_docx), "start_paragraph": 3, "end_paragraph": 99},
        )

    assert result.structured_content == {
        "paragraphs": [
            {
                "paragraph_index": 3,
                "text": "Background",
                "style_id": "Heading2",
                "heading_level": 2,
            },
            {"paragraph_index": 4, "text": "", "style_id": None, "heading_level": None},
            {
                "paragraph_index": 5,
                "text": "Final paragraph.",
                "style_id": None,
                "heading_level": None,
            },
        ],
        "toc": [{"paragraph_index": 3, "level": 2, "text": "Background"}],
        "tables": [],
        "footnotes": [],
        "total_paragraphs": 6,
    }


@pytest.mark.anyio
async def test_get_structure_example_table_rows_matches_spec(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/get_structure.md §9 - secondary example (structured.docx's merged-cell table)."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_structure", {"path": str(structured_docx)})

    assert result.structured_content["tables"][0]["rows"] == [
        ["Header A", "Header B"],
        ["Value 1", "Value 2"],
        ["Merged Row"],
    ]


@pytest.mark.anyio
async def test_get_metadata_example_matches_spec(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/get_metadata.md §9 / server.py docstring."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_metadata", {"path": str(structured_docx)})

    assert result.structured_content == {
        "title": "Structured Test Document",
        "author": "Test Author",
        "created": "2024-01-01T00:00:00Z",
        "modified": "2024-06-15T12:30:00Z",
        "word_count": 29,
    }


@pytest.mark.anyio
async def test_get_footnotes_example_matches_spec(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/get_footnotes.md §9 / server.py docstring - including the leading space."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(minimal_docx)})

    assert result.structured_content == {
        "footnotes": [{"id": "1", "paragraph_index": 2, "content": " This is a footnote."}]
    }


@pytest.mark.anyio
async def test_find_text_example_matches_spec(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/find_text.md §9 / server.py docstring."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "find_text", {"path": str(run_split_docx), "search_text": "receive"}
        )

    assert result.structured_content == {
        "matches": [
            {
                "location": {"paragraph_index": 0, "start_offset": 7, "end_offset": 14},
                "matched_text": "receive",
                "paragraph_text": "Please receive this message.",
            }
        ]
    }


@pytest.mark.anyio
async def test_replace_text_location_mode_example_matches_spec(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/replace_text.md §9 / server.py docstring - location-mode example."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "replace_text",
            {
                "path": str(run_split_docx),
                "search_text": "receive",
                "replacement_text": "REPLIED",
                "location": {"paragraph_index": 0, "start_offset": 7, "end_offset": 14},
            },
        )
        read_back = await client.call_tool("read_document", {"path": str(run_split_docx)})

    assert result.structured_content == {"replacements_made": 1}
    assert "Please REPLIED this message." in read_back.structured_content["text"]


@pytest.mark.anyio
async def test_replace_text_global_mode_example_matches_spec(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/replace_text.md §9 / server.py docstring - global-mode example."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "replace_text",
            {"path": str(run_split_docx), "search_text": "test", "replacement_text": "exam"},
        )

    assert result.structured_content == {"replacements_made": 4}


@pytest.mark.anyio
async def test_insert_paragraph_example_matches_spec(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/insert_paragraph.md §9 / server.py docstring."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "insert_paragraph",
            {
                "path": str(paragraph_edits_docx),
                "text": "Added by the agent.",
                "after_paragraph_index": 1,
            },
        )

    assert result.structured_content == {"paragraph_index": 2}


@pytest.mark.anyio
async def test_delete_paragraph_example_matches_spec(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/delete_paragraph.md §9 / server.py docstring."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "delete_paragraph",
            {
                "path": str(paragraph_edits_docx),
                "paragraph_index": 3,
                "expected_text": "Second body paragraph.",
            },
        )

    assert result.structured_content == {"deleted_text": "Second body paragraph."}


@pytest.mark.anyio
async def test_add_footnote_from_scratch_example_matches_spec(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/add_footnote.md §9 / server.py docstring - primary (from-scratch) example."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "add_footnote",
            {
                "path": str(structured_docx),
                "paragraph_index": 0,
                "content": "Source: internal Q1 report.",
            },
        )

    assert result.structured_content == {"footnote_id": "1"}


@pytest.mark.anyio
async def test_add_footnote_append_example_matches_spec(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/add_footnote.md §9 - secondary (append-path) example."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "add_footnote",
            {"path": str(footnotes_docx), "paragraph_index": 5, "content": "x"},
        )

    assert result.structured_content == {"footnote_id": "4"}


@pytest.mark.anyio
async def test_edit_footnote_example_matches_spec(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """specs/edit_footnote.md §9 / server.py docstring - including the leading space in
    expected_content, matching get_footnotes' documented content exactly."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {
                "path": str(minimal_docx),
                "footnote_id": "1",
                "content": "This is a revised footnote, with a source cited.",
                "expected_content": " This is a footnote.",
            },
        )

    assert result.structured_content == {"previous_content": " This is a footnote."}
