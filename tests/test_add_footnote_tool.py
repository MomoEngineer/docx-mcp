"""Contract and functional tests for the `add_footnote` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/add_footnote.md §8](../src/docx_mcp/specs/add_footnote.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.server import ADD_FOOTNOTE_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_add_footnote_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "add_footnote" in names


@pytest.mark.anyio
async def test_add_footnote_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "add_footnote")
    assert tool.description
    assert "footnote" in tool.description.lower()


@pytest.mark.anyio
async def test_add_footnote_input_schema_declares_expected_parameters(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "add_footnote")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert schema["properties"]["paragraph_index"]["type"] == "integer"
    assert schema["properties"]["content"]["type"] == "string"
    assert set(schema["required"]) == {"path", "paragraph_index", "content"}


@pytest.mark.anyio
async def test_add_footnote_output_schema_declares_footnote_id(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "add_footnote")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"footnote_id"}


@pytest.mark.anyio
async def test_add_footnote_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "add_footnote")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == ADD_FOOTNOTE_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_add_footnote_on_structured_fixture_creates_footnotes_xml(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """`structured.docx` (Phase 2) has zero footnotes - proves the
    from-scratch `word/footnotes.xml`/wiring creation path end-to-end."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "add_footnote",
            {"path": str(structured_docx), "paragraph_index": 0, "content": "New footnote."},
        )

    assert result.is_error is False
    assert result.structured_content["footnote_id"] == "1"
    footnotes = get_document_footnotes(structured_docx)
    assert [(f.id, f.paragraph_index, f.content) for f in footnotes] == [("1", 0, "New footnote.")]


@pytest.mark.anyio
async def test_call_add_footnote_on_footnotes_fixture_allocates_next_id(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "add_footnote",
            {"path": str(footnotes_docx), "paragraph_index": 5, "content": "Fourth footnote."},
        )

    assert result.is_error is False
    assert result.structured_content["footnote_id"] == "4"


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_add_footnote_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool(
            "add_footnote",
            {"path": str(outside_file), "paragraph_index": 0, "content": "x"},
        )

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_add_footnote_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, structured_docx: Path
) -> None:
    async with client:
        result = await client.call_tool(
            "add_footnote",
            {"path": str(structured_docx), "paragraph_index": 0, "content": "x"},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_add_footnote_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool(
            "add_footnote", {"path": str(missing), "paragraph_index": 0, "content": "x"}
        )

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_add_footnote_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool(
            "add_footnote", {"path": str(corrupt), "paragraph_index": 0, "content": "x"}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_add_footnote_out_of_range_paragraph_index_is_a_tool_error(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "add_footnote",
            {"path": str(structured_docx), "paragraph_index": 999, "content": "x"},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_add_footnote_missing_required_arguments_is_a_tool_error(
    client: Client,
) -> None:
    async with client:
        result = await client.call_tool("add_footnote", {})

    assert result.is_error is True
