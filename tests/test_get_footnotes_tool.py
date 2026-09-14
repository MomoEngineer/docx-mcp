"""Contract and functional tests for the `get_footnotes` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/get_footnotes.md §8](../src/docx_mcp/specs/get_footnotes.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.server import GET_FOOTNOTES_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_get_footnotes_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "get_footnotes" in names


@pytest.mark.anyio
async def test_get_footnotes_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_footnotes")
    assert tool.description
    assert "footnote" in tool.description.lower()


@pytest.mark.anyio
async def test_get_footnotes_input_schema_declares_required_path_string(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_footnotes")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert "path" in schema["required"]


@pytest.mark.anyio
async def test_get_footnotes_output_schema_declares_expected_top_level_field(
    client: Client,
) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_footnotes")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"footnotes"}


@pytest.mark.anyio
async def test_get_footnotes_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_footnotes")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == GET_FOOTNOTES_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_get_footnotes_on_footnotes_fixture_returns_all_three_entries(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(footnotes_docx)})

    assert result.is_error is False
    footnotes = result.structured_content["footnotes"]
    assert [(f["id"], f["paragraph_index"], f["content"]) for f in footnotes] == [
        ("1", 2, " First footnote."),
        ("2", 2, " Second footnote."),
        ("3", 4, " Third footnote."),
    ]


@pytest.mark.anyio
async def test_call_get_footnotes_on_minimal_fixture_resolves_the_single_footnote(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(minimal_docx)})

    assert result.is_error is False
    assert result.structured_content["footnotes"] == [
        {"id": "1", "paragraph_index": 2, "content": " This is a footnote."}
    ]


@pytest.mark.anyio
async def test_call_get_footnotes_on_structured_fixture_returns_empty_list(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    """`structured.docx` (Phase 2) has no footnotes at all."""
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(structured_docx)})

    assert result.is_error is False
    assert result.structured_content["footnotes"] == []


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_get_footnotes_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(outside_file)})

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_get_footnotes_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, footnotes_docx: Path
) -> None:
    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(footnotes_docx)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_get_footnotes_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(missing)})

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_get_footnotes_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool("get_footnotes", {"path": str(corrupt)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_get_footnotes_missing_path_argument_is_a_tool_error(client: Client) -> None:
    async with client:
        result = await client.call_tool("get_footnotes", {})

    assert result.is_error is True
