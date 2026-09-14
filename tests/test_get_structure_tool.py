"""Contract and integration tests for the `get_structure` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/get_structure.md §8](../src/docx_mcp/specs/get_structure.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.server import GET_STRUCTURE_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_get_structure_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "get_structure" in names


@pytest.mark.anyio
async def test_get_structure_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_structure")
    assert tool.description
    assert "structure" in tool.description.lower() or "heading" in tool.description.lower()


@pytest.mark.anyio
async def test_get_structure_input_schema_declares_required_path_string(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_structure")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert "path" in schema["required"]


@pytest.mark.anyio
async def test_get_structure_output_schema_declares_expected_top_level_fields(
    client: Client,
) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_structure")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"paragraphs", "toc", "tables", "footnotes"}


@pytest.mark.anyio
async def test_get_structure_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_structure")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == GET_STRUCTURE_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_get_structure_returns_expected_headings_and_table(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_structure", {"path": str(structured_docx)})

    assert result.is_error is False
    content = result.structured_content
    assert [p["heading_level"] for p in content["paragraphs"]] == [1, None, 2, 2, 3, None]
    assert [t["paragraph_index"] for t in content["toc"]] == [0, 2, 3, 4]
    assert len(content["tables"]) == 1
    assert content["tables"][0]["rows"] == [
        ["Header A", "Header B"],
        ["Value 1", "Value 2"],
        ["Merged Row"],
    ]
    assert content["footnotes"] == []


@pytest.mark.anyio
async def test_call_get_structure_on_minimal_fixture_finds_the_footnote_anchor(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_structure", {"path": str(minimal_docx)})

    assert result.is_error is False
    assert result.structured_content["footnotes"] == [{"id": "1", "paragraph_index": 2}]


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_get_structure_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool("get_structure", {"path": str(outside_file)})

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_get_structure_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, structured_docx: Path
) -> None:
    async with client:
        result = await client.call_tool("get_structure", {"path": str(structured_docx)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_get_structure_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool("get_structure", {"path": str(missing)})

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_get_structure_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool("get_structure", {"path": str(corrupt)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_get_structure_missing_path_argument_is_a_tool_error(client: Client) -> None:
    async with client:
        result = await client.call_tool("get_structure", {})

    assert result.is_error is True
