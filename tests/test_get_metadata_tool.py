"""Contract and integration tests for the `get_metadata` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/get_metadata.md §8](../src/docx_mcp/specs/get_metadata.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.server import GET_METADATA_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_get_metadata_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "get_metadata" in names


@pytest.mark.anyio
async def test_get_metadata_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_metadata")
    assert tool.description
    assert "metadata" in tool.description.lower() or "properties" in tool.description.lower()


@pytest.mark.anyio
async def test_get_metadata_input_schema_declares_required_path_string(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_metadata")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert "path" in schema["required"]


@pytest.mark.anyio
async def test_get_metadata_output_schema_declares_expected_fields(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_metadata")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {
        "title",
        "author",
        "created",
        "modified",
        "word_count",
    }


@pytest.mark.anyio
async def test_get_metadata_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "get_metadata")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == GET_METADATA_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_get_metadata_returns_expected_core_properties(
    client: Client, structured_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_metadata", {"path": str(structured_docx)})

    assert result.is_error is False
    content = result.structured_content
    assert content["title"] == "Structured Test Document"
    assert content["author"] == "Test Author"
    assert content["created"] == "2024-01-01T00:00:00Z"
    assert content["modified"] == "2024-06-15T12:30:00Z"
    assert content["word_count"] == 29


@pytest.mark.anyio
async def test_call_get_metadata_on_minimal_fixture_has_no_core_properties(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("get_metadata", {"path": str(minimal_docx)})

    assert result.is_error is False
    content = result.structured_content
    assert content["title"] is None
    assert content["word_count"] == 18


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_get_metadata_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool("get_metadata", {"path": str(outside_file)})

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_get_metadata_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, structured_docx: Path
) -> None:
    async with client:
        result = await client.call_tool("get_metadata", {"path": str(structured_docx)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_get_metadata_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool("get_metadata", {"path": str(missing)})

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_get_metadata_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool("get_metadata", {"path": str(corrupt)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_get_metadata_missing_path_argument_is_a_tool_error(client: Client) -> None:
    async with client:
        result = await client.call_tool("get_metadata", {})

    assert result.is_error is True
