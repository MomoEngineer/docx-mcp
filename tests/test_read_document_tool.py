"""Contract and integration tests for the `read_document` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair to exercise the real MCP
protocol (list_tools, call_tool) without a subprocess or stdio pipe, per the
"in-memory MCP client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/read_document.md §8](../src/docx_mcp/specs/read_document.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.server import READ_DOCUMENT_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_read_document_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "read_document" in names


@pytest.mark.anyio
async def test_read_document_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "read_document")
    assert tool.description
    assert "docx" in tool.description.lower() or "document" in tool.description.lower()


@pytest.mark.anyio
async def test_read_document_input_schema_declares_required_path_string(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "read_document")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert "path" in schema["required"]


@pytest.mark.anyio
async def test_read_document_output_schema_declares_text_string(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "read_document")
    assert tool.output_schema is not None
    assert tool.output_schema["properties"]["text"]["type"] == "string"


@pytest.mark.anyio
async def test_read_document_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "read_document")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == READ_DOCUMENT_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_read_document_returns_expected_text(
    client: Client, minimal_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool("read_document", {"path": str(minimal_docx)})

    assert result.is_error is False
    assert result.structured_content["text"] == (
        "# Introduction\n"
        "This is the first paragraph of the document.\n"
        "This paragraph has a footnote reference.[^1]\n"
        "## Background\n"
        "\n"
        "Final paragraph."
    )


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_read_document_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool("read_document", {"path": str(outside_file)})

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_read_document_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, minimal_docx: Path
) -> None:
    async with client:
        result = await client.call_tool("read_document", {"path": str(minimal_docx)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_read_document_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool("read_document", {"path": str(missing)})

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_read_document_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool("read_document", {"path": str(corrupt)})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_read_document_missing_path_argument_is_a_tool_error(client: Client) -> None:
    async with client:
        result = await client.call_tool("read_document", {})

    assert result.is_error is True
