"""Contract and functional tests for the `edit_footnote` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/edit_footnote.md §8](../src/docx_mcp/specs/edit_footnote.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.server import EDIT_FOOTNOTE_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_edit_footnote_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "edit_footnote" in names


@pytest.mark.anyio
async def test_edit_footnote_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "edit_footnote")
    assert tool.description
    assert "footnote" in tool.description.lower()


@pytest.mark.anyio
async def test_edit_footnote_input_schema_declares_expected_parameters(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "edit_footnote")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert schema["properties"]["footnote_id"]["type"] == "string"
    assert schema["properties"]["content"]["type"] == "string"
    assert "expected_content" in schema["properties"]
    assert set(schema["required"]) == {"path", "footnote_id", "content"}


@pytest.mark.anyio
async def test_edit_footnote_output_schema_declares_previous_content(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "edit_footnote")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"previous_content"}


@pytest.mark.anyio
async def test_edit_footnote_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "edit_footnote")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == EDIT_FOOTNOTE_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_edit_footnote_updates_content(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {"path": str(footnotes_docx), "footnote_id": "1", "content": "Updated text."},
        )

    assert result.is_error is False
    assert result.structured_content["previous_content"] == " First footnote."
    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == "Updated text."


@pytest.mark.anyio
async def test_call_edit_footnote_with_matching_expected_content(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {
                "path": str(footnotes_docx),
                "footnote_id": "1",
                "content": "Updated text.",
                "expected_content": " First footnote.",
            },
        )

    assert result.is_error is False
    assert result.structured_content["previous_content"] == " First footnote."


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_edit_footnote_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {"path": str(outside_file), "footnote_id": "1", "content": "x"},
        )

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_edit_footnote_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, footnotes_docx: Path
) -> None:
    async with client:
        result = await client.call_tool(
            "edit_footnote", {"path": str(footnotes_docx), "footnote_id": "1", "content": "x"}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_edit_footnote_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool(
            "edit_footnote", {"path": str(missing), "footnote_id": "1", "content": "x"}
        )

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_edit_footnote_unknown_footnote_id_is_a_tool_error(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {"path": str(footnotes_docx), "footnote_id": "999", "content": "x"},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_edit_footnote_reserved_boilerplate_id_is_a_tool_error(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {"path": str(footnotes_docx), "footnote_id": "-1", "content": "x"},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_edit_footnote_stale_expected_content_is_a_tool_error(
    client: Client, footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "edit_footnote",
            {
                "path": str(footnotes_docx),
                "footnote_id": "1",
                "content": "x",
                "expected_content": "Wrong content.",
            },
        )

    assert result.is_error is True
    assert "stale" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_edit_footnote_missing_required_arguments_is_a_tool_error(
    client: Client,
) -> None:
    async with client:
        result = await client.call_tool("edit_footnote", {})

    assert result.is_error is True
