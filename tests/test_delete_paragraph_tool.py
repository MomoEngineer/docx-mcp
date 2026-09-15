"""Contract and functional tests for the `delete_paragraph` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/delete_paragraph.md §8](../src/docx_mcp/specs/delete_paragraph.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.document import DOCUMENT_PART, NSMAP, get_body, paragraph_plain_text
from docx_mcp.ooxml import parse_xml, read_part, validate_and_open
from docx_mcp.server import DELETE_PARAGRAPH_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


def _paragraph_texts(docx_path: Path) -> list[str]:
    with validate_and_open(docx_path) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)
    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    body = get_body(root)
    return [paragraph_plain_text(p) for p in body.findall("w:p", namespaces=NSMAP)]


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_delete_paragraph_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "delete_paragraph" in names


@pytest.mark.anyio
async def test_delete_paragraph_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "delete_paragraph")
    assert tool.description
    assert "delete" in tool.description.lower() or "remove" in tool.description.lower()


@pytest.mark.anyio
async def test_delete_paragraph_input_schema_declares_expected_parameters(
    client: Client,
) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "delete_paragraph")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert schema["properties"]["paragraph_index"]["type"] == "integer"
    assert "expected_text" in schema["properties"]
    assert set(schema["required"]) == {"path", "paragraph_index"}


@pytest.mark.anyio
async def test_delete_paragraph_output_schema_declares_deleted_text(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "delete_paragraph")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"deleted_text"}


@pytest.mark.anyio
async def test_delete_paragraph_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "delete_paragraph")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == DELETE_PARAGRAPH_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_delete_paragraph_by_index(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "delete_paragraph", {"path": str(paragraph_edits_docx), "paragraph_index": 1}
        )

    assert result.is_error is False
    assert result.structured_content["deleted_text"] == "First body paragraph."
    assert "First body paragraph." not in _paragraph_texts(paragraph_edits_docx)


@pytest.mark.anyio
async def test_call_delete_paragraph_with_matching_expected_text(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "delete_paragraph",
            {
                "path": str(paragraph_edits_docx),
                "paragraph_index": 1,
                "expected_text": "First body paragraph.",
            },
        )

    assert result.is_error is False
    assert result.structured_content["deleted_text"] == "First body paragraph."


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_delete_paragraph_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool(
            "delete_paragraph", {"path": str(outside_file), "paragraph_index": 0}
        )

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_delete_paragraph_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, paragraph_edits_docx: Path
) -> None:
    async with client:
        result = await client.call_tool(
            "delete_paragraph", {"path": str(paragraph_edits_docx), "paragraph_index": 0}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_delete_paragraph_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool(
            "delete_paragraph", {"path": str(missing), "paragraph_index": 0}
        )

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_delete_paragraph_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool(
            "delete_paragraph", {"path": str(corrupt), "paragraph_index": 0}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_delete_paragraph_out_of_range_index_is_a_tool_error(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "delete_paragraph", {"path": str(paragraph_edits_docx), "paragraph_index": 999}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_delete_paragraph_stale_expected_text_is_a_tool_error(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "delete_paragraph",
            {
                "path": str(paragraph_edits_docx),
                "paragraph_index": 1,
                "expected_text": "Wrong text.",
            },
        )

    assert result.is_error is True
    assert "stale" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_delete_paragraph_missing_required_arguments_is_a_tool_error(
    client: Client,
) -> None:
    async with client:
        result = await client.call_tool("delete_paragraph", {})

    assert result.is_error is True
