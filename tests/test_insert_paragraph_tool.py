"""Contract and functional tests for the `insert_paragraph` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/insert_paragraph.md §8](../src/docx_mcp/specs/insert_paragraph.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.document import DOCUMENT_PART, NSMAP, get_body, paragraph_plain_text
from docx_mcp.ooxml import parse_xml, read_part, validate_and_open
from docx_mcp.server import INSERT_PARAGRAPH_TOOL_VERSION, mcp


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
async def test_insert_paragraph_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "insert_paragraph" in names


@pytest.mark.anyio
async def test_insert_paragraph_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "insert_paragraph")
    assert tool.description
    assert "insert" in tool.description.lower()


@pytest.mark.anyio
async def test_insert_paragraph_input_schema_declares_expected_parameters(
    client: Client,
) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "insert_paragraph")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert schema["properties"]["text"]["type"] == "string"
    assert "after_paragraph_index" in schema["properties"]
    assert "heading_level" in schema["properties"]
    assert set(schema["required"]) == {"path", "text"}


@pytest.mark.anyio
async def test_insert_paragraph_output_schema_declares_paragraph_index(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "insert_paragraph")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"paragraph_index"}


@pytest.mark.anyio
async def test_insert_paragraph_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "insert_paragraph")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == INSERT_PARAGRAPH_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_insert_paragraph_after_a_body_paragraph(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "insert_paragraph",
            {
                "path": str(paragraph_edits_docx),
                "text": "New body text.",
                "after_paragraph_index": 1,
            },
        )

    assert result.is_error is False
    assert result.structured_content["paragraph_index"] == 2
    assert _paragraph_texts(paragraph_edits_docx)[2] == "New body text."


@pytest.mark.anyio
async def test_call_insert_paragraph_as_a_heading(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "insert_paragraph",
            {
                "path": str(paragraph_edits_docx),
                "text": "New Section",
                "after_paragraph_index": 4,
                "heading_level": 2,
            },
        )

    assert result.is_error is False
    new_index = result.structured_content["paragraph_index"]
    assert new_index == 5
    assert _paragraph_texts(paragraph_edits_docx)[new_index] == "New Section"


@pytest.mark.anyio
async def test_call_insert_paragraph_at_the_start_without_after_paragraph_index(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "insert_paragraph",
            {"path": str(paragraph_edits_docx), "text": "New first paragraph."},
        )

    assert result.is_error is False
    assert result.structured_content["paragraph_index"] == 0
    assert _paragraph_texts(paragraph_edits_docx)[0] == "New first paragraph."


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_insert_paragraph_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool(
            "insert_paragraph", {"path": str(outside_file), "text": "x"}
        )

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_insert_paragraph_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, paragraph_edits_docx: Path
) -> None:
    async with client:
        result = await client.call_tool(
            "insert_paragraph", {"path": str(paragraph_edits_docx), "text": "x"}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_insert_paragraph_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool("insert_paragraph", {"path": str(missing), "text": "x"})

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_insert_paragraph_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool("insert_paragraph", {"path": str(corrupt), "text": "x"})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_insert_paragraph_out_of_range_after_paragraph_index_is_a_tool_error(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "insert_paragraph",
            {"path": str(paragraph_edits_docx), "text": "x", "after_paragraph_index": 999},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_insert_paragraph_invalid_heading_level_is_a_tool_error(
    client: Client, paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "insert_paragraph",
            {"path": str(paragraph_edits_docx), "text": "x", "heading_level": 10},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_insert_paragraph_missing_required_arguments_is_a_tool_error(
    client: Client,
) -> None:
    async with client:
        result = await client.call_tool("insert_paragraph", {})

    assert result.is_error is True
