"""Contract and functional tests for the `replace_text` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/replace_text.md §8](../src/docx_mcp/specs/replace_text.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.document import DOCUMENT_PART, NSMAP, get_body, paragraph_plain_text
from docx_mcp.ooxml import parse_xml, read_part, validate_and_open
from docx_mcp.server import REPLACE_TEXT_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


def _document_text(docx_path: Path) -> str:
    with validate_and_open(docx_path) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)
    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    body = get_body(root)
    return "\n".join(paragraph_plain_text(p) for p in body.findall("w:p", namespaces=NSMAP))


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_replace_text_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "replace_text" in names


@pytest.mark.anyio
async def test_replace_text_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "replace_text")
    assert tool.description
    assert "replace" in tool.description.lower()


@pytest.mark.anyio
async def test_replace_text_input_schema_declares_expected_parameters(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "replace_text")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert schema["properties"]["search_text"]["type"] == "string"
    assert schema["properties"]["replacement_text"]["type"] == "string"
    assert "location" in schema["properties"]
    assert set(schema["required"]) == {"path", "search_text", "replacement_text"}


@pytest.mark.anyio
async def test_replace_text_output_schema_declares_replacements_made(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "replace_text")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"replacements_made"}


@pytest.mark.anyio
async def test_replace_text_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "replace_text")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == REPLACE_TEXT_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_replace_text_at_a_location(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
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

    assert result.is_error is False
    assert result.structured_content["replacements_made"] == 1
    assert "Please REPLIED this message." in _document_text(run_split_docx)


@pytest.mark.anyio
async def test_call_replace_text_globally(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "replace_text",
            {"path": str(run_split_docx), "search_text": "test", "replacement_text": "exam"},
        )

    assert result.is_error is False
    assert result.structured_content["replacements_made"] == 4
    assert "test" not in _document_text(run_split_docx)


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_replace_text_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool(
            "replace_text",
            {"path": str(outside_file), "search_text": "a", "replacement_text": "b"},
        )

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_replace_text_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, run_split_docx: Path
) -> None:
    async with client:
        result = await client.call_tool(
            "replace_text",
            {"path": str(run_split_docx), "search_text": "a", "replacement_text": "b"},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_replace_text_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool(
            "replace_text",
            {"path": str(missing), "search_text": "a", "replacement_text": "b"},
        )

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_replace_text_empty_search_text_is_a_tool_error(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "replace_text",
            {"path": str(run_split_docx), "search_text": "", "replacement_text": "b"},
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_replace_text_zero_global_matches_is_a_tool_error(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "replace_text",
            {
                "path": str(run_split_docx),
                "search_text": "not present anywhere",
                "replacement_text": "b",
            },
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_replace_text_stale_location_is_a_tool_error(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    location = {"paragraph_index": 0, "start_offset": 7, "end_offset": 14}

    async with client:
        first = await client.call_tool(
            "replace_text",
            {
                "path": str(run_split_docx),
                "search_text": "receive",
                "replacement_text": "REPLIED",
                "location": location,
            },
        )
        assert first.is_error is False

        second = await client.call_tool(
            "replace_text",
            {
                "path": str(run_split_docx),
                "search_text": "receive",
                "replacement_text": "AGAIN",
                "location": location,
            },
        )

    assert second.is_error is True
    assert "stale" in second.content[0].text.lower()


@pytest.mark.anyio
async def test_call_replace_text_missing_required_arguments_is_a_tool_error(
    client: Client,
) -> None:
    async with client:
        result = await client.call_tool("replace_text", {})

    assert result.is_error is True
