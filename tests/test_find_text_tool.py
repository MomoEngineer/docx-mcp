"""Contract and functional tests for the `find_text` MCP tool.

Uses the SDK's in-memory `Client`/`MCPServer` pair, per the "in-memory MCP
client" convention in
[docs/repository-structure.md §2](../docs/repository-structure.md#2-conventions).
See [specs/find_text.md §8](../src/docx_mcp/specs/find_text.md#8-test-coverage).
"""

from __future__ import annotations

from pathlib import Path

import pytest
from mcp import Client

from docx_mcp.server import FIND_TEXT_TOOL_VERSION, mcp


@pytest.fixture
def client() -> Client:
    return Client(mcp, mode="legacy")


# --- Contract -----------------------------------------------------------------


@pytest.mark.anyio
async def test_find_text_is_discoverable(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    names = [tool.name for tool in result.tools]
    assert "find_text" in names


@pytest.mark.anyio
async def test_find_text_description_is_present_and_non_trivial(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "find_text")
    assert tool.description
    assert "text" in tool.description.lower()


@pytest.mark.anyio
async def test_find_text_input_schema_declares_expected_parameters(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "find_text")
    schema = tool.input_schema
    assert schema["type"] == "object"
    assert schema["properties"]["path"]["type"] == "string"
    assert schema["properties"]["search_text"]["type"] == "string"
    assert schema["properties"]["case_sensitive"]["type"] == "boolean"
    assert set(schema["required"]) == {"path", "search_text"}


@pytest.mark.anyio
async def test_find_text_output_schema_declares_matches_array(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "find_text")
    assert tool.output_schema is not None
    assert set(tool.output_schema["required"]) == {"matches"}


@pytest.mark.anyio
async def test_find_text_version_is_exposed_via_meta(client: Client) -> None:
    async with client:
        result = await client.list_tools()

    tool = next(t for t in result.tools if t.name == "find_text")
    assert tool.meta is not None
    assert tool.meta["docx_mcp.tool_version"] == FIND_TEXT_TOOL_VERSION == "1.0.0"


# --- Functional -----------------------------------------------------------------


@pytest.mark.anyio
async def test_call_find_text_returns_expected_location(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "find_text", {"path": str(run_split_docx), "search_text": "receive"}
        )

    assert result.is_error is False
    matches = result.structured_content["matches"]
    assert len(matches) == 1
    assert matches[0]["location"] == {"paragraph_index": 0, "start_offset": 7, "end_offset": 14}
    assert matches[0]["matched_text"] == "receive"
    assert matches[0]["paragraph_text"] == "Please receive this message."


@pytest.mark.anyio
async def test_call_find_text_case_insensitive(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "find_text",
            {"path": str(run_split_docx), "search_text": "RECEIVE", "case_sensitive": False},
        )

    assert result.is_error is False
    assert result.structured_content["matches"][0]["matched_text"] == "receive"


@pytest.mark.anyio
async def test_call_find_text_no_match_returns_empty_list(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "find_text", {"path": str(run_split_docx), "search_text": "not present anywhere"}
        )

    assert result.is_error is False
    assert result.structured_content["matches"] == []


# --- Error / edge (surfaced as ToolError -> is_error=True) --------------------


@pytest.mark.anyio
async def test_call_find_text_outside_allowed_roots_is_a_tool_error(
    client: Client, outside_root: Path, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    outside_file = outside_root / "secret.docx"
    outside_file.write_bytes(b"")

    async with client:
        result = await client.call_tool(
            "find_text", {"path": str(outside_file), "search_text": "x"}
        )

    message = result.content[0].text.lower()
    assert result.is_error is True
    assert "outside" in message or "allowed" in message


@pytest.mark.anyio
async def test_call_find_text_with_no_allowed_roots_configured_is_a_tool_error(
    client: Client, run_split_docx: Path
) -> None:
    async with client:
        result = await client.call_tool(
            "find_text", {"path": str(run_split_docx), "search_text": "x"}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_find_text_on_missing_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    missing = allowed_root / "does-not-exist.docx"

    async with client:
        result = await client.call_tool("find_text", {"path": str(missing), "search_text": "x"})

    assert result.is_error is True
    assert "not found" in result.content[0].text.lower()


@pytest.mark.anyio
async def test_call_find_text_on_corrupt_file_is_a_tool_error(
    client: Client, allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))
    corrupt = allowed_root / "corrupt.docx"
    corrupt.write_bytes(b"not a zip file")

    async with client:
        result = await client.call_tool("find_text", {"path": str(corrupt), "search_text": "x"})

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_find_text_empty_search_text_is_a_tool_error(
    client: Client, run_split_docx: Path, monkeypatch: pytest.MonkeyPatch, allowed_root: Path
) -> None:
    monkeypatch.setenv("DOCX_MCP_ALLOWED_ROOTS", str(allowed_root))

    async with client:
        result = await client.call_tool(
            "find_text", {"path": str(run_split_docx), "search_text": ""}
        )

    assert result.is_error is True


@pytest.mark.anyio
async def test_call_find_text_missing_required_arguments_is_a_tool_error(client: Client) -> None:
    async with client:
        result = await client.call_tool("find_text", {})

    assert result.is_error is True
