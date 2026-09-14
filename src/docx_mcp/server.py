"""The docx-mcp stdio MCP server: registers tools as thin wrappers, no business logic.

See [docs/repository-structure.md §2](../../docs/repository-structure.md#2-conventions):
`server.py` only wires a tool's arguments to its implementation module and
translates that module's exceptions into `ToolError`; parsing, security, and
configuration logic live in `docx_mcp.document`, `docx_mcp.security`, and
`docx_mcp.config`.
"""

from __future__ import annotations

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import BaseModel, Field

from docx_mcp import __version__
from docx_mcp.config import ServerConfig, load_config
from docx_mcp.document import InvalidDocumentError, extract_text
from docx_mcp.security import PathAccessError, resolve_safe_path

READ_DOCUMENT_TOOL_VERSION = "1.0.0"
"""Semantic version of `read_document`, per
[docs/documentation-standards.md §4.1](../../docs/documentation-standards.md#41-tool-versioning).
Exposed via the tool's `_meta` field (the MCP protocol has no native tool
version field) rather than folded into the description text."""


class ReadDocumentResult(BaseModel):
    """Output of `read_document` (see `specs/read_document.md` §3, Output Schema)."""

    text: str = Field(
        description=(
            "Full document body text, one paragraph per line, with Markdown-style "
            "heading markers (e.g. '# Heading') and inline footnote-anchor markers "
            "(e.g. '[^1]')."
        )
    )


def read_document(path: str) -> ReadDocumentResult:
    """Read a .docx file and return its full body text, with heading and footnote markers inline.

    Headings (Word's built-in Heading 1-9 styles) are rendered as Markdown
    headings ("# ", "## ", ...); a footnote reference is rendered inline, at
    its exact position, as "[^N]" (its content is not included - use
    get_footnotes once that tool exists). Table content, list numbering, and
    endnotes are not extracted. See specs/read_document.md for the full
    specification.

    Args:
        path: Filesystem path to a `.docx` file, absolute or relative. Must
            resolve inside one of the roots configured via
            `DOCX_MCP_ALLOWED_ROOTS`; an empty/unset allow-list denies every
            path.

    Returns:
        A `ReadDocumentResult` carrying the extracted text.

    Raises:
        ToolError: If `path` escapes the allowed roots, the file does not
            exist, or the file is not a valid `.docx` document.
    """
    config = load_config()
    try:
        safe_path = resolve_safe_path(path, config.allowed_roots)
    except PathAccessError as exc:
        raise ToolError(str(exc)) from exc

    try:
        text = extract_text(safe_path)
    except InvalidDocumentError as exc:
        raise ToolError(str(exc)) from exc

    return ReadDocumentResult(text=text)


def create_server(config: ServerConfig | None = None) -> MCPServer:
    """Build the docx-mcp `MCPServer`, with its process-wide log level applied.

    Args:
        config: Configuration to build the server from. Defaults to
            `load_config()` (the real process environment); tests pass an
            explicit `ServerConfig` instead of mutating `os.environ`, so
            `DOCX_MCP_LOG_LEVEL`'s effect on the server can be verified
            without process-global state.

    Returns:
        An `MCPServer` with `read_document` registered and logging configured
        at `config.log_level` (see `docs/documentation-standards.md` §6,
        Logging and observability).
    """
    resolved_config = config if config is not None else load_config()
    server = MCPServer(name="docx-mcp", version=__version__, log_level=resolved_config.log_level)
    server.add_tool(read_document, meta={"docx_mcp.tool_version": READ_DOCUMENT_TOOL_VERSION})
    return server


mcp = create_server()
"""The process-wide server instance, built from the real environment at import
time (log level and allowed roots as of that moment). `__main__.py` runs this
over stdio. `read_document` itself re-reads `DOCX_MCP_ALLOWED_ROOTS` on every
call, so changing it after import still takes effect per-call; only the log
level is fixed at construction, matching a normal server's startup-time
logging configuration."""
