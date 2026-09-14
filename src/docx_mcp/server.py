"""The docx-mcp stdio MCP server: registers tools as thin wrappers, no business logic.

See [docs/repository-structure.md §2](../../docs/repository-structure.md#2-conventions):
`server.py` only wires a tool's arguments to its implementation module and
translates that module's exceptions into `ToolError`; parsing, security, and
configuration logic live in `docx_mcp.document`, `docx_mcp.structure`,
`docx_mcp.metadata`, `docx_mcp.security`, and `docx_mcp.config`.
"""

from __future__ import annotations

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import BaseModel, Field

from docx_mcp import __version__
from docx_mcp.config import ServerConfig, load_config
from docx_mcp.document import InvalidDocumentError, extract_text
from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.metadata import get_document_metadata
from docx_mcp.security import PathAccessError, resolve_safe_path
from docx_mcp.structure import get_document_structure

READ_DOCUMENT_TOOL_VERSION = "1.0.0"
"""Semantic version of `read_document`, per
[docs/documentation-standards.md §4.1](../../docs/documentation-standards.md#41-tool-versioning).
Exposed via the tool's `_meta` field (the MCP protocol has no native tool
version field) rather than folded into the description text."""

GET_STRUCTURE_TOOL_VERSION = "1.0.0"
"""Semantic version of `get_structure` (see `docs/documentation-standards.md` §4.1)."""

GET_METADATA_TOOL_VERSION = "1.0.0"
"""Semantic version of `get_metadata` (see `docs/documentation-standards.md` §4.1)."""

GET_FOOTNOTES_TOOL_VERSION = "1.0.0"
"""Semantic version of `get_footnotes` (see `docs/documentation-standards.md` §4.1)."""


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


class ParagraphEntry(BaseModel):
    """One top-level body paragraph (see `specs/get_structure.md` §3)."""

    paragraph_index: int = Field(description="0-based position among top-level body paragraphs.")
    text: str = Field(description="Plain paragraph text - no heading or footnote markers.")
    style_id: str | None = Field(description="Raw w:pPr/w:pStyle id, or null if unstyled.")
    heading_level: int | None = Field(
        description="Resolved heading level (1-9), or null if not a heading."
    )


class TocEntry(BaseModel):
    """One heading, filtered from `paragraphs` (see `specs/get_structure.md` §3)."""

    paragraph_index: int = Field(description="Index into the paragraphs list.")
    level: int = Field(description="Resolved heading level (1-9).")
    text: str = Field(description="The heading's plain text.")


class TableEntry(BaseModel):
    """One top-level table, rendered as rows of cell text (see `specs/get_structure.md` §3)."""

    table_index: int = Field(description="0-based position among top-level tables.")
    rows: list[list[str]] = Field(
        description=(
            "One entry per row, each a list of cell texts. Rows are not padded to a "
            "uniform column count - a horizontally merged cell yields a shorter row."
        )
    )


class FootnoteAnchorEntry(BaseModel):
    """One footnote reference's anchor location (see `specs/get_structure.md` §3)."""

    id: str = Field(description="The footnote's w:id.")
    paragraph_index: int = Field(description="Index of the paragraph anchoring this footnote.")


class GetStructureResult(BaseModel):
    """Output of `get_structure` (see `specs/get_structure.md` §3, Output Schema)."""

    paragraphs: list[ParagraphEntry] = Field(
        description="Every top-level body paragraph, in document order."
    )
    toc: list[TocEntry] = Field(description="Every heading, derived from paragraphs.")
    tables: list[TableEntry] = Field(description="Every top-level table, in document order.")
    footnotes: list[FootnoteAnchorEntry] = Field(
        description="Footnote-anchor index (id + paragraph_index only, no resolved content)."
    )


def get_structure(path: str) -> GetStructureResult:
    """Return a .docx file's structural outline: paragraphs, TOC, tables, footnote anchors.

    Every top-level body paragraph is indexed with its resolved heading level
    (via styles.xml's outlineLvl/basedOn chain, falling back to the HeadingN
    style-id heuristic); the table of contents is derived from those headings;
    every top-level table is rendered as rows of cell text; footnotes are
    listed as a lightweight id+anchor index only (use get_footnotes, once that
    tool exists, for resolved footnote content). See specs/get_structure.md
    for the full specification, including the paragraph_index scope
    (top-level body paragraphs only - not text inside table cells) and the
    table-row-length caveat for merged cells.

    Args:
        path: Filesystem path to a `.docx` file, absolute or relative. Must
            resolve inside one of the roots configured via
            `DOCX_MCP_ALLOWED_ROOTS`; an empty/unset allow-list denies every
            path.

    Returns:
        A `GetStructureResult` carrying the paragraph index, table of
        contents, tables, and footnote-anchor index.

    Raises:
        ToolError: If `path` escapes the allowed roots, the file does not
            exist, the file is not a valid `.docx` document, or `word/styles.xml`
            is present but malformed.
    """
    config = load_config()
    try:
        safe_path = resolve_safe_path(path, config.allowed_roots)
    except PathAccessError as exc:
        raise ToolError(str(exc)) from exc

    try:
        structure = get_document_structure(safe_path)
    except InvalidDocumentError as exc:
        raise ToolError(str(exc)) from exc

    return GetStructureResult(
        paragraphs=[
            ParagraphEntry(
                paragraph_index=p.paragraph_index,
                text=p.text,
                style_id=p.style_id,
                heading_level=p.heading_level,
            )
            for p in structure.paragraphs
        ],
        toc=[
            TocEntry(paragraph_index=t.paragraph_index, level=t.level, text=t.text)
            for t in structure.toc
        ],
        tables=[
            TableEntry(table_index=t.table_index, rows=[list(row) for row in t.rows])
            for t in structure.tables
        ],
        footnotes=[
            FootnoteAnchorEntry(id=f.id, paragraph_index=f.paragraph_index)
            for f in structure.footnotes
        ],
    )


class GetMetadataResult(BaseModel):
    """Output of `get_metadata` (see `specs/get_metadata.md` §3, Output Schema)."""

    title: str | None = Field(description="docProps/core.xml's dc:title, or null.")
    author: str | None = Field(description="docProps/core.xml's dc:creator, or null.")
    created: str | None = Field(
        description="docProps/core.xml's dcterms:created, verbatim (not reformatted), or null."
    )
    modified: str | None = Field(
        description="docProps/core.xml's dcterms:modified, verbatim (not reformatted), or null."
    )
    word_count: int = Field(
        description="Whitespace-token count across all paragraph and table-cell text."
    )


def get_metadata(path: str) -> GetMetadataResult:
    """Return a .docx file's core document properties (title, author, dates) and word count.

    Reads docProps/core.xml for title/author/created/modified - null for any
    field that is absent, empty, or if the whole part is missing or malformed
    (a missing/malformed docProps/core.xml never fails the call, unlike a
    missing/malformed word/document.xml). word_count is a whitespace-token
    count over the document's plain text (no heading/footnote markers),
    always computed regardless of docProps/core.xml. See specs/get_metadata.md
    for the full specification.

    Args:
        path: Filesystem path to a `.docx` file, absolute or relative. Must
            resolve inside one of the roots configured via
            `DOCX_MCP_ALLOWED_ROOTS`; an empty/unset allow-list denies every
            path.

    Returns:
        A `GetMetadataResult` carrying title, author, created, modified, and word_count.

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
        metadata = get_document_metadata(safe_path)
    except InvalidDocumentError as exc:
        raise ToolError(str(exc)) from exc

    return GetMetadataResult(
        title=metadata.title,
        author=metadata.author,
        created=metadata.created,
        modified=metadata.modified,
        word_count=metadata.word_count,
    )


class FootnoteEntry(BaseModel):
    """One footnote reference, resolved to its content (see `specs/get_footnotes.md` §3)."""

    id: str = Field(description="The footnote's w:id.")
    paragraph_index: int = Field(description="Index of the paragraph anchoring this footnote.")
    content: str | None = Field(
        description=(
            "The footnote's text, or null if word/footnotes.xml is absent or does not "
            "declare a matching id - never an error by itself."
        )
    )


class GetFootnotesResult(BaseModel):
    """Output of `get_footnotes` (see `specs/get_footnotes.md` §3, Output Schema)."""

    footnotes: list[FootnoteEntry] = Field(
        description=(
            "One entry per w:footnoteReference found in the body, in document order. "
            "Two different footnote ids anchored in the same paragraph produce two "
            "separate entries with the same paragraph_index."
        )
    )


def get_footnotes(path: str) -> GetFootnotesResult:
    """Return every footnote a .docx file's body references: its ID, anchor paragraph, and content.

    Population is anchor-driven: one entry per w:footnoteReference actually
    found in the body (the same population get_structure's footnote index
    uses), not one per footnote declared in word/footnotes.xml - an orphaned,
    never-referenced footnote is intentionally excluded. Two different
    footnote ids anchored in the same paragraph both appear, as separate
    entries with the same paragraph_index, never collapsed into one. content
    is resolved from word/footnotes.xml with the same marker-free text
    assembly get_structure uses for plain paragraph text; it is null (not an
    error) when word/footnotes.xml is absent or doesn't declare a matching
    id. See specs/get_footnotes.md for the full specification.

    Args:
        path: Filesystem path to a `.docx` file, absolute or relative. Must
            resolve inside one of the roots configured via
            `DOCX_MCP_ALLOWED_ROOTS`; an empty/unset allow-list denies every
            path.

    Returns:
        A `GetFootnotesResult` carrying the footnote list.

    Raises:
        ToolError: If `path` escapes the allowed roots, the file does not
            exist, the file is not a valid `.docx` document, or
            `word/footnotes.xml` is present but malformed.
    """
    config = load_config()
    try:
        safe_path = resolve_safe_path(path, config.allowed_roots)
    except PathAccessError as exc:
        raise ToolError(str(exc)) from exc

    try:
        footnotes = get_document_footnotes(safe_path)
    except InvalidDocumentError as exc:
        raise ToolError(str(exc)) from exc

    return GetFootnotesResult(
        footnotes=[
            FootnoteEntry(id=f.id, paragraph_index=f.paragraph_index, content=f.content)
            for f in footnotes
        ]
    )


def create_server(config: ServerConfig | None = None) -> MCPServer:
    """Build the docx-mcp `MCPServer`, with its process-wide log level applied.

    Args:
        config: Configuration to build the server from. Defaults to
            `load_config()` (the real process environment); tests pass an
            explicit `ServerConfig` instead of mutating `os.environ`, so
            `DOCX_MCP_LOG_LEVEL`'s effect on the server can be verified
            without process-global state.

    Returns:
        An `MCPServer` with `read_document`, `get_structure`, `get_metadata`,
        and `get_footnotes` registered, and logging configured at
        `config.log_level` (see `docs/documentation-standards.md` §6,
        Logging and observability).
    """
    resolved_config = config if config is not None else load_config()
    server = MCPServer(name="docx-mcp", version=__version__, log_level=resolved_config.log_level)
    server.add_tool(read_document, meta={"docx_mcp.tool_version": READ_DOCUMENT_TOOL_VERSION})
    server.add_tool(get_structure, meta={"docx_mcp.tool_version": GET_STRUCTURE_TOOL_VERSION})
    server.add_tool(get_metadata, meta={"docx_mcp.tool_version": GET_METADATA_TOOL_VERSION})
    server.add_tool(get_footnotes, meta={"docx_mcp.tool_version": GET_FOOTNOTES_TOOL_VERSION})
    return server


mcp = create_server()
"""The process-wide server instance, built from the real environment at import
time (log level and allowed roots as of that moment). `__main__.py` runs this
over stdio. Every tool re-reads `DOCX_MCP_ALLOWED_ROOTS` on every call, so
changing it after import still takes effect per-call; only the log level is
fixed at construction, matching a normal server's startup-time logging
configuration."""
