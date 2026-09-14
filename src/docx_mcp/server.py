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
from docx_mcp.text_edit import TextEditError, find_text_matches, replace_text_in_document
from docx_mcp.text_edit import TextLocation as _TextLocation

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

FIND_TEXT_TOOL_VERSION = "1.0.0"
"""Semantic version of `find_text` (see `docs/documentation-standards.md` §4.1)."""

REPLACE_TEXT_TOOL_VERSION = "1.0.0"
"""Semantic version of `replace_text` (see `docs/documentation-standards.md` §4.1)."""


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


class TextLocation(BaseModel):
    """A half-open character span in one top-level body paragraph's plain
    text - `find_text`'s match location and `replace_text`'s input, sharing
    one shape (see `specs/find_text.md` §3, `specs/replace_text.md` §2)."""

    paragraph_index: int = Field(description="0-based position among top-level body paragraphs.")
    start_offset: int = Field(description="Start offset (inclusive) into the paragraph's text.")
    end_offset: int = Field(description="End offset (exclusive) into the paragraph's text.")


class TextMatch(BaseModel):
    """One `find_text` match (see `specs/find_text.md` §3)."""

    location: TextLocation = Field(description="Where the match was found.")
    matched_text: str = Field(description="The actual substring found, in its original casing.")
    paragraph_text: str = Field(description="The full plain text of the matched paragraph.")


class FindTextResult(BaseModel):
    """Output of `find_text` (see `specs/find_text.md` §3, Output Schema)."""

    matches: list[TextMatch] = Field(
        description="Every non-overlapping match, in document order. Empty if none found."
    )


def find_text(path: str, search_text: str, case_sensitive: bool = True) -> FindTextResult:
    """Search a .docx file's body for literal text, returning locations replace_text can act on.

    Finds every non-overlapping, literal (non-regex) occurrence of
    search_text across top-level body paragraphs, even when Word has split
    the matched text across multiple internal runs. Each match's location
    (paragraph_index, start_offset, end_offset) is a half-open span into
    that paragraph's plain text (no heading '#'/footnote '[^N]' markers) and
    can be passed directly as replace_text's location argument. See
    specs/find_text.md for the full specification, including why regex,
    table-cell text, and cross-paragraph matches are out of scope.

    Args:
        path: Filesystem path to a `.docx` file, absolute or relative. Must
            resolve inside one of the roots configured via
            `DOCX_MCP_ALLOWED_ROOTS`; an empty/unset allow-list denies every
            path.
        search_text: The literal text to search for. Must be non-empty.
        case_sensitive: `True` (default) for exact-case matching only;
            `False` for Unicode-aware case-insensitive matching (matched_text
            still reflects the text's actual, original casing).

    Returns:
        A `FindTextResult` carrying every match, in document order.

    Raises:
        ToolError: If `path` escapes the allowed roots, the file does not
            exist, the file is not a valid `.docx` document, or `search_text`
            is empty.
    """
    config = load_config()
    try:
        safe_path = resolve_safe_path(path, config.allowed_roots)
    except PathAccessError as exc:
        raise ToolError(str(exc)) from exc

    try:
        matches = find_text_matches(safe_path, search_text, case_sensitive=case_sensitive)
    except (InvalidDocumentError, TextEditError) as exc:
        raise ToolError(str(exc)) from exc

    return FindTextResult(
        matches=[
            TextMatch(
                location=TextLocation(
                    paragraph_index=m.location.paragraph_index,
                    start_offset=m.location.start_offset,
                    end_offset=m.location.end_offset,
                ),
                matched_text=m.matched_text,
                paragraph_text=m.paragraph_text,
            )
            for m in matches
        ]
    )


class ReplaceTextResult(BaseModel):
    """Output of `replace_text` (see `specs/replace_text.md` §3, Output Schema)."""

    replacements_made: int = Field(
        description=(
            "Number of occurrences replaced: always 1 in location mode, always >= 1 "
            "in global mode (zero matches in global mode is an error, not a 0 result)."
        )
    )


def replace_text(
    path: str,
    search_text: str,
    replacement_text: str,
    location: TextLocation | None = None,
    case_sensitive: bool = True,
) -> ReplaceTextResult:
    """Replace literal text in a .docx file, at one location or everywhere, formatting-safely.

    Without location, replaces every occurrence of search_text in the
    document. With location (as returned by find_text), replaces exactly
    that one occurrence - but first re-verifies the document still contains
    search_text there, failing with a clear "stale location" error if it
    does not, so a caller can never silently edit the wrong text because the
    document changed since find_text was called. A boundary run's untouched
    prefix/suffix text keeps its original formatting; the replacement text
    inherits the leftmost matched run's formatting. The write is atomic (see
    docs/security-model.md §3): the original file is only ever touched by a
    single, final replace step. See specs/replace_text.md for the full
    specification, including why regex, table-cell text, cross-paragraph
    matches, and a match spanning a hyperlink/tracked-change boundary are
    out of scope or rejected.

    Args:
        path: Filesystem path to a `.docx` file, absolute or relative. Must
            resolve inside one of the roots configured via
            `DOCX_MCP_ALLOWED_ROOTS`; an empty/unset allow-list denies every
            path.
        search_text: The literal text to replace. Must be non-empty; in
            location mode, also the value re-verified against `location`.
        replacement_text: The literal replacement text; may be empty (a
            pure deletion).
        location: `None` (default) replaces every occurrence. Given (as
            returned by `find_text`), replaces exactly that one occurrence.
        case_sensitive: `True` (default) for exact-case matching/verification;
            `False` for Unicode-aware case-insensitive matching.

    Returns:
        A `ReplaceTextResult` carrying the number of occurrences replaced.

    Raises:
        ToolError: If `path` escapes the allowed roots, the file does not
            exist or is not a valid `.docx` document, `search_text` is
            empty, `location` is invalid or stale, no occurrences are found
            in global mode, or a match spans a run-container boundary.
    """
    config = load_config()
    try:
        safe_path = resolve_safe_path(path, config.allowed_roots)
    except PathAccessError as exc:
        raise ToolError(str(exc)) from exc

    internal_location = (
        _TextLocation(location.paragraph_index, location.start_offset, location.end_offset)
        if location is not None
        else None
    )

    try:
        replacements_made = replace_text_in_document(
            safe_path,
            search_text,
            replacement_text,
            location=internal_location,
            case_sensitive=case_sensitive,
        )
    except (InvalidDocumentError, TextEditError) as exc:
        raise ToolError(str(exc)) from exc

    return ReplaceTextResult(replacements_made=replacements_made)


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
        `get_footnotes`, `find_text`, and `replace_text` registered, and
        logging configured at `config.log_level` (see
        `docs/documentation-standards.md` §6, Logging and observability).
    """
    resolved_config = config if config is not None else load_config()
    server = MCPServer(name="docx-mcp", version=__version__, log_level=resolved_config.log_level)
    server.add_tool(read_document, meta={"docx_mcp.tool_version": READ_DOCUMENT_TOOL_VERSION})
    server.add_tool(get_structure, meta={"docx_mcp.tool_version": GET_STRUCTURE_TOOL_VERSION})
    server.add_tool(get_metadata, meta={"docx_mcp.tool_version": GET_METADATA_TOOL_VERSION})
    server.add_tool(get_footnotes, meta={"docx_mcp.tool_version": GET_FOOTNOTES_TOOL_VERSION})
    server.add_tool(find_text, meta={"docx_mcp.tool_version": FIND_TEXT_TOOL_VERSION})
    server.add_tool(replace_text, meta={"docx_mcp.tool_version": REPLACE_TEXT_TOOL_VERSION})
    return server


mcp = create_server()
"""The process-wide server instance, built from the real environment at import
time (log level and allowed roots as of that moment). `__main__.py` runs this
over stdio. Every tool re-reads `DOCX_MCP_ALLOWED_ROOTS` on every call, so
changing it after import still takes effect per-call; only the log level is
fixed at construction, matching a normal server's startup-time logging
configuration."""
