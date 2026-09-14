# docx-mcp

An MCP (Model Context Protocol) server that lets AI agents read and edit `.docx` Word documents — headings, paragraphs, tables, and footnotes — while preserving the original formatting.

## Why

Most tools that let an LLM touch a `.docx` file fall into one of two traps: they flatten the document to plain text and lose all formatting on write, or they wrap a high-level library (like `python-docx`) that silently mangles anything it doesn't fully understand — fields, tables of contents, footnotes, custom styles. Real-world Word documents (academic papers, business reports, contracts) rely on exactly those features.

`docx-mcp` manipulates the underlying OOXML (the XML/ZIP structure inside every `.docx` file) directly. This keeps existing formatting, styles, tables, images, and footnotes intact even when the agent only touches text — instead of round-tripping the document through a lossy object model.

## Features

- **Read with context, not just text** — full document text plus structure: heading hierarchy, table of contents, tables (as text), and footnotes.
- **Footnotes are first-class** — footnote markers are surfaced inline in the text and can be listed, added, and edited individually, not just dumped as an afterthought.
- **Precise, run-aware text edits** — find and replace text correctly even when Word has split it across multiple internal XML runs (a common source of corruption in naive `.docx` editors).
- **Paragraph-level edit primitives** — insert, delete, and address paragraphs individually so an agent can make targeted edits without rewriting the whole document.
- **Formatting-preserving** — edits reuse the styles and formatting already present in the surrounding document instead of resetting them.
- **Local-first** — runs over stdio, operates on files already on disk. No document content is round-tripped through a remote service.

### Out of scope (for now)

- Editing table structure (adding/removing rows or columns), images, or document theme/styles.
- Comments and tracked changes.
- Remote/HTTP transport (see [Roadmap.md](Roadmap.md)).

## Installation

```bash
uvx docx-mcp
```

or install with `pip`:

```bash
pip install docx-mcp
```

Requires Python 3.11+.

## Usage

### Claude Desktop / Claude Code

Add to your MCP client configuration:

```json
{
  "mcpServers": {
    "docx": {
      "command": "uvx",
      "args": ["docx-mcp"]
    }
  }
}
```

The server operates on local file paths passed as tool arguments — the file must be accessible on the filesystem where the server runs.

## Available tools

| Tool | Type | Description |
|---|---|---|
| `read_document` | Read | Full document text, with heading and footnote markers inline. |
| `get_structure` | Read | Document outline: heading hierarchy, table of contents, table summaries, footnote list, and paragraph index. |
| `get_footnotes` | Read | All footnotes with their ID, anchor location in the text, and content. |
| `get_metadata` | Read | Core document properties (title, author, created/modified dates, word count). |
| `find_text` | Read | Search for literal text; returns match locations (paragraph index and character offsets) for use with `replace_text`. |
| `replace_text` | Write | Replace text at a specific location or globally, without corrupting surrounding formatting. |
| `insert_paragraph` | Write | Insert a new paragraph (body text or heading) at a given position, inheriting the style of its context. |
| `delete_paragraph` | Write | Remove a paragraph by its index/ID. |
| `add_footnote` | Write | Attach a new footnote to a location in the text. |
| `edit_footnote` | Write | Update the content of an existing footnote. |

This is the initial tool surface; it will evolve as real usage surfaces gaps — see [Roadmap.md](Roadmap.md).

## Project status

Early development. Interfaces and tool names may still change.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT
