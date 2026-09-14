# Repository Structure and Definition of Done

This document defines the **binding folder structure** and summarizes the **Definition of Done**. Structure and discipline are modeled on `research-graphrag`, right-sized for a single, focused MCP server (see [CONTRIBUTING.md](../CONTRIBUTING.md)).

---

## 1. Binding folder structure

```
docx-mcp/
├─ README.md                     # Project scope, features, tool list
├─ CONTRIBUTING.md                # Central rulebook
├─ Roadmap.md                     # Phased implementation plan
├─ pyproject.toml                 # (Phase 1) src layout, dependencies
├─ .env.example                   # Example environment variables (no secrets)
├─ .gitignore
├─ docs/
│  ├─ repository-structure.md     # this document
│  ├─ testing.md                  # test strategy: fixture-based, offline, spec-first
│  ├─ documentation-standards.md  # docstrings, typing, tool-spec, security, logging
│  ├─ security-model.md           # path-sandboxing rules, atomic-write contract, threat model
│  └─ adr/
│     ├─ README.md                # ADR process
│     ├─ 0001-*.md                # (from Phase 1) OOXML library and module layout
│     └─ 0002-*.md                # (from Phase 1) dependency/lockfile strategy
├─ templates/
│  ├─ tool-spec.md                # per-tool specification template
│  └─ adr-template.md             # ADR template
├─ src/docx_mcp/                  # (from Phase 1) package skeleton
│  ├─ __init__.py                 #   package version
│  ├─ __main__.py                 #   entry point (python -m docx_mcp)
│  ├─ server.py                   #   MCPServer factory + tool registration, as thin wrappers
│  ├─ config.py                   #   environment-variable configuration (DOCX_MCP_ALLOWED_ROOTS, DOCX_MCP_LOG_LEVEL)
│  ├─ security.py                 #   path sandboxing (resolve_safe_path)
│  ├─ document.py                 #   OOXML parsing/extraction for read_document (lxml, hardened parser)
│  └─ specs/                      #   per-tool specifications (one file per tool, from templates/tool-spec.md)
└─ tests/                         # mirrors src/docx_mcp/
   └─ fixtures/                   # small synthetic .docx fixtures — never real personal documents
```

> **Phase 0 note (historical):** as of Phase 0, only `README.md`, `CONTRIBUTING.md`, `Roadmap.md`, `docs/`, `templates/`, `.gitignore`, and `.env.example` existed as real content; `src/docx_mcp/`, `tests/`, and `tests/fixtures/` existed as empty folders (kept in git via `.gitkeep`). As of Phase 1 ([ADR-0001](adr/0001-ooxml-library-and-module-layout.md), [ADR-0002](adr/0002-dependency-and-lockfile-strategy.md)), `pyproject.toml` and the module layout above exist for real; the `.gitkeep` placeholders are gone.

---

## 2. Conventions

- **Tool logic = one module** under `src/docx_mcp/`; `server.py` registers tools as **thin wrappers**, consistent with `research-graphrag`'s convention that the server module itself carries no business logic.
- **Tool contracts** are checked via an in-memory MCP client, exercising the real protocol (name, description, input schema, version) against the tool spec.
- **One tool = one specification** under `src/docx_mcp/specs/<tool>.md`.
- The submodule layout of `src/docx_mcp/` shown above was decided by [ADR-0001](adr/0001-ooxml-library-and-module-layout.md) once Phase 1's real parsing need existed (left open in Phase 0 on purpose); `document.py` stays a single module until a second read path (Phase 2/3) would otherwise duplicate logic against it.
- **Tool names** are descriptive and domain-specific, matching [README.md](../README.md#available-tools): `read_document`, `get_structure`, `get_footnotes`, `get_metadata`, `find_text`, `replace_text`, `insert_paragraph`, `delete_paragraph`, `add_footnote`, `edit_footnote`.

---

## 3. Definition of Done (per contribution)

Identical to [CONTRIBUTING.md §5](../CONTRIBUTING.md#5-definition-of-done), summarized here as a checklist:

- [ ] Tool specification exists, written **before** the implementation.
- [ ] Tests exist, were written from the spec before the implementation, and are green.
- [ ] Filesystem-facing changes: a path-traversal rejection test.
- [ ] Document-write changes: a test proving a failed/interrupted write leaves the original untouched.
- [ ] Full type hints; `mypy` clean.
- [ ] `ruff check` and `ruff format --check` clean.
- [ ] Docstrings for all public functions/tools.
- [ ] Module documentation current for every touched module.
- [ ] Architecturally relevant decisions: an ADR exists.
