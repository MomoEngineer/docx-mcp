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
│     ├─ 0002-*.md                # (from Phase 1) dependency/lockfile strategy
│     ├─ 0003-*.md                # (from Phase 2) module layout: ooxml.py/structure.py/metadata.py
│     ├─ 0004-*.md                # (from Phase 3) footnotes.py and shared anchor resolution
│     ├─ 0005-*.md                # (from Phase 4) text_edit.py and the atomic-write contract
│     ├─ 0006-*.md                # (from Phase 5) paragraph_edit.py and style-inheritance rules
│     └─ 0007-*.md                # (from Phase 6) footnote_edit.py and multi-part atomic writes
├─ templates/
│  ├─ tool-spec.md                # per-tool specification template
│  └─ adr-template.md             # ADR template
├─ src/docx_mcp/                  # (from Phase 1) package skeleton
│  ├─ __init__.py                 #   package version
│  ├─ __main__.py                 #   entry point (python -m docx_mcp)
│  ├─ server.py                   #   MCPServer factory + tool registration, as thin wrappers
│  ├─ config.py                   #   environment-variable configuration (DOCX_MCP_ALLOWED_ROOTS, DOCX_MCP_LOG_LEVEL)
│  ├─ security.py                 #   path sandboxing (resolve_safe_path)
│  ├─ ooxml.py                    #   (from Phase 2) generic OOXML/zip plumbing shared by every parsing module; (from Phase 4) atomic_write_part
│  ├─ document.py                 #   OOXML parsing/extraction for read_document; owns WordprocessingML paragraph rendering, reused by structure.py/metadata.py
│  ├─ structure.py                #   (from Phase 2) get_structure: heading resolution via styles.xml, tables, footnote-anchor index
│  ├─ metadata.py                 #   (from Phase 2) get_metadata: docProps/core.xml + word count
│  ├─ footnotes.py                #   (from Phase 3) get_footnotes: word/footnotes.xml content, resolved against document.py's shared anchor index
│  ├─ text_edit.py                #   (from Phase 4) find_text/replace_text: offset<->XML-node mapping, run-splitting replace algorithm
│  ├─ paragraph_edit.py           #   (from Phase 5) insert_paragraph/delete_paragraph: paragraph_index resolution, style-inheritance rules
│  ├─ footnote_edit.py            #   (from Phase 6) add_footnote/edit_footnote: footnote content building, id allocation, package wiring
│  └─ specs/                      #   per-tool specifications (one file per tool, from templates/tool-spec.md)
└─ tests/                         # mirrors src/docx_mcp/
   └─ fixtures/                   # small synthetic .docx fixtures — never real personal documents
```

> **Phase 0 note (historical):** as of Phase 0, only `README.md`, `CONTRIBUTING.md`, `Roadmap.md`, `docs/`, `templates/`, `.gitignore`, and `.env.example` existed as real content; `src/docx_mcp/`, `tests/`, and `tests/fixtures/` existed as empty folders (kept in git via `.gitkeep`). As of Phase 1 ([ADR-0001](adr/0001-ooxml-library-and-module-layout.md), [ADR-0002](adr/0002-dependency-and-lockfile-strategy.md)), `pyproject.toml` and the module layout above exist for real; the `.gitkeep` placeholders are gone. As of Phase 2 ([ADR-0003](adr/0003-phase-2-module-layout.md)), `document.py` is split: `ooxml.py` holds the generic zip/hardened-parser plumbing every parsing module shares, and `structure.py`/`metadata.py` are new tool-logic modules that reuse `document.py`'s paragraph-rendering helpers rather than duplicating them. As of Phase 3 ([ADR-0004](adr/0004-phase-3-footnote-module-and-shared-anchor-resolution.md)), `footnotes.py` is a new tool-logic module for `get_footnotes`; the footnote-anchor-finding loop `structure.py` previously implemented inline moved into `document.py` as a shared, presentation-agnostic helper (`find_footnote_anchors`), reused by both `structure.py` and `footnotes.py` so the two tools cannot disagree on which paragraph anchors which footnote id. As of Phase 4 ([ADR-0005](adr/0005-phase-4-text-edit-module-and-atomic-write.md)), `text_edit.py` is a new tool-logic module owning **both** `find_text` and `replace_text` (a deliberate, phase-scoped exception to the one-module-per-tool convention below, since the two tools share one offset-addressing contract); `ooxml.py` gains `atomic_write_part`, the project's first write-path plumbing, reused by every write tool from this phase onward. As of Phase 5 ([ADR-0006](adr/0006-phase-5-paragraph-edit-module-layout.md)), `paragraph_edit.py` is a new tool-logic module owning **both** `insert_paragraph` and `delete_paragraph` (the same kind of exception as `text_edit.py`, since both address the same top-level paragraph list by the same `paragraph_index` convention); it reuses `atomic_write_part` unchanged and `document.py`'s existing built-in-heading-id heuristic for its style-inheritance rule, without opening `word/styles.xml`. As of Phase 6 ([ADR-0007](adr/0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md)), `footnote_edit.py` is a new tool-logic module owning **both** `add_footnote` and `edit_footnote` (the same kind of exception, since both build a footnote's paragraph content from a `content` string and resolve a footnote by id); `ooxml.py` gains `atomic_write_parts`, a more permissive sibling of `atomic_write_part` (mirroring this module's existing `read_part`/`read_optional_part` split) that `add_footnote` uses to add `word/footnotes.xml` - plus, when it doesn't exist yet, its `[Content_Types].xml`/`.rels` wiring - together with the `word/document.xml` edit that references it, in one atomic step; `footnotes.py` gains an exported `render_footnote_content`, reused by `footnote_edit.py` rather than duplicated.

---

## 2. Conventions

- **Tool logic = one module** under `src/docx_mcp/`; `server.py` registers tools as **thin wrappers**, consistent with `research-graphrag`'s convention that the server module itself carries no business logic. `text_edit.py` (Phase 4), `paragraph_edit.py` (Phase 5), and `footnote_edit.py` (Phase 6) are documented exceptions, each covering **one closely-coupled tool pair** (`find_text`/`replace_text`; `insert_paragraph`/`delete_paragraph`; `add_footnote`/`edit_footnote`) rather than one tool — see [ADR-0005](adr/0005-phase-4-text-edit-module-and-atomic-write.md), [ADR-0006](adr/0006-phase-5-paragraph-edit-module-layout.md), [ADR-0007](adr/0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md); none of these is a precedent for merging unrelated tools into one module.
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
