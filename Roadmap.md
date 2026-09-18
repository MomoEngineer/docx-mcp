# Roadmap – docx-mcp

Phased implementation plan for docx-mcp. The plan is **iterative**: a thin, working vertical slice first, then targeted extensions. **Deliberately no time estimates** — progress is measured by the "Definition of Done" (DoD) per phase.

> Complements the [README](README.md) and [CONTRIBUTING.md](CONTRIBUTING.md).

---

## Guiding principles

- **Formatting fidelity first:** edits manipulate the underlying OOXML directly; nothing an edit doesn't touch may change.
- **Footnotes are first-class:** findable, resolvable to their anchor, and — once Phase 6 lands — editable, not a side effect of generic text handling.
- **Spec and tests before code:** every tool gets a written specification and failing tests before it gets an implementation (see [CONTRIBUTING.md](CONTRIBUTING.md#4-workflow-spec--tests--docs--implementation)).
- **Security by default:** filesystem access is sandboxed to an explicit root; writes are atomic. Not a later hardening pass grafted on, but a property every phase from Phase 1 onward is tested against.
- **Local-first:** `stdio` transport only; no network access anywhere in the server.
- **Small, precisely specified tools:** a tool a caller can use correctly from its description alone beats a large tool that does many things ambiguously.

---

## Phase 0 – Repository Foundations

**Goal:** the organizational and documentation structure of the repository — folder layout, templates, and rules — before any feature code exists. No document parsing, no MCP server, no tool implementation in this phase.

**Status:**

| Deliverable | Status |
| --- | --- |
| [README.md](README.md) — project scope, features, tool list | Done |
| [CONTRIBUTING.md](CONTRIBUTING.md) — rulebook, workflow, Definition of Done | Done |
| [Roadmap.md](Roadmap.md) — this document | Done |
| Folder skeleton (`src/docx_mcp/`, `tests/`, `tests/fixtures/`, `docs/`, `docs/adr/`, `templates/`) | Done |
| `templates/tool-spec.md` — tool specification template | Done |
| `templates/adr-template.md` — ADR template | Done |
| `docs/adr/README.md` — ADR process | Done |
| `docs/repository-structure.md` — binding folder structure | Done |
| `docs/testing.md` — test strategy (fixture-based, offline, spec-first) | Done |
| `docs/documentation-standards.md` — docstrings, typing, tool-spec, security, logging | Done |
| `docs/security-model.md` — path-sandboxing rules, atomic-write contract, threat model | Done |

**Definition of Done:** every deliverable above exists, cross-references the others where relevant, and no deliverable describes a rule that a later phase then has to contradict. No `pyproject.toml`, no source code — that begins in Phase 1.

---

## Phase 1 – Thin Vertical Slice: Server Skeleton + `read_document`

**Goal:** a minimal, actually running MCP server, reachable over `stdio`, exposing exactly one tool — `read_document` — that opens a real `.docx` fixture and returns its full text with heading and footnote markers inline. This proves the raw-OOXML approach end-to-end before anything else is built on top of it.

**Scope:**
- `pyproject.toml`, package skeleton `src/docx_mcp/`.
- ADR deciding the OOXML library (`lxml` vs. standard-library `xml.etree`) — the first real architectural decision, deferred from Phase 0 on purpose because it needs a real parsing need to evaluate against.
- Small synthetic fixture `.docx` files in `tests/fixtures/` (headings, a footnote, plain paragraphs — not a real personal document).
- Tool spec for `read_document`, written first.
- Tests against the fixtures, written from the spec, before implementation.
- Docstrings and a module doc for the new package.

**Definition of Done:** `pip install -e .` works; the server starts over `stdio` and registers `read_document`; the tool returns correct text plus inline heading/footnote markers against the fixtures; tests are green; spec and docs exist and predate the implementation (per [CONTRIBUTING.md](CONTRIBUTING.md)).

---

## Phase 2 – Structural Read: `get_structure`, `get_metadata`

**Goal:** expose the document's structure — heading hierarchy, table-of-contents entries, table contents rendered as text, a paragraph index, and a lightweight footnote-anchor index (id + anchor paragraph, no resolved content — see [specs/get_structure.md §3](src/docx_mcp/specs/get_structure.md#3-output-schema) for why this doesn't pull Phase 3's content-resolution scope forward) — plus core document metadata (title, author, created/modified, word count). Still read-only.

**Scope:** tool specs and tests for `get_structure` and `get_metadata`, written first; fixtures extended with a table and a multi-level heading structure; implementation against those fixtures.

**Definition of Done:** `get_structure` correctly reports heading levels (resolved via `styles.xml`'s `outlineLvl`/`basedOn` chain, not just the built-in style-id heuristic), table-of-contents entries, table content as text, and the footnote-anchor index against the fixtures; `get_metadata` reports core properties and word count correctly; tests green; no regression in Phase 1's `read_document`.

---

## Phase 3 – Footnotes, Read Path: `get_footnotes`

**Goal:** a dedicated tool that lists every footnote with its ID, its anchor location in the text, and its content — and confirm that `read_document`'s inline footnote markers (Phase 1) and `get_structure`'s footnote list (Phase 2) all resolve consistently against the same data.

**Scope:** tool spec and tests for `get_footnotes` first; fixtures extended with multiple footnotes, including two footnotes anchored in the same paragraph (the edge case most likely to break a naive anchor-to-paragraph mapping).

**Definition of Done:** `get_footnotes` correctly maps every footnote to its anchor paragraph, including the multi-footnote-per-paragraph fixture; `read_document` and `get_structure` output stays consistent with it; tests green.

---

## Phase 4 – Precise Text Edits: `find_text`, `replace_text`

**Goal:** locate and replace text reliably even when Word has split it across multiple internal XML runs — a common, easy-to-get-wrong source of silent corruption in naive `.docx` editors — without altering the formatting of anything not touched.

**Scope:** tool specs and tests first, including a fixture deliberately constructed with text split across runs (e.g. produced by spell-check or an earlier manual edit in Word); `find_text` returns stable, addressable locations that `replace_text` can act on unambiguously.

**Definition of Done:** `replace_text` round-trips correctly on the run-split fixture without losing or altering surrounding formatting; `find_text` locations remain valid inputs to `replace_text`; tests green; no regression in Phases 1–3.

---

## Phase 5 – Paragraph-Level Structure Edits: `insert_paragraph`, `delete_paragraph`

**Goal:** insert and delete whole paragraphs — body text or headings — with the new paragraph inheriting the style of its surrounding context.

**Scope:** tool specs and tests first; fixtures covering insertion after a heading, after a body paragraph, and at the end of the document; explicit scope note that a Word table-of-contents field is **not** recomputed by this tool (it requires a field refresh inside Word itself) — documented as a boundary, not fixed as a bug.

**Definition of Done:** inserted paragraphs carry the correct style (`pStyle`) for their context; `delete_paragraph` removes a paragraph without leaving orphaned XML *elements* in `word/document.xml` itself (no dangling siblings, no malformed tree) — a footnote a deleted paragraph anchored is a deliberate, documented exception: its `<w:footnote>` entry is left in `word/footnotes.xml`, orphaned but harmless, per [ADR-0006 §Decision.5](docs/adr/0006-phase-5-paragraph-edit-module-layout.md#decision); the TOC-non-recompute limitation is documented in the tool spec; tests green; no regression in Phases 1–4.

---

## Phase 6 – Footnote Editing: `add_footnote`, `edit_footnote`

**Goal:** extend `footnotes.xml`, the relationship wiring, and the in-text anchor consistently when a footnote is added or edited, keeping footnote IDs unique and valid.

**Scope:** tool specs and tests first; a round-trip fixture test that re-parses the written document and confirms the new/edited footnote resolves correctly through `get_footnotes` and `read_document`.

**Definition of Done:** a newly added footnote is discoverable and correctly anchored via `get_footnotes`; an edited footnote's content is updated without breaking its anchor or ID; tests green; no regression in Phases 1–5.

---

## Phase 7 – Security Hardening: Path Sandboxing, Atomic Writes

**Goal:** enforce the security principles from [CONTRIBUTING.md](CONTRIBUTING.md#1-guiding-principles) as tested behavior, not just as a written rule: every filesystem-facing tool resolves its path against an explicit allow-listed root and rejects anything that escapes it; every write goes through a temp-file-then-replace sequence so a crash mid-write can never corrupt the original document; malformed or implausibly large `.docx` files are rejected with a clear error instead of crashing the server.

**Scope:** a path-sandboxing test suite covering traversal attempts (`../`, absolute paths outside the root, symlink escapes) across every tool that touches the filesystem; a simulated-crash test for the atomic-write path; a review of every tool description against the "smell" categories from [CONTRIBUTING.md §1.5](CONTRIBUTING.md#1-guiding-principles) (purpose, parameters, limitations, examples).

**Definition of Done:** every filesystem-facing tool has a passing path-traversal-rejection test; a simulated crash mid-write leaves the original file byte-for-byte intact; every tool spec covers purpose, parameters, limitations, and at least one example; no regression in Phases 1–6.

---

## Phase 8 – QA & Documentation Pass

**Goal:** a full Definition-of-Done sweep across the whole tool surface before calling the initial tool set complete: every spec matches its implementation, every module has current documentation, the README's tool table matches reality, and the ADR log is internally consistent.

**Scope:** cross-check every tool spec against its actual input/output schema; cross-check the README tool table; review the ADR log for contradictions; a coverage pass against the guideline in `docs/testing.md`.

**Definition of Done:** `pytest`, `ruff check`, `ruff format --check`, and `mypy` all clean; every tool spec, module doc, and the README tool table agree with the implementation; no open contradictions between ADRs.

---

## Phase 9 – Scoped Paragraph-Range Reads

**Goal:** let `read_document` and `get_structure` return a caller-chosen `[start_paragraph, end_paragraph)` slice instead of always the whole document, so checking a growing document before the next edit costs tokens proportional to what changed, not to the document's total size — see [ADR-0008](docs/adr/0008-scoped-paragraph-range-reads.md).

**Scope:** a shared range-resolution helper (`docx_mcp.document.resolve_paragraph_range`, Python-slice-like: out-of-range bounds clamp, only a negative or reversed range is rejected); both tools' specs and tests updated first; `get_structure` gains a `total_paragraphs` output field; `read_document`'s output schema is deliberately left unchanged (ADR-0008, point 3).

**Definition of Done:** both tools accept `start_paragraph`/`end_paragraph`; omitting both reproduces every existing test's expected output byte-for-byte (no regression); `get_structure`'s `toc`/`footnotes` are scoped consistently with `paragraphs`, `tables` is documented as unscoped; a range whose `end_paragraph` exceeds the document's length returns the available tail instead of erroring; a negative `start_paragraph` or `end_paragraph < start_paragraph` is rejected with a clear error; tests green; `ruff`/`mypy` clean; both tool specs' Version bumped to `1.1.0`.

---

## Deliberately excluded (for now)

Table structure edits (rows/columns), image insertion or replacement, document theme/style changes, comments, tracked changes, and remote/HTTP transport are out of scope for this roadmap — see the README's [Out of scope](README.md#out-of-scope-for-now) section. Each would need its own phase, its own ADR, and its own case for why the added complexity is worth it; none is assumed as a foregone next step.
