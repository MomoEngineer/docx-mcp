# Test Strategy

This document defines how code and MCP tools are tested. Offline-capable and right-sized (see [CONTRIBUTING.md](../CONTRIBUTING.md)).

> **Core rule:** for every non-trivial MCP tool, tests are written **from its specification, before the implementation exists** (red → green). No tool counts as done without tests (see [CONTRIBUTING.md §5](../CONTRIBUTING.md#5-definition-of-done)).

---

## 1. Test scope in the repo

All tests live under `tests/`, mirroring the package structure:

```
tests/
├─ fixtures/         # small, synthetic .docx fixtures — never real personal documents
├─ integration/      # cross-tool tests (from Phase 2) — see §3 below
└─ ...                # mirrors src/docx_mcp/, populated from Phase 1 onward
```

As of Phase 0, only `tests/` and `tests/fixtures/` existed, as empty folders. Phase 1 added `tests/fixtures/minimal.docx` (headings, a footnote, plain paragraphs) and the tests that read it. Phase 2 added `tests/fixtures/structured.docx` (a table with a merged cell, three heading levels including one custom-named style, deterministic core properties) and `tests/integration/`; fixtures are extended per later phase as new tools need them — multiple footnotes including two anchored in the same paragraph in Phase 3, text split across runs in Phase 4, and so on (see [Roadmap.md](../Roadmap.md)).

---

## 2. Test types (per MCP tool)

> **Order (test-first):** contract and error/edge-case tests are derived **from the tool specification, before the core logic exists** (red → green). This requires a minimal, registered tool skeleton first.

### 2.1 Contract test (MCP protocol)
- The tool is found during discovery (`list_tools`).
- Name, description, **input schema**, and **tool version** match the specification.
- The **output** matches the documented output schema.

### 2.2 Functional / happy-path test
- Representative, valid input produces the expected result, checked against small, versioned fixtures.

### 2.3 Formatting-fidelity test

This category has no equivalent in a plain text-processing tool and is docx-mcp's core guarantee (see [CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), principle 1): after an edit, everything the edit didn't touch — styles, formatting, tables, images, unrelated paragraphs and runs — must be provably unchanged. Where feasible this is checked by comparing the OOXML of untouched parts before and after the edit, not just by re-reading the visible text.

### 2.4 Error / edge-case test
- Invalid input produces a clear, well-formed error.
- Security boundaries (paths escaping the allowed root) are rejected — see [docs/security-model.md](security-model.md).
- Malformed or implausibly large `.docx` input is rejected without crashing the server.

---

## 3. Integration tests (cross-tool)

For chains where tools build on each other — e.g. `insert_paragraph` followed by `read_document` must show the new paragraph, or `add_footnote` followed by `get_footnotes` must resolve it consistently (see [Roadmap.md](../Roadmap.md), Phase 3 and Phase 6) — integration tests are added under `tests/integration/`.

The same folder also holds tests for tools that expose the *same* underlying data two independently-computed ways without one calling the other, where staying consistent (or, where a spec deliberately allows divergence, staying divergent in exactly the documented place) is itself part of the contract: `tests/integration/test_read_document_and_get_structure_consistency.py` (Phase 2) checks that `read_document`'s inline markers and `get_structure`'s `heading_level`/`footnotes` agree everywhere the specs promise, and pins the one documented exception (custom-named heading styles, see [specs/read_document.md §5](../src/docx_mcp/specs/read_document.md#5-limitations-non-goals)) so it stays an intentional, tested contract rather than something either tool's implementation could silently drift into or out of.

---

## 4. Tooling and configuration

- **Test runner:** `pytest` (`python -m pytest tests -q`).
- **Static checks:** `ruff check`, `ruff format --check`, `mypy src` (see [CONTRIBUTING.md §6](../CONTRIBUTING.md#6-running-quality-checks-locally)).
- Test dependencies (`pytest`) and the "no dedicated lockfile" strategy were decided in Phase 1 alongside the OOXML library ADR (see [CONTRIBUTING.md §3](../CONTRIBUTING.md#3-technical-stack), [ADR-0002](adr/0002-dependency-and-lockfile-strategy.md)).

No fixed line-coverage threshold is mandated; the binding bar is the Definition of Done in [CONTRIBUTING.md §5](../CONTRIBUTING.md#5-definition-of-done) — every path described in a tool's specification is covered by a green test.

---

## 5. Principles for good tests

1. **Deterministic.** No unmanaged dependency on the system clock, randomness, or network.
2. **Isolated.** No state leaks between tests; each test uses its own fixture copy so edits (including write tools) never mutate a shared fixture on disk.
3. **Meaningful.** A failing test names the violated expectation clearly.
4. **Reproducible.** Fixtures are versioned; runs are repeatable on another machine.
5. **Close to the specification.** Tests check exactly the inputs/outputs the spec promises.
