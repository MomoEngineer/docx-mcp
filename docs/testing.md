# Test Strategy

This document defines how code and MCP tools are tested. Offline-capable and right-sized (see [CONTRIBUTING.md](../CONTRIBUTING.md)).

> **Core rule:** for every non-trivial MCP tool, tests are written **from its specification, before the implementation exists** (red → green). No tool counts as done without tests (see [CONTRIBUTING.md §5](../CONTRIBUTING.md#5-definition-of-done)).

---

## 1. Test scope in the repo

All tests live under `tests/`, mirroring the package structure:

```
tests/
├─ fixtures/         # small, synthetic .docx fixtures — never real personal documents
└─ ...                # mirrors src/docx_mcp/, populated from Phase 1 onward
```

As of Phase 0, only `tests/` and `tests/fixtures/` exist, as empty folders. Fixtures are added starting Phase 1 (headings, a footnote, plain paragraphs), extended per phase as new tools need them — a table and multi-level headings in Phase 2, multiple footnotes including two anchored in the same paragraph in Phase 3, text split across runs in Phase 4, and so on (see [Roadmap.md](../Roadmap.md)).

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

---

## 4. Tooling and configuration

- **Test runner:** `pytest` (`python -m pytest tests -q`).
- **Static checks:** `ruff check`, `ruff format --check`, `mypy src` (see [CONTRIBUTING.md §6](../CONTRIBUTING.md#6-running-quality-checks-locally)).
- Test dependencies and any lockfile strategy are decided in Phase 1, alongside the OOXML library ADR (see [CONTRIBUTING.md §3](../CONTRIBUTING.md#3-technical-stack)).

No fixed line-coverage threshold is mandated; the binding bar is the Definition of Done in [CONTRIBUTING.md §5](../CONTRIBUTING.md#5-definition-of-done) — every path described in a tool's specification is covered by a green test.

---

## 5. Principles for good tests

1. **Deterministic.** No unmanaged dependency on the system clock, randomness, or network.
2. **Isolated.** No state leaks between tests; each test uses its own fixture copy so edits (including write tools) never mutate a shared fixture on disk.
3. **Meaningful.** A failing test names the violated expectation clearly.
4. **Reproducible.** Fixtures are versioned; runs are repeatable on another machine.
5. **Close to the specification.** Tests check exactly the inputs/outputs the spec promises.
