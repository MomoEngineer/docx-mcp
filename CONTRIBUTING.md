# Contributing to docx-mcp

This document is the rulebook for working on **docx-mcp**. It describes *how* the repository is worked on and *what* is required of code, tests, and documentation — modeled on the structure and discipline of `research-graphrag`, adjusted to a small, focused MCP server.

---

## 1. Guiding principles

1. **Formatting fidelity first.** Edits are made by manipulating the underlying OOXML (the XML/ZIP structure inside every `.docx`) directly, not through a lossy high-level object model. An edit must never silently strip or reset formatting, styles, or structure it didn't touch.
2. **Footnotes are first-class, not an afterthought.** Every read path that exposes document text also resolves footnote markers to their content; footnotes get their own dedicated tool rather than being buried inside a generic "structure" blob.
3. **Spec and tests before code, always.** For every MCP tool: write the tool specification ([templates/tool-spec.md](templates/tool-spec.md)) first, then the tests from that spec (red), then the documentation, and only then the implementation (green). This is a deliberate departure from `research-graphrag`, where documentation follows implementation — here, spec, tests, and docs are written up front, in that order, before a line of implementation exists.
4. **Security by default, not by review.** docx-mcp reads and writes files on the local filesystem on behalf of an autonomous agent. Every filesystem-facing tool resolves paths against an explicit allow-listed root and rejects anything that escapes it (no directory traversal) — the single most-cited concrete risk for filesystem-backed MCP servers in the security literature (e.g. Almeida et al., *"MCP – Landscape, Security Threats, and Future Research Directions"*, arXiv:2503.23278, §5: a file-system server "that lacks directory res[triction]" allows access outside its intended scope). Writes are atomic (write to a temp file, then replace) so a crash mid-write can never corrupt the original document.
5. **Tools are specified, not just named.** A tool description must give an LLM caller what it needs to use the tool correctly on the first try: purpose, parameter semantics, limitations, and at least one example — the categories empirically shown to matter for agent tool-calling reliability (Hasan et al., *"MCP Tool Descriptions Are Smelly"*, arXiv:2602.14878). A tool spec missing any of these is incomplete, not just under-documented.
6. **Local-first.** The server runs over `stdio`. No network access, no remote transport — see [Roadmap.md](Roadmap.md) for why and what would have to change first.
7. **Traceability.** Architecturally relevant decisions get an [ADR](docs/adr/README.md). Not every decision needs one — right-sized, as `research-graphrag`'s CONTRIBUTING puts it — but anything that changes the shape of the OOXML handling, the tool surface, or the security model does.

---

## 2. Further documents

Phase 0 ([Roadmap.md](Roadmap.md#phase-0--repository-foundations)) establishes the following structure. Until each document exists, this file is the canonical rulebook for that area.

| Document | Content |
| --- | --- |
| [README.md](README.md) | What docx-mcp is, its feature scope, installation, and the tool list |
| [Roadmap.md](Roadmap.md) | Phased implementation plan, measured by Definition of Done per phase |
| `docs/repository-structure.md` | Binding folder structure and what belongs where |
| `docs/testing.md` | Test strategy: fixture-based, offline, spec-first (TDD) |
| `docs/documentation-standards.md` | Docstrings, typing, tool-spec requirements, security notes, logging |
| `docs/security-model.md` | Path-sandboxing rules, atomic-write contract, threat model |
| `docs/adr/README.md` | Process for Architecture Decision Records |
| `templates/tool-spec.md` | Tool specification template (input/output schema, limitations, examples) |
| `templates/adr-template.md` | ADR template |

---

## 3. Technical stack

- **Language:** Python, `requires-python = ">=3.11"`.
- **MCP server:** the official **MCP Python SDK** (`mcp`, including FastMCP), transport `stdio` — consistent with `research-graphrag`.
- **OOXML handling:** direct manipulation of the `.docx` ZIP/XML structure (not a high-level library like `python-docx`), so existing formatting, styles, tables, and footnotes survive edits untouched. The exact XML library (`lxml` vs. the standard-library `xml.etree`) is decided in Phase 1 via ADR, once the first real parsing/editing needs are known.
- **Tests:** `pytest`, offline, against small synthetic fixture `.docx` files checked into `tests/fixtures/` — never against real personal documents.
- **Static quality:** `ruff` (lint + format) and `mypy` (types).
- **Dependencies/lockfile:** decided in Phase 1, once the OOXML library choice is fixed.

---

## 4. Workflow (spec → tests → docs → implementation)

For every new MCP tool or non-trivial change:

1. **Document the decision** if it's architecturally relevant (ADR, see [docs/adr/README.md](docs/adr/README.md)).
2. **Write the tool specification first** ([templates/tool-spec.md](templates/tool-spec.md)) — the input/output schema is the single source of truth, and it exists before any code does.
3. **Write the tests from the spec**, before the implementation exists (contract tests, functional tests, error-path tests — see `docs/testing.md`). They are expected to fail (red) at this point.
4. **Write the documentation** (docstrings, module docs) alongside the spec and tests — not deferred until the code is green.
5. **Implement**, until the tests pass (green).
6. **Check the Definition of Done** (Section 5).

This order — spec, tests, docs, *then* code — is the one binding process rule of this project; it is stricter than `research-graphrag`'s (which documents after implementing) by explicit choice.

---

## 5. Definition of Done

A contribution counts as done when the applicable points are satisfied:

- [ ] Tool specification exists and was written **before** the implementation ([templates/tool-spec.md](templates/tool-spec.md)), covering purpose, parameters, limitations, and at least one example.
- [ ] Tests exist, were written from the spec before the implementation, and are green (`pytest`).
- [ ] For any change touching filesystem access: a test proving path-traversal / outside-root access is rejected.
- [ ] For any change touching document writes: a test proving a failed/interrupted write leaves the original file untouched.
- [ ] Full type hints; `mypy` clean.
- [ ] `ruff check` and `ruff format --check` clean.
- [ ] Docstrings for all public functions/tools, written alongside the spec (not after).
- [ ] Module documentation current for every touched module.
- [ ] For architecturally relevant decisions: an ADR exists.

---

## 6. Running quality checks locally

```bash
python -m ruff check .
python -m ruff format --check .
python -m mypy src
python -m pytest tests -q
```

---

## 7. Documentation language

Documentation, comments, and commit messages are written in **English**, consistent with [README.md](README.md) and [Roadmap.md](Roadmap.md). Code identifiers (function and variable names) are English.
