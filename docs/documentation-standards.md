# Documentation Standards

This document defines the requirements for documenting code and results. The goal is consistent **traceability and maintainability**. Right-sized for a small, focused MCP server (see [CONTRIBUTING.md](../CONTRIBUTING.md)).

---

## 1. Docstrings and typing

- **Type hints are mandatory** for all public functions, tools, and return values. `python -m mypy src` must pass without errors.
- **Docstrings are mandatory** for every publicly exposed MCP tool, and for public functions and classes.
- A tool docstring contains at least: **purpose** (one sentence, doubles as the MCP tool description), **parameters**, **return value**, **errors/limitations**, and **at least one concrete usage example**. The docstring — not the markdown spec — is what an LLM caller actually sees as the MCP tool description at call time, so it is the artifact these categories must hold on, not just the human-facing spec file; the example categories are the ones empirically shown to matter for agent tool-calling reliability (see [CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), principle 5, citing Hasan et al., *"MCP Tool Descriptions Are Smelly"*, arXiv:2602.14878).
- Style: consistent across the repo (Google style recommended).

---

## 2. Per-tool specification

- Every **non-trivial** MCP tool has a specification following [templates/tool-spec.md](../templates/tool-spec.md), under `src/docx_mcp/specs/<tool>.md`. The input/output schema is the single source of truth for tests and contract checks; the specification exists **before** the code (see [CONTRIBUTING.md §4](../CONTRIBUTING.md#4-workflow-spec--tests--docs--implementation)).
- If spec and code diverge, the **spec is the binding contract**; the implementation is corrected to match it, not the other way around.

---

## 3. Architecture Decision Records (ADRs)

Architectural and policy decisions are recorded as an ADR (process/template: [docs/adr/README.md](adr/README.md), [templates/adr-template.md](../templates/adr-template.md)).

ADR-required, among others:

- Choosing or changing the OOXML handling library.
- Choosing or changing the dependency/lockfile strategy.
- Choosing a language other than Python for a component.
- Operating the server over a transport other than `stdio`.
- Deliberately accepted redundancy.

---

## 4. Reproducibility and tool versioning

- **Version pinning:** minimum-version constraints in `pyproject.toml`, no dedicated lockfile tool for now — decided in Phase 1 via [ADR-0002](adr/0002-dependency-and-lockfile-strategy.md).
- Run metadata is recorded wherever a result could vary: tool version, timestamp.

### 4.1 Tool versioning

- Every MCP tool carries a **semantic version** (`MAJOR.MINOR.PATCH`) in its tool specification and exposes it via MCP.
- **MAJOR** for a breaking schema change, **MINOR** for a backward-compatible extension, **PATCH** for internal fixes.

---

## 5. Configuration and secrets

- Server/tool configuration is read from **environment variables**; safe defaults are documented alongside the config (see [.env.example](../.env.example)).
- Naming convention: `DOCX_MCP_ALLOWED_ROOTS` (allow-listed filesystem roots, `os.pathsep`-separated; empty means access is denied — see [docs/security-model.md](security-model.md)), `DOCX_MCP_LOG_LEVEL` (default `INFO`).
- **No secrets** (tokens, passwords, API keys) in code, tool specs, README, or MCP client config — docx-mcp has no external service dependency that would need one.
- A tool **must not** access resources outside the allowed roots.

---

## 6. Logging and observability (baseline)

- Logging via the standard `logging` module, to **stderr** only (never stdout — with `stdio` transport, stdout is reserved for the MCP protocol).
- **No secrets and no document content** in the log.
- Recommended default level `INFO`, configurable via `DOCX_MCP_LOG_LEVEL`.
- Errors are logged with the error category from the tool's error behavior section (see [templates/tool-spec.md](../templates/tool-spec.md)).

---

## 7. Documentation language

Documentation, comments, and commit messages are written in **English** (see [CONTRIBUTING.md §7](../CONTRIBUTING.md#7-documentation-language)); code identifiers are English.
