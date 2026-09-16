# 0002 – Dependency and Lockfile Strategy

- **Status:** Accepted
- **Date:** 2026-09-14

## Context

CONTRIBUTING.md §3 defers the dependency/lockfile strategy to Phase 1, tied to the OOXML library decision ([ADR-0001](0001-ooxml-library-and-module-layout.md)). CONTRIBUTING.md §6 already describes the local dev workflow as plain `python -m venv` + `pip`, not a third-party dependency manager; README.md's `uvx docx-mcp` is an **end-user installation** instruction (running a published package via `uv`'s tool runner), not a statement about how the project itself is developed. `uv` is not installed in this environment, and nothing else in the repository currently assumes it.

## Decision

- **Packaging:** `pyproject.toml` with PEP 621 metadata, build backend `hatchling`.
- **Runtime dependencies:** `mcp>=2.2,<3` (the official MCP Python SDK) and `lxml>=5`, generally pinned with a minimum-version constraint (`>=`) rather than an exact pin, since docx-mcp is a library/server meant to be installed into a caller's own environment (`uvx docx-mcp`, `pip install docx-mcp`) where over-constraining would cause resolver conflicts. `mcp` is the one exception with an explicit upper bound (`<3`): its ergonomic server API moved from `mcp.server.fastmcp.FastMCP` (1.x) to `mcp.server.mcpserver.MCPServer` (2.x) — a breaking rename this project hit directly while implementing `server.py` — so an unbounded `>=` would let a resolver install a 1.x release this code cannot import, or a future 3.x with the same kind of break. `lxml` has no such observed history, so it keeps a plain lower bound.
- **Dev dependencies:** declared as an optional-dependency group `dev` (`pip install -e .[dev]`): `pytest`, `ruff`, `mypy`, `lxml-stubs` (type stubs for `mypy`, since `lxml` ships without inline types), and `python-docx` (fixture generation only, never imported by `src/docx_mcp/`).
- **No dedicated lockfile tool** (no `uv.lock`, no `poetry.lock`, no `pip-tools` compile step) for now. Reproducibility for contributors is achieved by pinning minimum versions in `pyproject.toml` and running tests in CI/locally against whatever those constraints resolve to.

## Alternatives

- **`uv` + `uv.lock`.** Would give exact reproducibility and matches the `uvx docx-mcp` end-user path, but requires installing a new tool on every contributor's machine and would require rewriting CONTRIBUTING.md §6's already-documented venv/pip workflow. Rejected as disproportionate for a project at Phase 1 with two runtime dependencies; revisit if dependency-resolution drift actually causes a problem (open follow-up, not a blocker).
- **`pip-tools` (`requirements.in` → compiled `requirements.txt`).** Solves reproducibility without a new runtime tool, but adds a compile step contributors must remember to re-run. Rejected for the same "not proportionate yet" reason.
- **Exact version pins in `pyproject.toml` itself.** Rejected because docx-mcp is distributed as an installable package (not an application with its own deployment), so exact pins would fight the installing environment's own resolver.

## Consequences

- **Positive:** the dependency setup matches the dev workflow CONTRIBUTING.md §6 already documents; no new tool required to contribute; minimal-version pins keep the package installable alongside other packages in a caller's environment.
- **Negative / Effort:** no automatic detection of a dependency update silently breaking the build between CI runs on different days (mitigated by `pytest`/`mypy`/`ruff` all running on every change, which would catch a breaking dependency update as a test failure).
- **Follow-up decisions:** if the project later needs exact reproducibility (e.g. a security incident traced to an unpinned transitive dependency), introduce `uv.lock` via a new ADR superseding this one — not a silent change.

## Addendum (found during a documentation audit, not a new decision)

`src/docx_mcp/server.py` imports `pydantic.BaseModel`/`Field` directly (for every tool's result model, since Phase 1) but `pydantic` was never listed in `pyproject.toml`'s `dependencies` — it installed correctly only because `mcp` itself depends on `pydantic>=2.12.0`, a transitive dependency this project's own directly-imported code should not rely on silently. Corrected by adding `pydantic>=2.12` (matching `mcp`'s own floor) to `dependencies` directly. This does not change the Decision above (no new dependency was added — `pydantic` was always resolved and installed, per `mcp`'s requirement) or require a new ADR (declaring an already-present transitive dependency explicitly is a correctness fix, not an architectural change, per [docs/adr/README.md §1](README.md#1-when-is-an-adr-required)).
