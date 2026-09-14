# 0001 – OOXML Handling Library and Module Layout

- **Status:** Accepted
- **Date:** 2026-09-14

## Context

Phase 1 ([Roadmap.md](../../Roadmap.md#phase-1--thin-vertical-slice-server-skeleton--read_document)) requires the first tool, `read_document`, to open a real `.docx` file and extract its text with heading and footnote markers inline. This is the first point at which the project needs to parse OOXML (the XML/ZIP structure inside a `.docx`), so it is the first point at which the choice between `lxml` and the standard-library `xml.etree.ElementTree` (CONTRIBUTING.md §3) can be made against a real need rather than a guess.

Two concerns shape the decision:

1. **Forward fit with later phases.** Phase 2 (`get_structure`) needs heading-hierarchy and table extraction; Phase 4 (`find_text`/`replace_text`) needs to locate and rewrite text that Word has split across multiple `w:r` runs; Phase 5 needs paragraph-level insertion/deletion that preserves surrounding `pStyle`; Phase 6 needs to extend `footnotes.xml` and its relationship wiring. All of these are namespace-heavy, XPath-shaped problems (e.g. "find every `w:r` under this `w:p` that isn't inside a `w:del`").
2. **Security.** docx-mcp parses `.docx` files that may originate from an autonomous agent's filesystem access, i.e. untrusted input. Generic XML parsing is subject to XML External Entity (XXE) expansion (a maliciously crafted external entity can read arbitrary local files or, if network access were allowed, exfiltrate data) and entity-expansion ("billion laughs") denial-of-service. Neither `docs/security-model.md` nor `CONTRIBUTING.md` covered this threat category prior to this ADR — a gap this ADR closes together with the library choice (see [docs/security-model.md §4](../security-model.md#4-input-validation)). This class of "minimal server, outsized blast radius" risk matches the broader pattern documented for filesystem-adjacent MCP servers (Almeida et al., *"MCP – Landscape, Security Threats, and Future Research Directions"*, arXiv:2503.23278, §5; consistent with the attack patterns in Rehberger, *"MCP Safety Audit"*, arXiv:2504.03767, and Wu et al., *"Trivial Trojans"*, arXiv:2507.19880).

## Decision

Use **`lxml`** (`lxml.etree`) for all OOXML parsing and manipulation, with a hardened parser configuration everywhere a `.docx`-supplied XML part is parsed:

```python
def _hardened_parser() -> etree.XMLParser:
    return etree.XMLParser(
        resolve_entities=False,  # never resolve DTD-declared entities (blocks XXE file/SSRF reads)
        no_network=True,  # never fetch external DTDs/entities over network or file access
        huge_tree=False,  # keep libxml2's built-in entity-expansion/depth limits enabled
        dtd_validation=False,
        load_dtd=False,
    )
```

`resolve_entities=False` is set explicitly rather than relying on `lxml`'s default, since the default has changed across `lxml` releases and must not be assumed. `huge_tree=False` (the default) keeps libxml2's built-in protection against entity-expansion ("billion laughs") attacks active; it is never set to `True` for a `.docx`-supplied part. This lives in one small parser-factory function (`docx_mcp.document._hardened_parser()`) so every call site gets the same hardened configuration — there is exactly one place where a `.docx` XML part is turned into a tree.

**Module layout for Phase 1** (the part of this decision that answers `docs/repository-structure.md`'s deferred "exact submodule layout" question):

```
src/docx_mcp/
├── __init__.py     # package version
├── __main__.py     # entry point (python -m docx_mcp)
├── server.py       # MCPServer instance; registers tools as thin wrappers, no business logic
├── config.py       # environment-variable configuration (DOCX_MCP_ALLOWED_ROOTS, DOCX_MCP_LOG_LEVEL)
├── security.py     # path sandboxing (resolve_safe_path)
├── document.py     # opens a .docx as a zip, parses word/document.xml with lxml, extracts text with inline heading/footnote-reference markers (footnote content itself is Phase 3's get_footnotes, not this module)
└── specs/
    └── read_document.md
```

`document.py` is deliberately a single module for Phase 1's thin slice — splitting OOXML handling into further submodules (e.g. separate parsing vs. rendering vs. footnote resolution) is deferred until Phase 2 or 3 introduce a second read path that would otherwise duplicate logic against the same module. This keeps with the "no abstraction before it's needed" principle; it is not meant to be the final shape of the package.

## Alternatives

- **`xml.etree.ElementTree` (standard library).** Rejected for two reasons: (1) the Python documentation itself states it "is not secure against maliciously constructed data" and recommends `defusedxml` for untrusted input — adding `defusedxml` as a dependency anyway erodes the "no extra dependency" argument for the standard library; (2) it lacks full XPath 1.0 support (only a small subset), which later phases need for precise run/paragraph queries. Splitting the OOXML library between "stdlib for reading, `lxml` for writing" was considered and rejected as unnecessary complexity — one library, one parsing contract, is easier to keep secure and consistent.
- **`defusedxml.ElementTree`.** Solves the XXE problem for the standard library's API but keeps its limited XPath support; doesn't change the Phase 4–6 ergonomics problem. Rejected for the same XPath reason as plain `xml.etree`.
- **A high-level library (`python-docx`).** Out of scope by the project's founding principle (README.md "Why", CONTRIBUTING.md §1.1) — it round-trips through a lossy object model. Retained only as a **dev-only dependency** for generating test fixtures (see [ADR-0002](0002-dependency-and-lockfile-strategy.md)), never imported by the server itself.

## Consequences

- **Positive:** one consistent, namespace-aware, XPath-capable library for all current and future OOXML work; the XXE/entity-expansion gap in the threat model is closed at the same time the parser is introduced, not retrofitted later; the module layout question from Phase 0 is now answered and traceable.
- **Negative / Effort:** `lxml` is a compiled dependency (ships as a wheel on the supported platforms/Python versions; no C toolchain needed for `pip install`, but it is a larger dependency than the standard library).
- **Follow-up decisions:** none required by this ADR. If a future phase needs schema validation (`.xsd`) or C14N canonicalization for formatting-fidelity tests, that can reuse the same `lxml` dependency without a new ADR.
