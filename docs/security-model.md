# Security Model

docx-mcp reads and writes files on the local filesystem on behalf of an autonomous LLM agent. This document defines the threat model and the concrete rules every filesystem-facing and write-capable tool must follow (see [CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), principle 4). These rules are enforced and tested starting Phase 1, with a dedicated hardening/test pass in [Roadmap.md, Phase 7](../Roadmap.md#phase-7--security-hardening-path-sandboxing-atomic-writes) — they are a property every phase from Phase 1 onward is tested against, not a later addition.

---

## 1. Threat model

- **Path traversal / path escape.** A path argument supplied by the calling agent (directly, or indirectly via content the agent read elsewhere) could point outside the directory the caller intended the server to operate in — via `../` segments, an absolute path outside the allowed roots, or a symlink that resolves outside them. This is the most-cited concrete risk for filesystem-backed MCP servers (Almeida et al., *"MCP – Landscape, Security Threats, and Future Research Directions"*, arXiv:2503.23278, §5).
- **Path resolution triggering network access.** On Windows, canonicalizing a UNC path (`\\host\share\...`) makes the standard-library path resolution perform a real SMB/DNS network round-trip to `host` — confirmed during Phase 1 verification to block for several seconds against an unreachable hostname before any sandboxing check runs. A caller-supplied path must never be able to make the server itself originate a network connection, however transiently: that is both a denial-of-service vector (many such calls exhaust request-handling capacity) and, against a reachable or attacker-controlled host, a violation of the local-first / no-network-access principle ([CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), item 6). A UNC path is therefore rejected lexically, before any OS-level resolution is attempted.
- **Corruption from interrupted writes.** A crash, kill, or exception partway through writing a document must never leave the original file in a broken or partially written state.
- **Malformed or oversized input.** A `.docx` file that isn't a valid ZIP/XML structure, or one that is implausibly large, must be rejected with a clear error instead of crashing the server or exhausting memory.

**Out of scope:** network access (none — `stdio` transport only, no remote transport, see [README.md](../README.md#out-of-scope-for-now)); the UNC-path rule above exists precisely to keep that true even as an unintended side effect of path resolution, not as an exception to it. Also out of scope: the trustworthiness of a document's *content*. docx-mcp returns document text and structure as-is; judging whether text embedded in a document should be treated as instructions is the calling agent's responsibility, not something docx-mcp enforces.

---

## 2. Path-sandboxing rules

- Every filesystem-facing tool resolves its path argument against an explicit **allow-listed set of roots** (`DOCX_MCP_ALLOWED_ROOTS`, see [docs/documentation-standards.md](documentation-standards.md)). An empty allow-list means every path is denied — there is no implicit default root.
- A UNC network path (`\\host\share\...`, or `//host/share/...`) is rejected immediately, by lexical inspection only, **before** the canonicalization step below ever runs — see the threat-model entry above.
- Resolution **canonicalizes** the path (resolves `..` segments and symlinks) before comparing it against the allowed roots, so a symlink inside an allowed root cannot be used to escape it.
- A path that does not resolve inside at least one allowed root is rejected **before any file is opened** — the tool must not attempt to read or write it first and check afterward.
- The rejection error names the currently configured `DOCX_MCP_ALLOWED_ROOTS` roots (or states that none are configured), so the calling agent or the user reading the error can tell *why* the path was rejected without a second round trip. This is safe to disclose: the server is local-only, single-user `stdio` (see [CONTRIBUTING.md §1, item 6](../CONTRIBUTING.md#1-guiding-principles)), so the caller and the person who configured the allow-list are the same party — there is no confidentiality boundary between them for this server's own configuration.

## 3. Atomic-write contract

- Every write operation writes its result to a temporary file **in the same directory** as the target (so the final step is a rename on the same filesystem/volume), then atomically replaces the original.
- The original file is only ever touched by that final, atomic replace step.
- A crash or interruption at any point before the atomic replace leaves the original file byte-for-byte intact; there is no window in which the original is partially overwritten.

## 4. Input validation

- A `.docx` file that fails to open as a valid ZIP archive, or whose required internal XML parts are missing or malformed, is rejected with a clear, structured error.
- An implausibly large input file is rejected before it is fully parsed.
- **XML entity attacks.** Every internal XML part of a `.docx` that any tool parses — `word/document.xml` (Phase 1), `word/styles.xml` and `docProps/core.xml` (Phase 2, `get_structure`/`get_metadata`), `word/footnotes.xml` (Phase 3), `[Content_Types].xml` and `word/_rels/document.xml.rels` (Phase 6, `add_footnote`'s from-scratch footnote-package wiring) — is parsed through the single hardened-parser entry point `docx_mcp.ooxml.parse_xml` (see [ADR-0003](adr/0003-phase-2-module-layout.md)), which never resolves DTD-declared entities and never performs network or local-file access while parsing (`resolve_entities=False`, `no_network=True`; see [ADR-0001](adr/0001-ooxml-library-and-module-layout.md)). This blocks XML External Entity (XXE) attacks — where a crafted external entity reads an arbitrary local file or attempts network/SSRF access — and keeps libxml2's built-in entity-expansion ("billion laughs") limits active. A `.docx` is untrusted input in exactly the same sense as a path argument: it may be a file the calling agent read or was handed from an untrusted source, not something docx-mcp can assume is well-behaved.

## 5. Testing obligations

- Every filesystem-facing tool has a path-traversal-rejection test (`../`, absolute paths outside the allowed roots, symlink escapes) and a test proving a UNC path is rejected fast, without waiting on any network resolution.
- Every write-capable tool has a simulated-crash test proving the atomic-write contract (see [docs/testing.md §2.4](testing.md#24-error--edge-case-test) and [Roadmap.md, Phase 7](../Roadmap.md#phase-7--security-hardening-path-sandboxing-atomic-writes)).
- Every tool that parses a `.docx`'s internal XML has a test proving that a crafted external-entity declaration does not get resolved (e.g. an entity attempting to read a local file), from Phase 1 onward.
