# Security Model

docx-mcp reads and writes files on the local filesystem on behalf of an autonomous LLM agent. This document defines the threat model and the concrete rules every filesystem-facing and write-capable tool must follow (see [CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), principle 4). These rules are enforced and tested starting Phase 1, with a dedicated hardening/test pass in [Roadmap.md, Phase 7](../Roadmap.md#phase-7--security-hardening-path-sandboxing-atomic-writes) — they are a property every phase from Phase 1 onward is tested against, not a later addition.

---

## 1. Threat model

- **Path traversal / path escape.** A path argument supplied by the calling agent (directly, or indirectly via content the agent read elsewhere) could point outside the directory the caller intended the server to operate in — via `../` segments, an absolute path outside the allowed roots, or a symlink that resolves outside them. This is the most-cited concrete risk for filesystem-backed MCP servers (Almeida et al., *"MCP – Landscape, Security Threats, and Future Research Directions"*, arXiv:2503.23278, §5).
- **Corruption from interrupted writes.** A crash, kill, or exception partway through writing a document must never leave the original file in a broken or partially written state.
- **Malformed or oversized input.** A `.docx` file that isn't a valid ZIP/XML structure, or one that is implausibly large, must be rejected with a clear error instead of crashing the server or exhausting memory.

**Out of scope:** network access (none — `stdio` transport only, no remote transport, see [README.md](../README.md#out-of-scope-for-now)); the trustworthiness of a document's *content*. docx-mcp returns document text and structure as-is; judging whether text embedded in a document should be treated as instructions is the calling agent's responsibility, not something docx-mcp enforces.

---

## 2. Path-sandboxing rules

- Every filesystem-facing tool resolves its path argument against an explicit **allow-listed set of roots** (`DOCX_MCP_ALLOWED_ROOTS`, see [docs/documentation-standards.md](documentation-standards.md)). An empty allow-list means every path is denied — there is no implicit default root.
- Resolution **canonicalizes** the path (resolves `..` segments and symlinks) before comparing it against the allowed roots, so a symlink inside an allowed root cannot be used to escape it.
- A path that does not resolve inside at least one allowed root is rejected **before any file is opened** — the tool must not attempt to read or write it first and check afterward.

## 3. Atomic-write contract

- Every write operation writes its result to a temporary file **in the same directory** as the target (so the final step is a rename on the same filesystem/volume), then atomically replaces the original.
- The original file is only ever touched by that final, atomic replace step.
- A crash or interruption at any point before the atomic replace leaves the original file byte-for-byte intact; there is no window in which the original is partially overwritten.

## 4. Input validation

- A `.docx` file that fails to open as a valid ZIP archive, or whose required internal XML parts are missing or malformed, is rejected with a clear, structured error.
- An implausibly large input file is rejected before it is fully parsed.

## 5. Testing obligations

- Every filesystem-facing tool has a path-traversal-rejection test (`../`, absolute paths outside the allowed roots, symlink escapes).
- Every write-capable tool has a simulated-crash test proving the atomic-write contract (see [docs/testing.md §2.4](testing.md#24-error--edge-case-test) and [Roadmap.md, Phase 7](../Roadmap.md#phase-7--security-hardening-path-sandboxing-atomic-writes)).
