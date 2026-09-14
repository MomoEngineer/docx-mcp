# Tool Specification: `<tool_name>`

> Template for the per-tool specification. Copy this file to
> `src/docx_mcp/specs/<tool_name>.md` and fill it in **before** the
> implementation exists. It is the single source of truth for the tests
> written from it (see [docs/testing.md](../docs/testing.md)).

---

## Metadata

| Field | Value |
| --- | --- |
| **Tool name** | `<tool_name>` (e.g. `read_document`) |
| **Version** | `MAJOR.MINOR.PATCH` (see [docs/documentation-standards.md](../docs/documentation-standards.md), Section 4.1) |
| **Type** | Read / Write (see [README.md](../README.md#available-tools)) |
| **Status** | Draft / Implemented / Stable / Deprecated |

---

## 1. Purpose

One or two sentences: what the tool does and why it exists. (Doubles as the MCP tool description shown to the calling agent.)

## 2. Input Schema

| Parameter | Type | Required | Description / Value range |
| --- | --- | --- | --- |
| `example_param` | `str` | yes | ... |

## 3. Output Schema

Structure and meaning of the return value (fields, types, which are required). Sketch as JSON:

```json
{
  "field": "type and meaning"
}
```

> Input and output schema are additionally recorded as a **machine-readable JSON Schema** (single source of truth for contract tests, see [docs/testing.md](../docs/testing.md), Section 2.1).

## 4. Assumptions and Preconditions

- Expected input formats, preconditions, required resources (e.g. "the target path resolves inside an allowed root").

## 5. Limitations (Non-Goals)

- What the tool deliberately does **not** do. Boundary against neighboring tools (no overlap). Document known structural limitations here rather than treating them as bugs later (e.g. the table-of-contents-recompute boundary noted for `insert_paragraph`/`delete_paragraph` in [Roadmap.md](../Roadmap.md)).

## 6. Formatting & Safety Guarantees

- What formatting, styles, or structure this tool must never alter, even implicitly (see [CONTRIBUTING.md §1](../CONTRIBUTING.md#1-guiding-principles), principle 1).
- For write tools: confirmation that the tool follows the atomic-write contract ([docs/security-model.md](../docs/security-model.md)).
- For filesystem-facing tools: confirmation that the path is resolved against an allow-listed root before any file is opened ([docs/security-model.md](../docs/security-model.md)).

## 7. Error Behavior

- Defined error cases and how the tool surfaces them to the caller.
- Every filesystem-facing tool rejects a path that escapes its allowed root(s) (see [docs/security-model.md](../docs/security-model.md)).
- Malformed or implausibly large `.docx` input is rejected with a clear error instead of crashing the server.

## 8. Test Coverage

- Reference to `tests/<module>/test_<tool_name>.py`.
- Planned contract, functional, and error/edge-case tests (short list) — written from this spec, before the implementation (see [docs/testing.md](../docs/testing.md)).
