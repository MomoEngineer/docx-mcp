# 0008 – Scoped Paragraph-Range Reads for `read_document` and `get_structure`

- **Status:** Accepted
- **Date:** 2026-09-18

## Context

`read_document` and `get_structure` always return **every** top-level body paragraph, with
no way to ask for a subset. This is fine for the fixture-sized documents this project has
been tested against so far, but it does not scale to the primary real-world use case named in
[README.md](../../README.md) — an autonomous agent drafting a long document (a report, a
thesis) incrementally, checking the current state before every edit. As that document grows,
every such check re-transmits the *entire* document text or structure, even though the agent
usually only needs the last few paragraphs (to decide where to append next) or a small,
already-known range (to verify a recent edit). The cost of a check grows without bound with
the document, which is exactly backwards for an operation meant to run before every write.

This is not a hypothetical concern: `research-graphrag` (a sibling MCP server commonly paired
with this one for literature-grounded writing) already treats unbounded response growth as a
concrete problem it hit in practice — its own ADR 0037 ("Obergrenze für MCP-Tool-Antworten")
caps every retrieval tool's result count after a real response exceeded half the MCP transport
limit. Anthropic's own guidance for designing agent tools makes the same point in general terms:
implement "pagination, filtering, or truncation for any data that could be large, ensuring
smaller default responses" ("Writing effective tools for AI agents", Anthropic engineering
blog, 2026). `docx-mcp` has no such mechanism on its two read tools today.

Doing nothing is not an option once a document exceeds a few dozen paragraphs; the question is
how narrowly to scope the fix.

## Decision

Add two optional, keyword-only parameters — `start_paragraph: int | None = None` and
`end_paragraph: int | None = None` — to both `read_document` and `get_structure`. They form a
half-open range `[start_paragraph, end_paragraph)`, addressed by the **same** `paragraph_index`
convention every other tool in this project already uses (`find_text`, `replace_text`,
`insert_paragraph`, `delete_paragraph`).

1. **Default (`None`/`None`) is today's behavior, unchanged.** Every existing call, and every
   existing test, keeps working without modification — this is a purely additive extension
   (`MINOR` version bump per
   [docs/documentation-standards.md §4.1](../documentation-standards.md#41-tool-versioning)),
   not a breaking one.
2. **Range resolution is Python-slice-like, not index-lookup-like.** A negative
   `start_paragraph`, or an `end_paragraph` less than the (given or defaulted) `start_paragraph`,
   is rejected as a caller error (`ParagraphRangeError` → `ToolError`) — that can never reflect a
   legitimate range. But a bound *beyond* the document's actual paragraph count is silently
   **clamped**, not rejected. This is deliberate: the dominant use case is a caller checking a
   *growing* document without already knowing its current exact length (`start_paragraph=200,
   end_paragraph=250` against a 210-paragraph document should return paragraphs 200–209, not
   fail) — forcing an exact upper bound would reintroduce the same "read everything first to
   find out how much there is" cost this ADR removes. This mirrors `insert_paragraph`'s existing
   asymmetry between a single out-of-range index (rejected — a caller addressing one specific,
   named paragraph that doesn't exist is almost always a bug) and a range whose *edges* may
   legitimately be approximate.
3. **`read_document`'s output schema is unchanged.** `extract_text` (the underlying function in
   `docx_mcp.document`) gains the two new keyword-only parameters but keeps returning a plain
   `str` — the sliced text, rendered exactly as today. This keeps every one of `extract_text`'s
   ~30 existing call sites in `tests/test_document.py` and four `tests/integration/*.py` files
   working unmodified; only `server.py`'s thin wrapper gains the two new tool parameters.
4. **`get_structure`'s `paragraphs`/`toc`/`footnotes` are scoped to the range; `tables` is not.**
   `toc` and `footnotes` entries are addressed by `paragraph_index`, so filtering them to the
   same range keeps the response internally consistent and proportional to the requested slice —
   returning a `paragraphs` slice of 10 entries alongside a `toc` built from all 500 headings in
   the whole document would defeat the purpose. `tables` is addressed by an independent
   `table_index` sequence with no recorded relationship to `paragraph_index` (see
   [specs/get_structure.md §5](../../src/docx_mcp/specs/get_structure.md#5-limitations-non-goals),
   "`table_index` and `paragraph_index` are independent sequences") — inventing that relationship
   just to scope `tables` is a separate, bigger change with no demonstrated need yet, so `tables`
   continues to return every table, unscoped, documented as a limitation.
5. **`get_structure` gains one new output field, `total_paragraphs: int`** — the document's true
   paragraph count, independent of the requested range. Without it, a caller that requests a
   range has no way to tell whether it has reached the end of the document short of requesting
   an adjacent range and finding it empty. `read_document` gains no equivalent field: its output
   schema is deliberately left untouched (point 3), and a caller that needs the total count can
   get it from `get_structure` — including cheaply, via an empty probe range
   (`start_paragraph=0, end_paragraph=0`) that still returns `total_paragraphs` without paying for
   any paragraph text.
6. **Internal computation is not optimized for the range** — `get_document_structure` still
   builds the full `paragraphs`/`toc` lists exactly as before, then slices/filters the result. The
   actual problem this ADR solves is **response size** (the token cost the calling agent pays),
   not server-side CPU time; a `.docx`'s XML is small enough that parsing it fully is not the
   bottleneck. Skipping per-paragraph work for out-of-range paragraphs would be a legitimate
   future optimization if profiling ever shows it matters, but is not needed to solve the
   problem this ADR is scoped to.

## Alternatives

- **A dedicated new tool (e.g. `get_paragraphs`) instead of extending the existing two.**
  Rejected: it would overlap heavily with `get_structure` (same paragraph population, same
  addressing scheme) while adding a third tool surface for a caller to learn — against
  [CONTRIBUTING.md §1](../../CONTRIBUTING.md#1-guiding-principles) principle 6 ("a tool a caller
  can use correctly from its description alone beats a large tool that does many things
  ambiguously" cuts both ways: it also argues against near-duplicate tools with overlapping
  scope). Extending the two tools whose job this already is, with a backward-compatible optional
  parameter, is the smaller change.
- **Changing `extract_text`'s return type to a small dataclass carrying `total_paragraphs`
  alongside the text**, so `read_document`'s output could expose the same pagination metadata
  `get_structure` now does. Rejected for this phase: it would touch every one of `extract_text`'s
  ~30 existing call sites across `tests/test_document.py` and four integration test files for a
  benefit already available via `get_structure`'s new field — a materially larger, riskier
  change for a caller-facing gain that already has a path (point 5). Revisit only if real usage
  shows callers need it from `read_document` directly.
- **Hard-rejecting any `end_paragraph` beyond the actual paragraph count** (index-lookup-like,
  matching `insert_paragraph`/`delete_paragraph`'s single-index strictness). Rejected: it would
  force a caller to already know the document's exact length before requesting a tail range,
  defeating the purpose for the dominant "check the end of a growing document" use case (point
  2).
- **Also scoping `tables` to the range.** Rejected for this phase: `table_index` and
  `paragraph_index` are independent, unrelated sequences today (point 4); recovering a
  paragraph-relative position for each table is a separate, larger change with no demonstrated
  need yet.

## Consequences

- **Positive:** The dominant "check a growing document before the next edit" workflow no longer
  costs the calling agent tokens proportional to the whole document; `get_structure` gives a
  caller an explicit, cheap way to learn the document's total length independent of the range it
  actually wants to read. Both changes are purely additive — no existing caller, test, or schema
  consumer needs to change.
- **Negative / Effort:** Two tools now carry range-validation logic (`resolve_paragraph_range`,
  shared from `docx_mcp.document` to avoid duplicating it in `docx_mcp.structure`); a new error
  case (`ParagraphRangeError`) is added to both tools' error tables. `get_structure`'s output
  schema grows by one field, a `MINOR` bump per
  [docs/documentation-standards.md §4.1](../documentation-standards.md#41-tool-versioning)
  (`1.0.0` → `1.1.0` for both tools, since `read_document`'s *input* schema also grows even
  though its output does not).
- **Follow-up decisions:** A batching improvement for the *write* side (letting `insert_paragraph`
  attach one or more footnotes in the same call, to cut the round-trip count when drafting a
  citation-heavy paragraph) was considered during the same review that produced this ADR but is
  **deliberately deferred** — no current evidence shows the write-side round-trip count is
  actually a bottleneck the way the read-side response size demonstrably is, and
  `research-graphrag`'s own project culture (see its Roadmap, "Erst messen") is to measure before
  adding an optimization, not to add one speculatively. It is recorded here, explicitly, as a
  candidate for a future ADR if real usage shows it matters — not silently dropped.
