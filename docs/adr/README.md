# Architecture Decision Records (ADRs)

An **Architecture Decision Record (ADR)** documents a significant architectural or policy decision: the context, the decision made, and its consequences. ADRs make it traceable **why** the system is the way it is — central to this repository's traceability principle ([CONTRIBUTING.md §1](../../CONTRIBUTING.md#1-guiding-principles), principle 7).

---

## 1. When is an ADR required?

An ADR is created for, among others:

- Choosing the OOXML handling library (`lxml` vs. the standard-library `xml.etree`) — the first architectural decision, deliberately deferred from Phase 0 (see [Roadmap.md](../../Roadmap.md#phase-1--thin-vertical-slice-server-skeleton--read_document)).
- Any change to the shape of the OOXML handling, the tool surface, or the security model.
- Choosing or changing the dependency/lockfile strategy.
- Operating the server over a transport other than `stdio`.
- Deliberately accepted redundancy.
- Any decision that is hard to reverse or affects multiple modules.

---

## 2. Process

1. Copy [templates/adr-template.md](../../templates/adr-template.md).
2. Number sequentially: `NNNN-short-title.md`.
3. Set the status to `Proposed` while the decision is under discussion.
4. After acceptance: set the status to `Accepted`.
5. If a decision is later superseded: set the old ADR's status to `Superseded by NNNN` — do not delete it.

---

## 3. Status values

- **Proposed** — under discussion.
- **Accepted** — in effect.
- **Rejected** — discarded, kept for traceability.
- **Superseded** — replaced by a newer ADR (with a reference).

---

## 4. Index of ADRs

| No. | Title | Status |
| --- | --- | --- |
| [0001](0001-ooxml-library-and-module-layout.md) | OOXML Handling Library and Module Layout | Accepted |
| [0002](0002-dependency-and-lockfile-strategy.md) | Dependency and Lockfile Strategy | Accepted |
| [0003](0003-phase-2-module-layout.md) | Phase 2 Module Layout: Splitting `document.py` | Accepted |
| [0004](0004-phase-3-footnote-module-and-shared-anchor-resolution.md) | Phase 3 Module Layout: `footnotes.py` and Shared Footnote-Anchor Resolution | Accepted |
| [0005](0005-phase-4-text-edit-module-and-atomic-write.md) | Phase 4 Module Layout: `text_edit.py` and the Atomic-Write Contract | Accepted |
| [0006](0006-phase-5-paragraph-edit-module-layout.md) | Phase 5 Module Layout: `paragraph_edit.py` and Style-Inheritance Rules for `insert_paragraph`/`delete_paragraph` | Accepted |
| [0007](0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md) | Phase 6 Module Layout: `footnote_edit.py`, Multi-Part Atomic Writes, and Anchor/Content Rules for `add_footnote`/`edit_footnote` | Accepted |

---

## 5. Open/planned ADRs

- None currently open. The seven ADRs required so far (OOXML library, dependency/lockfile strategy, Phase 2 module layout, Phase 3 module layout, Phase 4 module layout and atomic-write contract, Phase 5 module layout and style-inheritance rules, Phase 6 module layout and multi-part atomic writes) are decided above. [ADR-0007](0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md)'s follow-up note asked Phase 7 (security hardening) to extend its path-traversal/atomic-write test sweep to cover `ooxml.atomic_write_parts` alongside `atomic_write_part` — verified during Phase 7 as already satisfied by Phase 6's own test suite (`tests/test_ooxml_atomic_write.py`'s `atomic_write_parts` simulated-crash tests; `tests/test_add_footnote_tool.py`'s path-traversal-rejection test, `add_footnote` being the sole caller of `atomic_write_parts`), so no new ADR was needed for it. None of Phase 7's other findings (a missing "Examples" section across every tool spec/docstring, closed by adding one to each plus a permanent `tests/test_spec_examples.py` regression check; an inconsistent MCP-protocol-level malformed-input test present for every read tool but no write tool, closed by adding the missing five) changed the shape of the OOXML handling, the tool surface, or the security model either — see [docs/documentation-standards.md §3](../documentation-standards.md#3-architecture-decision-records-adrs) for the ADR-required categories — so Phase 7 added no new ADR.
