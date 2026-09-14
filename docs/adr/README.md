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
| — | *(none yet — the first ADR is expected in Phase 1, deciding the OOXML library)* | — |

---

## 5. Open/planned ADRs

- Phase 1 requires an ADR deciding the OOXML library (`lxml` vs. `xml.etree`) and, tied to it, the dependency/lockfile strategy (see [CONTRIBUTING.md §3](../../CONTRIBUTING.md#3-technical-stack)).
