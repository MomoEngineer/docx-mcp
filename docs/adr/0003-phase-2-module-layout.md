# 0003 – Phase 2 Module Layout: Splitting `document.py`

- **Status:** Accepted
- **Date:** 2026-09-14

## Context

[ADR-0001](0001-ooxml-library-and-module-layout.md) deliberately kept all OOXML handling in a single `document.py` module for Phase 1's thin slice, and explicitly named the trigger for splitting it: "deferred until Phase 2 or 3 introduce a second read path that would otherwise duplicate logic against the same module."

Phase 2 ([Roadmap.md](../../Roadmap.md#phase-2--structural-read-get_structure-get_metadata)) adds exactly that: `get_structure` needs to open a `.docx`, resolve heading levels via `word/styles.xml` (`w:outlineLvl`, `w:basedOn` chains), render table cells, and build a footnote-anchor index; `get_metadata` needs to read `docProps/core.xml` and compute a word count. Both need the same zip-opening, size-checking, and hardened-XML-parsing logic `document.py`'s `extract_text` already implements (a hardened `lxml.etree.XMLParser`, `MAX_DOCX_SIZE_BYTES`, `InvalidDocumentError`, the zip/part-reading sequence). Without a split, that logic would be copy-pasted three times, and `structure.py`'s paragraph rendering (headings, tabs, footnote markers, for table-cell text) would duplicate `document.py`'s paragraph renderer — the exact duplication ADR-0001 named as the split trigger.

## Decision

Split into four modules with a one-directional dependency chain (`ooxml.py` → `document.py` → {`structure.py`, `metadata.py`}), no cycles:

```
src/docx_mcp/
├── ooxml.py       # generic OOXML/zip plumbing shared by every parsing module:
│                  #   hardened lxml parser, MAX_DOCX_SIZE_BYTES, InvalidDocumentError,
│                  #   "open a .docx part as a parsed element tree" helper.
│                  #   No WordprocessingML-specific rendering logic lives here.
├── document.py    # WordprocessingML paragraph rendering: extract_text (Phase 1,
│                  #   unchanged behavior) plus two exported helpers reused by
│                  #   structure.py/metadata.py - a marker-free plain-text
│                  #   renderer (paragraph_plain_text, used for paragraphs[]
│                  #   and word_count) and the marker-inclusive one
│                  #   (render_paragraph, reused for table-cell text, which
│                  #   per specs/get_structure.md §3 gets the same "# "/"[^N]"
│                  #   markers as read_document) - so "how a paragraph's text
│                  #   is assembled from its runs" has exactly one
│                  #   implementation each, not duplicated per module.
├── structure.py   # get_structure: heading-level resolution via styles.xml,
│                  #   paragraph index, table-as-text, footnote-anchor index.
├── metadata.py    # get_metadata: docProps/core.xml + word count (via
│                  #   document.py's plain-text helper, not extract_text's
│                  #   marker-annotated output, so heading '#'/footnote '[^N]'
│                  #   markers never inflate the count).
└── specs/
    ├── read_document.md
    ├── get_structure.md
    └── get_metadata.md
```

`document.py` keeps ownership of paragraph-level WordprocessingML rendering (the one place `w:t`/`w:tab`/`w:footnoteReference` traversal happens) rather than moving it into `ooxml.py`, which stays deliberately generic (it has no knowledge of `w:p`, `w:tbl`, or any WordprocessingML tag). `structure.py` and `metadata.py` both depend on `document.py` for paragraph text, never the reverse — `document.py`'s Phase 1 public API (`extract_text`, `InvalidDocumentError`) is unchanged, so no regression in `read_document`.

## Alternatives

- **Keep everything in `document.py`.** Rejected: this is precisely the duplication ADR-0001 flagged as the split trigger, now realized.
- **One `parsing.py` for all four tools' logic.** Rejected: collapses tool-specific logic (heading resolution, table rendering, metadata reading) into one large module with no clear ownership boundary, working against `docs/repository-structure.md` §2's "tool logic = one module" convention.
- **Fold the shared zip/parser plumbing into `document.py` itself instead of a new `ooxml.py`, and have `structure.py`/`metadata.py` import from `document.py`.** Rejected: would make `document.py` (Phase 1's `read_document` implementation) a dependency of unrelated tools' security-relevant parsing setup, muddying which module owns "safe to open a `.docx` part" versus "how to render a paragraph." A dedicated `ooxml.py` makes the hardened-parser contract ([ADR-0001](0001-ooxml-library-and-module-layout.md), [docs/security-model.md §4](../security-model.md#4-input-validation)) visibly shared infrastructure, not an implicit side effect of importing `document.py`.

## Consequences

- **Positive:** no duplicated hardened-parser/size-check/zip-opening logic across three tool modules; `document.py`'s paragraph-rendering logic has exactly one implementation, reused (not copied) by `structure.py` and `metadata.py`; each module's ownership matches one tool family, consistent with `docs/repository-structure.md` §2.
- **Negative / Effort:** one more module to navigate than a single `document.py`; `document.py` is refactored (its private helpers move to `ooxml.py`) even though its own public behavior and tests must stay green — verified by re-running the full Phase 1 test suite unchanged after the refactor.
- **Follow-up decisions:** `docs/repository-structure.md`'s binding folder-structure diagram is updated alongside this ADR to list `ooxml.py`, `structure.py`, and `metadata.py`. None of Phase 3's `get_footnotes` is decided here — it will need its own look at whether it extends `structure.py` (footnote anchors) or gets a fifth module, deferred to that phase's own ADR discussion if warranted.
