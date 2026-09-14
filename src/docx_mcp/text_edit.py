"""`find_text`/`replace_text`: literal text search and run-splitting replacement.

Owns both tools' logic (a deliberate, phase-scoped exception to one-module-
per-tool - see [ADR-0005](../../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md)):
they share one `TextLocation` addressing shape and one offset-computation
rule, and `replace_text`'s core correctness requirement - re-verifying a
caller-supplied location before writing - *is* "recompute what `find_text`
would compute at those offsets, and compare."

A location is always `(paragraph_index, start_offset, end_offset)`, a
half-open slice of `document.py`'s `paragraph_plain_text(paragraph)` - the
same marker-free text `get_structure` exposes, never `read_document`'s
`"# "`/`"[^N]"`-annotated rendering. Matching is literal-substring only, no
regex (see [specs/find_text.md §5](specs/find_text.md#5-limitations-non-goals)).

The replace algorithm (see
[ADR-0005](../../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md)
for the full rationale) walks a paragraph's `w:t`/`w:tab` descendants in the
same document order `document.py`'s internal traversal uses, maps the
matched character span onto them, truncates boundary runs' text in place
(keeping their original, untouched `w:rPr`), and inserts one new run - a
deep copy of the *leftmost* matched run's `w:rPr` - to hold the replacement
text. A match whose runs sit under different nearest special-container
ancestors (`w:hyperlink`/`w:ins`/`w:del`/`w:smartTag`) is rejected rather
than silently spliced across the boundary. A global replace (no `location`)
computes every match against the original, untouched tree first, then
applies them per paragraph in descending offset order so earlier offsets
never drift.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from docx_mcp.document import (
    DOCUMENT_PART,
    NSMAP,
    WORD_NS,
    get_body,
    paragraph_plain_text,
)
from docx_mcp.ooxml import (
    MAX_DOCX_SIZE_BYTES,
    atomic_write_part,
    parse_xml,
    read_part,
    validate_and_open,
)

__all__ = [
    "TextEditError",
    "TextLocation",
    "TextMatch",
    "find_text_matches",
    "replace_text_in_document",
]

_W_P = f"{{{WORD_NS}}}p"
_W_T = f"{{{WORD_NS}}}t"
_W_TAB = f"{{{WORD_NS}}}tab"
_W_RPR = f"{{{WORD_NS}}}rPr"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class TextEditError(Exception):
    """Raised for a `find_text`/`replace_text`-specific error.

    See [specs/find_text.md §7](specs/find_text.md#7-error-behavior) and
    [specs/replace_text.md §7](specs/replace_text.md#7-error-behavior) for
    the full error table.
    """


@dataclass(frozen=True)
class TextLocation:
    """A half-open character span within one top-level body paragraph's
    `paragraph_plain_text` (see specs/find_text.md §3, `location`)."""

    paragraph_index: int
    start_offset: int
    end_offset: int


@dataclass(frozen=True)
class TextMatch:
    """One `find_text` match: its location, the matched text, and the
    matched paragraph's full text (see specs/find_text.md §3)."""

    location: TextLocation
    matched_text: str
    paragraph_text: str


def _find_spans(text: str, search_text: str, *, case_sensitive: bool) -> list[tuple[int, int]]:
    """Every non-overlapping occurrence of `search_text` in `text`, left to
    right, as half-open `(start, end)` spans.

    Case-insensitive comparison compares same-length windows of the
    *original* text against `search_text`, never a whole-string `.lower()`
    pass first - some Unicode case folding changes length (e.g. `'İ'.lower()`
    has two characters), which would desynchronize offsets from `text` as
    given. `matched_text` (built by the caller from these spans) therefore
    always carries the text's actual, original casing, never `search_text`'s.
    Assumes `search_text` is non-empty - callers reject an empty
    `search_text` before this is ever called.
    """
    spans: list[tuple[int, int]] = []
    needle_length = len(search_text)
    comparison_needle = search_text if case_sensitive else search_text.lower()
    limit = len(text) - needle_length + 1
    index = 0
    while index < limit:
        window = text[index : index + needle_length]
        candidate = window if case_sensitive else window.lower()
        if candidate == comparison_needle:
            spans.append((index, index + needle_length))
            index += needle_length
        else:
            index += 1
    return spans


def _iter_offset_nodes(
    paragraph: etree._Element,
) -> list[tuple[etree._Element, bool, int, int]]:
    """Every `w:t`/`w:tab` descendant of `paragraph`, in the same document
    order `document.py`'s internal paragraph-text traversal uses, as
    `(node, is_tab, start, end)` with cumulative character offsets matching
    `paragraph_plain_text`'s output exactly - a `w:t` contributes
    `len(text)` characters, a `w:tab` contributes exactly one, and (as in
    `document.py`) any other node, including `w:footnoteReference`,
    contributes zero and is therefore never returned here."""
    nodes: list[tuple[etree._Element, bool, int, int]] = []
    cursor = 0
    for node in paragraph.iter():
        if node.tag == _W_T:
            length = len(node.text or "")
            nodes.append((node, False, cursor, cursor + length))
            cursor += length
        elif node.tag == _W_TAB:
            nodes.append((node, True, cursor, cursor + 1))
            cursor += 1
    return nodes


def _parent_element(node: etree._Element) -> etree._Element:
    """`node.getparent()`, asserted non-`None`.

    Every `w:t`/`w:tab`/`w:r` node reached by walking a real, parsed
    paragraph always has a parent element (at minimum the paragraph itself);
    this narrows lxml's `_Element | None` return type accordingly rather
    than threading an unreachable `None` case through every call site.
    """
    parent = node.getparent()
    assert parent is not None
    return parent


def _container_of(run: etree._Element) -> etree._Element | None:
    """The run's nearest special-container ancestor (`w:hyperlink`, `w:ins`,
    `w:del`, `w:smartTag`, or any other non-`w:p` immediate parent), or
    `None` if the run is a direct child of `w:p`. Used to detect a match
    that would splice across such a boundary (see module docstring,
    [ADR-0005](../../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md))."""
    parent = run.getparent()
    if parent is None or parent.tag == _W_P:
        return None
    return parent


def _apply_replacement(
    paragraph: etree._Element, start: int, end: int, replacement_text: str
) -> None:
    """Replace `paragraph_plain_text(paragraph)[start:end]` with
    `replacement_text`, in place, per the algorithm in the module docstring.

    Raises:
        TextEditError: If the matched span's runs sit under different
            nearest special-container ancestors.
    """
    involved = [
        (node, is_tab, s, e)
        for node, is_tab, s, e in _iter_offset_nodes(paragraph)
        if e > start and s < end
    ]
    if not involved:
        return

    runs = [_parent_element(node) for node, _is_tab, _s, _e in involved]
    containers = [_container_of(run) for run in runs]
    if any(container is not containers[0] for container in containers):
        raise TextEditError(
            "match spans a run-container boundary (e.g. a hyperlink or tracked-change "
            "block) - not supported"
        )

    leftmost_run = runs[0]
    leftmost_rpr = leftmost_run.find(_W_RPR)
    replacement_rpr = copy.deepcopy(leftmost_rpr) if leftmost_rpr is not None else None

    for node, is_tab, node_start, node_end in involved:
        run = _parent_element(node)
        if is_tab:
            run.remove(node)
            continue

        text = node.text or ""
        overlap_start = max(node_start, start) - node_start
        overlap_end = min(node_end, end) - node_start
        prefix = text[:overlap_start]
        suffix = text[overlap_end:]

        if prefix and suffix:
            # The match lies entirely within this one run's text, with
            # untouched text on both sides - the trivial single-run case.
            # One element can't hold two disjoint spans, so the suffix
            # becomes a new sibling run, a deep copy of the original run
            # (preserving its rPr and any other attributes) with its text
            # replaced.
            suffix_run = copy.deepcopy(run)
            suffix_t = suffix_run.find(_W_T)
            assert suffix_t is not None
            suffix_t.text = suffix
            suffix_t.set(_XML_SPACE, "preserve")
            run.addnext(suffix_run)
            node.text = prefix
            node.set(_XML_SPACE, "preserve")
        elif prefix:
            node.text = prefix
            node.set(_XML_SPACE, "preserve")
        elif suffix:
            node.text = suffix
            node.set(_XML_SPACE, "preserve")
        else:
            run.remove(node)

    if replacement_text:
        replacement_run = etree.Element(f"{{{WORD_NS}}}r")
        if replacement_rpr is not None:
            replacement_run.append(replacement_rpr)
        replacement_t = etree.SubElement(replacement_run, _W_T)
        replacement_t.text = replacement_text
        replacement_t.set(_XML_SPACE, "preserve")
        # The leftmost matched node keeps an untouched *prefix* in place only
        # when the match starts strictly after that node's own start; then
        # the replacement must follow it (`addnext`). Otherwise - the node is
        # either fully consumed or keeps only a *suffix* - the replacement
        # must precede it (`addprevious`), or it would land after surviving
        # suffix text instead of where the removed text used to be.
        leftmost_node_start = involved[0][2]
        if start > leftmost_node_start:
            leftmost_run.addnext(replacement_run)
        else:
            leftmost_run.addprevious(replacement_run)

    for run in dict.fromkeys(runs):  # de-duplicated, original relative order
        remaining = [child for child in run if child.tag != _W_RPR]
        if not remaining:
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)


def find_text_matches(
    docx_path: Path,
    search_text: str,
    *,
    case_sensitive: bool = True,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> tuple[TextMatch, ...]:
    """Find every literal, non-overlapping occurrence of `search_text` in a `.docx`'s body.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots; this function
            does not perform any sandboxing itself.
        search_text: The literal text to search for. Must be non-empty.
        case_sensitive: `True` for exact-case matching only; `False` for
            Unicode-aware case-insensitive matching (see `_find_spans`).
        max_size_bytes: Reject the file if it exceeds this, before it is opened.

    Returns:
        Every match, in document order, per
        [specs/find_text.md §3](specs/find_text.md#3-output-schema).

    Raises:
        TextEditError: If `search_text` is empty.
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, or is missing/has a
            malformed `word/document.xml`.
    """
    if not search_text:
        raise TextEditError("search_text must not be empty")

    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)

    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    body = get_body(root)

    matches: list[TextMatch] = []
    for paragraph_index, paragraph in enumerate(body.findall("w:p", namespaces=NSMAP)):
        text = paragraph_plain_text(paragraph)
        for start, end in _find_spans(text, search_text, case_sensitive=case_sensitive):
            matches.append(
                TextMatch(
                    location=TextLocation(paragraph_index, start, end),
                    matched_text=text[start:end],
                    paragraph_text=text,
                )
            )
    return tuple(matches)


def replace_text_in_document(
    docx_path: Path,
    search_text: str,
    replacement_text: str,
    *,
    location: TextLocation | None = None,
    case_sensitive: bool = True,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> int:
    """Replace `search_text` with `replacement_text`, at `location` or everywhere.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots; this function
            does not perform any sandboxing itself.
        search_text: The literal text to replace. Must be non-empty; in
            location mode, also the value re-verified against `location`
            before anything is written.
        replacement_text: The literal replacement text; may be empty (a
            pure deletion).
        location: `None` (default) replaces every occurrence in the
            document. Given, replaces exactly that one occurrence, after
            confirming the document still holds `search_text` there.
        case_sensitive: Matching/verification case sensitivity.
        max_size_bytes: Reject the file if it exceeds this, before it is opened.

    Returns:
        The number of occurrences replaced: always `1` in location mode,
        always `>= 1` in global mode (see
        [specs/replace_text.md §3](specs/replace_text.md#3-output-schema)).

    Raises:
        TextEditError: Per [specs/replace_text.md §7](specs/replace_text.md#7-error-behavior):
            empty `search_text`; `location` referencing a nonexistent
            paragraph or out-of-range/invalid offsets; a stale `location`;
            zero matches in global mode; a match spanning a run-container
            boundary.
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, or is missing/has a
            malformed `word/document.xml`.
    """
    if not search_text:
        raise TextEditError("search_text must not be empty")

    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)

    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    body = get_body(root)
    paragraphs = body.findall("w:p", namespaces=NSMAP)

    if location is not None:
        if location.paragraph_index < 0 or location.paragraph_index >= len(paragraphs):
            raise TextEditError("location references a paragraph that does not exist")

        paragraph = paragraphs[location.paragraph_index]
        text = paragraph_plain_text(paragraph)
        start, end = location.start_offset, location.end_offset
        if start < 0 or end > len(text) or start >= end:
            raise TextEditError("location is out of range")

        candidate = text[start:end]
        candidate_matches = (
            candidate == search_text if case_sensitive else candidate.lower() == search_text.lower()
        )
        if not candidate_matches:
            raise TextEditError(
                "location is stale: the document no longer contains search_text there "
                "- re-run find_text"
            )

        _apply_replacement(paragraph, start, end, replacement_text)
        replacements_made = 1
    else:
        spans_by_paragraph: dict[int, list[tuple[int, int]]] = {}
        replacements_made = 0
        for index, paragraph in enumerate(paragraphs):
            text = paragraph_plain_text(paragraph)
            spans = _find_spans(text, search_text, case_sensitive=case_sensitive)
            if spans:
                spans_by_paragraph[index] = spans
                replacements_made += len(spans)

        if replacements_made == 0:
            raise TextEditError("no occurrences of search_text found")

        for index, spans in spans_by_paragraph.items():
            paragraph = paragraphs[index]
            for start, end in sorted(spans, reverse=True):
                _apply_replacement(paragraph, start, end, replacement_text)

    new_document_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    atomic_write_part(docx_path, DOCUMENT_PART, new_document_xml, max_size_bytes=max_size_bytes)
    return replacements_made
