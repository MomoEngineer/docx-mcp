"""Tests for `docx_mcp.text_edit`, from `specs/find_text.md`/`specs/replace_text.md`.

Covers the functional, formatting-fidelity, and error/edge test categories
required by [docs/testing.md §2](../docs/testing.md#2-test-types-per-mcp-tool)
and [docs/security-model.md §5](../docs/security-model.md#5-testing-obligations),
following the same synthetic-XML-fixture convention `tests/test_footnotes.py`/
`tests/test_structure.py` established for the error/edge matrix, plus the
real `tests/fixtures/run_split.docx` fixture (Phase 4) for the run-splitting/
tab/multi-match functional happy paths.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

import docx_mcp.ooxml as ooxml_module
from docx_mcp.document import (
    DOCUMENT_PART,
    NSMAP,
    WORD_NS,
    InvalidDocumentError,
    get_body,
    paragraph_plain_text,
)
from docx_mcp.footnotes import get_document_footnotes
from docx_mcp.ooxml import parse_xml, read_part, validate_and_open
from docx_mcp.text_edit import (
    TextEditError,
    TextLocation,
    find_text_matches,
    replace_text_in_document,
)

_W_B = f"{{{WORD_NS}}}b"
_W_I = f"{{{WORD_NS}}}i"


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_xml(body_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{body_inner_xml}</w:body></w:document>'
    ).encode()


def _paragraphs(docx_path: Path) -> list[etree._Element]:
    with validate_and_open(docx_path) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)
    root = parse_xml(raw_xml, part_name=DOCUMENT_PART)
    body = get_body(root)
    return body.findall("w:p", namespaces=NSMAP)


def _paragraph_texts(docx_path: Path) -> list[str]:
    return [paragraph_plain_text(p) for p in _paragraphs(docx_path)]


def _runs(paragraph: etree._Element) -> list[etree._Element]:
    return paragraph.findall("w:r", namespaces=NSMAP)


# --- Functional: find_text_matches (against run_split.docx) --------------------


def test_find_split_run_match(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "receive")

    assert len(matches) == 1
    assert matches[0].location == TextLocation(0, 7, 14)
    assert matches[0].matched_text == "receive"
    assert matches[0].paragraph_text == "Please receive this message."


def test_find_match_entirely_within_one_run(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "simple")

    assert len(matches) == 1
    assert matches[0].location == TextLocation(1, 10, 16)


def test_find_match_spanning_a_tab(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "Left\tRight")

    assert len(matches) == 1
    assert matches[0].location == TextLocation(2, 0, 10)


def test_find_case_insensitive_preserves_original_casing(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "RECEIVE", case_sensitive=False)

    assert len(matches) == 1
    assert matches[0].matched_text == "receive"


def test_find_case_sensitive_default_finds_nothing_for_different_casing(
    run_split_docx: Path,
) -> None:
    matches = find_text_matches(run_split_docx, "RECEIVE")

    assert matches == ()


def test_find_three_non_overlapping_matches_in_one_paragraph(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "test")

    same_paragraph = [m for m in matches if m.location.paragraph_index == 3]
    assert [m.location for m in same_paragraph] == [
        TextLocation(3, 24, 28),
        TextLocation(3, 30, 34),
        TextLocation(3, 40, 44),
    ]


def test_find_matches_span_multiple_paragraphs(run_split_docx: Path) -> None:
    matches = find_text_matches(run_split_docx, "test")

    assert len(matches) == 4
    assert {m.location.paragraph_index for m in matches} == {3, 4}


def test_find_no_match_returns_empty_tuple(run_split_docx: Path) -> None:
    assert find_text_matches(run_split_docx, "not present anywhere") == ()


def test_find_empty_search_text_is_rejected(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="empty"):
        find_text_matches(run_split_docx, "")


def test_find_matches_immediately_adjacent_to_a_footnote_reference_on_both_sides(
    tmp_path: Path,
) -> None:
    """A footnote reference contributes zero characters to
    `paragraph_plain_text` (module docstring, `_iter_offset_nodes`), so text
    ending right before it and text starting right after it are ordinary,
    correctly-offset matches, and the anchor itself never appears (as a
    `"[^N]"` marker or otherwise) in `paragraph_text` - `find_text` operates
    on the same marker-free text as `get_structure`, never `read_document`'s
    rendering (specs/find_text.md §5)."""
    path = tmp_path / "footnote_adjacent.docx"
    body = (
        "<w:p><w:r><w:t>See the important</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r>'
        "<w:r><w:t> note right after.</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    before_matches = find_text_matches(path, "important")
    after_matches = find_text_matches(path, "note right")

    assert len(before_matches) == 1
    assert before_matches[0].location == TextLocation(0, 8, 17)
    assert "[^" not in before_matches[0].paragraph_text
    assert len(after_matches) == 1
    assert after_matches[0].location == TextLocation(0, 18, 28)


# --- Functional: replace_text_in_document, location mode (run_split.docx) ------


def test_replace_split_run_preserves_boundary_run_formatting(run_split_docx: Path) -> None:
    """ "receive" spans a bold run ("rec") and a not-bold run ("eive"). The
    replacement must inherit the *leftmost* matched run's formatting (bold),
    and - since the whole match consumes both runs entirely here - no
    boundary-run text survives to check separately; the untouched runs
    before/after the match (spec §6) are what this test's sibling
    (`test_replace_partial_boundary_...`) exercises."""
    count = replace_text_in_document(
        run_split_docx, "receive", "REPLIED", location=TextLocation(0, 7, 14)
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[0] == "Please REPLIED this message."

    paragraph = _paragraphs(run_split_docx)[0]
    runs = _runs(paragraph)
    # "Please " (untouched) + "REPLIED" (new, bold - leftmost matched run's
    # rPr) + " this message." (untouched)
    texts = [r.findtext("w:t", namespaces=NSMAP) for r in runs]
    assert texts == ["Please ", "REPLIED", " this message."]
    replacement_run = runs[1]
    assert replacement_run.find("w:rPr", namespaces=NSMAP).find(_W_B) is not None


def test_replace_partial_boundary_keeps_original_untouched_formatting(
    run_split_docx: Path,
) -> None:
    """Replacing only "ceive" (a right-partial overlap of the bold "rec" run
    plus all of the not-bold "eive" run) must leave "re" behind in the
    original bold run, untouched."""
    count = replace_text_in_document(run_split_docx, "ceive", "X", location=TextLocation(0, 9, 14))

    assert count == 1
    assert _paragraph_texts(run_split_docx)[0] == "Please reX this message."

    paragraph = _paragraphs(run_split_docx)[0]
    runs = _runs(paragraph)
    texts = [r.findtext("w:t", namespaces=NSMAP) for r in runs]
    assert texts == ["Please ", "re", "X", " this message."]
    prefix_run = runs[1]
    assert prefix_run.find("w:rPr", namespaces=NSMAP).find(_W_B) is not None


def test_replace_trivial_single_run_match_copies_formatting(run_split_docx: Path) -> None:
    """ "simple" is fully inside one italic run, with untouched text on both
    sides - the trivial case, verifying the replacement copies the run's
    rPr (italic), not just "didn't crash", and that the surviving prefix and
    suffix keep the same italic formatting."""
    count = replace_text_in_document(
        run_split_docx, "simple", "basic", location=TextLocation(1, 10, 16)
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[1] == "This is a basic example sentence."

    paragraph = _paragraphs(run_split_docx)[1]
    runs = _runs(paragraph)
    texts = [r.findtext("w:t", namespaces=NSMAP) for r in runs]
    assert texts == ["This is a ", "basic", " example sentence."]
    for run in runs:
        assert run.find("w:rPr", namespaces=NSMAP).find(_W_I) is not None


def test_replace_spanning_a_tab_removes_the_tab_element(run_split_docx: Path) -> None:
    count = replace_text_in_document(
        run_split_docx, "Left\tRight", "Middle", location=TextLocation(2, 0, 10)
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[2] == "Middle"

    paragraph = _paragraphs(run_split_docx)[2]
    assert paragraph.find(f".//{{{WORD_NS}}}tab") is None


def test_replace_with_longer_replacement_text(run_split_docx: Path) -> None:
    count = replace_text_in_document(
        run_split_docx, "simple", "considerably longer", location=TextLocation(1, 10, 16)
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[1] == "This is a considerably longer example sentence."


def test_replace_pure_deletion_leaves_no_stray_empty_run(run_split_docx: Path) -> None:
    count = replace_text_in_document(
        run_split_docx, "simple ", "", location=TextLocation(1, 10, 17)
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[1] == "This is a example sentence."

    paragraph = _paragraphs(run_split_docx)[1]
    for run in _runs(paragraph):
        has_text = run.find("w:t", namespaces=NSMAP) is not None
        has_tab = run.find("w:tab", namespaces=NSMAP) is not None
        assert has_text or has_tab


# --- Functional: replace_text_in_document, global mode (run_split.docx) -------


def test_replace_global_mode_replaces_all_occurrences_across_paragraphs(
    run_split_docx: Path,
) -> None:
    count = replace_text_in_document(run_split_docx, "test", "exam")

    assert count == 4
    texts = _paragraph_texts(run_split_docx)
    assert texts[3] == "This paragraph mentions exam, exam, and exam again."
    assert texts[4] == "One more exam paragraph for global replace."


def test_replace_global_mode_same_paragraph_multi_match_no_offset_drift(
    run_split_docx: Path,
) -> None:
    """Reverse-order-per-paragraph application (ADR-0005) must not corrupt
    offsets when replacement_text has a different length than search_text."""
    count = replace_text_in_document(run_split_docx, "test", "a longer word")

    assert count == 4
    assert _paragraph_texts(run_split_docx)[3] == (
        "This paragraph mentions a longer word, a longer word, and a longer word again."
    )


def test_replace_global_mode_zero_matches_is_an_error(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="no occurrences"):
        replace_text_in_document(run_split_docx, "not present anywhere", "x")


# --- Formatting-fidelity ---------------------------------------------------------


def test_replace_does_not_alter_untouched_paragraphs_or_zip_parts(run_split_docx: Path) -> None:
    before_paragraphs = _paragraphs(run_split_docx)
    untouched_before = {i: etree.tostring(p) for i, p in enumerate(before_paragraphs) if i != 1}
    with zipfile.ZipFile(run_split_docx) as archive:
        other_parts_before = {
            name: archive.read(name) for name in archive.namelist() if name != DOCUMENT_PART
        }

    replace_text_in_document(run_split_docx, "simple", "basic", location=TextLocation(1, 10, 16))

    after_paragraphs = _paragraphs(run_split_docx)
    for i, xml_before in untouched_before.items():
        assert etree.tostring(after_paragraphs[i]) == xml_before

    with zipfile.ZipFile(run_split_docx) as archive:
        other_parts_after = {
            name: archive.read(name) for name in archive.namelist() if name != DOCUMENT_PART
        }
    assert other_parts_after == other_parts_before


# --- Footnote survival (minimal.docx) -------------------------------------------


def test_replace_next_to_a_footnote_reference_leaves_it_untouched(minimal_docx: Path) -> None:
    """`minimal.docx`'s paragraph 2 is "This paragraph has a footnote
    reference." with footnote id 1 anchored immediately after "reference.".
    Editing text ending right before the anchor must not disturb it."""
    before = get_document_footnotes(minimal_docx)
    assert [(f.id, f.paragraph_index) for f in before] == [("1", 2)]

    replace_text_in_document(minimal_docx, "footnote reference", "annotation marker")

    assert _paragraph_texts(minimal_docx)[2] == "This paragraph has a annotation marker."
    after = get_document_footnotes(minimal_docx)
    assert [(f.id, f.paragraph_index, f.content) for f in after] == [
        ("1", 2, " This is a footnote."),
    ]


def test_replace_text_starting_immediately_after_a_footnote_reference_leaves_it_untouched(
    tmp_path: Path,
) -> None:
    """The "after" counterpart to `test_replace_next_to_a_footnote_reference_...`
    above: `minimal.docx`'s only anchor sits at the very end of its
    paragraph, with no trailing text, so a synthetic document (matching
    `test_find_matches_immediately_adjacent_to_a_footnote_reference_...`'s
    structure) is needed to exercise editing text that starts right after
    an anchor."""
    path = tmp_path / "footnote_after.docx"
    body = (
        "<w:p><w:r><w:t>See the important</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r>'
        "<w:r><w:t> note right after.</w:t></w:r></w:p>"
    )
    footnotes = '<w:footnote w:id="1"><w:p><w:r><w:t>A note.</w:t></w:r></w:p></w:footnote>'
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "word/footnotes.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:footnotes xmlns:w="{WORD_NS}">{footnotes}</w:footnotes>'
            ).encode(),
        },
    )

    replace_text_in_document(path, "note right", "correction")

    assert _paragraph_texts(path) == ["See the important correction after."]
    after = get_document_footnotes(path)
    assert [(f.id, f.paragraph_index, f.content) for f in after] == [("1", 0, "A note.")]


# --- Error / edge ----------------------------------------------------------------


def test_replace_empty_search_text_is_rejected(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="empty"):
        replace_text_in_document(run_split_docx, "", "x")


def test_replace_location_paragraph_index_out_of_range(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="paragraph"):
        replace_text_in_document(run_split_docx, "receive", "x", location=TextLocation(999, 0, 1))


def test_replace_location_negative_paragraph_index(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="paragraph"):
        replace_text_in_document(run_split_docx, "receive", "x", location=TextLocation(-1, 0, 1))


def test_replace_location_offsets_out_of_range(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="out of range"):
        replace_text_in_document(run_split_docx, "receive", "x", location=TextLocation(0, 0, 9999))


def test_replace_location_start_not_before_end(run_split_docx: Path) -> None:
    with pytest.raises(TextEditError, match="out of range"):
        replace_text_in_document(run_split_docx, "receive", "x", location=TextLocation(0, 10, 5))


def test_replace_stale_location_is_rejected(run_split_docx: Path) -> None:
    """A location correctly obtained via find_text, then invalidated by an
    intervening edit to the same paragraph, must be rejected when reused -
    the core race/hallucination-closing check (ADR-0005, specs/replace_text.md §7)."""
    location = find_text_matches(run_split_docx, "receive")[0].location

    replace_text_in_document(run_split_docx, "receive", "REPLIED", location=location)

    with pytest.raises(TextEditError, match="stale"):
        replace_text_in_document(run_split_docx, "receive", "AGAIN", location=location)


def test_replace_match_spanning_a_hyperlink_boundary_is_rejected(tmp_path: Path) -> None:
    """A match starting inside a `w:hyperlink` and continuing into a plain
    run outside it must be rejected, not silently spliced across the
    boundary (ADR-0005)."""
    path = tmp_path / "hyperlink_boundary.docx"
    body = (
        "<w:p>"
        "<w:hyperlink><w:r><w:t>click he</w:t></w:r></w:hyperlink>"
        "<w:r><w:t>re now</w:t></w:r>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(TextEditError, match="run-container boundary"):
        replace_text_in_document(path, "here", "there", location=TextLocation(0, 6, 10))


def test_replace_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"
    with pytest.raises(InvalidDocumentError):
        replace_text_in_document(missing, "a", "b")


def test_find_xxe_entity_is_never_resolved(tmp_path: Path) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe.docx"
    malicious_xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:document [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>'
        "<w:p><w:r><w:t>Safe text.</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>&xxe;</w:t></w:r></w:p>"
        "</w:body></w:document>"
    ).encode()
    _write_zip(path, {"word/document.xml": malicious_xml})

    matches = find_text_matches(path, "Safe")

    assert len(matches) == 1
    assert "TOP-SECRET-CONTENT" not in matches[0].paragraph_text


def test_replace_xxe_entity_is_never_resolved_and_edit_still_succeeds(tmp_path: Path) -> None:
    """The hardened parser never resolves the entity while `replace_text`
    reads and edits the document. Note: `etree.tostring` on the parsed
    element (not a full `ElementTree`) does not re-emit the original
    `DOCTYPE`/internal subset when serializing the edit - a deliberate,
    documented consequence of the write path (specs/replace_text.md §6): a
    `.docx` with such a declaration is not standard Word output, only ever
    adversarial, so the rewritten `word/document.xml` legitimately no longer
    declares (or needs) the now-unused entity. This test therefore checks
    the raw written bytes directly, rather than re-parsing the whole file
    with the same strict parser afterward, which would correctly reject the
    now-orphaned `&xxe;` reference still literally present, unresolved, in
    the untouched second paragraph."""
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe.docx"
    malicious_xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:document [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>'
        "<w:p><w:r><w:t>Safe text.</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>&xxe;</w:t></w:r></w:p>"
        "</w:body></w:document>"
    ).encode()
    _write_zip(path, {"word/document.xml": malicious_xml})

    count = replace_text_in_document(path, "Safe", "Edited")

    assert count == 1
    with zipfile.ZipFile(path) as archive:
        raw = archive.read(DOCUMENT_PART)
    # "Edited" and " text." land in separate <w:t> elements (the replacement
    # run and the original run's surviving suffix), so check each piece
    # rather than asserting they're byte-contiguous in the raw XML.
    assert b"Edited" in raw
    assert b" text." in raw
    assert b"TOP-SECRET-CONTENT" not in raw


# --- Error / edge: shared validation (mirrors tests/test_document.py) ----------


@pytest.mark.parametrize("call", ["find", "replace"])
def test_oversized_file_is_rejected_before_parsing(tmp_path: Path, call: str) -> None:
    path = tmp_path / "oversized.docx"
    _write_zip(
        path, {"word/document.xml": _document_xml("<w:p><w:r><w:t>Small.</w:t></w:r></w:p>")}
    )
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        if call == "find":
            find_text_matches(path, "Small", max_size_bytes=actual_size - 1)
        else:
            replace_text_in_document(path, "Small", "Tiny", max_size_bytes=actual_size - 1)


def test_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    _write_zip(
        path, {"word/document.xml": _document_xml("<w:p><w:r><w:t>Small.</w:t></w:r></w:p>")}
    )
    actual_size = path.stat().st_size

    matches = find_text_matches(path, "Small", max_size_bytes=actual_size)
    assert len(matches) == 1


def test_not_a_valid_zip_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not_a_zip.docx"
    path.write_bytes(b"this is not a zip file")

    with pytest.raises(InvalidDocumentError):
        find_text_matches(path, "x")
    with pytest.raises(InvalidDocumentError):
        replace_text_in_document(path, "x", "y")


def test_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_document_xml.docx"
    _write_zip(path, {"word/styles.xml": b"<styles/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        find_text_matches(path, "x")


def test_malformed_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed.docx"
    _write_zip(path, {"word/document.xml": b"<w:document><unclosed>"})

    with pytest.raises(InvalidDocumentError):
        find_text_matches(path, "x")


def test_missing_body_element_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_body.docx"
    xml = f'<w:document xmlns:w="{WORD_NS}"></w:document>'.encode()
    _write_zip(path, {"word/document.xml": xml})

    with pytest.raises(InvalidDocumentError, match="body"):
        find_text_matches(path, "x")


# --- Self-referential global replace (ADR-0005) ---------------------------------


def test_global_replace_terminates_correctly_when_replacement_contains_search_text(
    tmp_path: Path,
) -> None:
    """Replacing "cat" -> "cats" must replace exactly the 3 occurrences that
    existed in the original text, never re-matching freshly-inserted "cat"
    substrings inside "cats" and looping or over-replacing."""
    path = tmp_path / "self_referential.docx"
    body = "<w:p><w:r><w:t>cat cat cat</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    count = replace_text_in_document(path, "cat", "cats")

    assert count == 3
    assert _paragraph_texts(path) == ["cats cats cats"]


def test_global_replace_with_longer_replacement_does_not_drift_across_paragraphs(
    tmp_path: Path,
) -> None:
    path = tmp_path / "multi_paragraph.docx"
    body = (
        "<w:p><w:r><w:t>cat and cat</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>another cat here</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    count = replace_text_in_document(path, "cat", "elephant")

    assert count == 3
    assert _paragraph_texts(path) == [
        "elephant and elephant",
        "another elephant here",
    ]


def test_atomicity_no_partial_write_when_one_of_several_matches_is_rejected(
    tmp_path: Path,
) -> None:
    """A global replace touching multiple paragraphs, where one match spans a
    run-container boundary, must write nothing at all - not even the other,
    perfectly valid matches (specs/replace_text.md §7)."""
    path = tmp_path / "partial_reject.docx"
    body = (
        "<w:p><w:r><w:t>cat cat</w:t></w:r></w:p>"
        "<w:p><w:ins><w:r><w:t>c</w:t></w:r></w:ins><w:r><w:t>at</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(TextEditError, match="run-container boundary"):
        replace_text_in_document(path, "cat", "dog")

    assert _paragraph_texts(path) == ["cat cat", "cat"]


# --- Multi-run (3+) splits -------------------------------------------------------


def test_replace_spanning_three_runs_with_prefix_and_suffix_kept_on_both_ends(
    tmp_path: Path,
) -> None:
    """The general, "sandwich" case: the leftmost run keeps an untouched
    prefix, the middle run is fully consumed and removed, and the rightmost
    run keeps an untouched suffix - all three repositioning rules exercised
    together in one match (the regression this whole batch was written to
    catch: the replacement must land between the kept prefix and the kept
    suffix, not after the suffix)."""
    path = tmp_path / "sandwich.docx"
    body = (
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>PREFIX-rec</w:t></w:r>"
        "<w:r><w:t>ei</w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>ve-SUFFIX</w:t></w:r>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    text = "PREFIX-receive-SUFFIX"
    start = text.index("receive")

    count = replace_text_in_document(
        path, "receive", "X", location=TextLocation(0, start, start + 7)
    )

    assert count == 1
    assert _paragraph_texts(path) == ["PREFIX-X-SUFFIX"]
    paragraph = _paragraphs(path)[0]
    runs = _runs(paragraph)
    texts = [r.findtext("w:t", namespaces=NSMAP) for r in runs]
    assert texts == ["PREFIX-", "X", "-SUFFIX"]
    assert runs[0].find("w:rPr", namespaces=NSMAP).find(_W_B) is not None
    assert runs[2].find("w:rPr", namespaces=NSMAP).find(_W_I) is not None


def test_replace_prefix_empty_match_on_a_non_first_paragraph(run_split_docx: Path) -> None:
    """The insertion-side fix (addprevious when the leftmost matched node
    has no prefix) must generalize beyond paragraph 0 - regresses against
    the exact bug class this test batch found."""
    count = replace_text_in_document(run_split_docx, "This", "That", location=TextLocation(1, 0, 4))

    assert count == 1
    assert _paragraph_texts(run_split_docx)[1] == "That is a simple example sentence."


def test_replace_spanning_three_runs_removes_the_fully_consumed_middle_run(
    tmp_path: Path,
) -> None:
    path = tmp_path / "three_runs.docx"
    body = (
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>r</w:t></w:r>"
        "<w:r><w:t>ecei</w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>ve</w:t></w:r>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    matches = find_text_matches(path, "receive")
    assert matches[0].location == TextLocation(0, 0, 7)

    count = replace_text_in_document(path, "receive", "X", location=matches[0].location)

    assert count == 1
    assert _paragraph_texts(path) == ["X"]
    paragraph = _paragraphs(path)[0]
    runs = _runs(paragraph)
    assert len(runs) == 1
    assert runs[0].find("w:rPr", namespaces=NSMAP).find(_W_B) is not None


# --- Run-container boundaries: w:ins (allowed inside, rejected across) ---------


def test_match_fully_inside_one_w_ins_block_is_allowed(tmp_path: Path) -> None:
    """Text entirely inside one tracked-change insertion remains a normal,
    matchable/replaceable target - `paragraph_plain_text` already includes
    it as current content (see module docstring)."""
    path = tmp_path / "inside_ins.docx"
    body = "<w:p><w:ins><w:r><w:t>hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:ins></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    count = replace_text_in_document(path, "hello world", "greetings")

    assert count == 1
    assert _paragraph_texts(path) == ["greetings"]


def test_match_spanning_into_w_ins_from_plain_text_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "ins_boundary.docx"
    body = "<w:p><w:ins><w:r><w:t>Hello wo</w:t></w:r></w:ins><w:r><w:t>rld</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(TextEditError, match="run-container boundary"):
        replace_text_in_document(path, "world", "earth")


def test_match_spanning_two_different_w_ins_blocks_is_rejected(tmp_path: Path) -> None:
    """Two distinct `w:ins` elements (same tag, different identity) must
    still be treated as different containers."""
    path = tmp_path / "two_ins.docx"
    body = (
        "<w:p><w:ins><w:r><w:t>ab</w:t></w:r></w:ins><w:ins><w:r><w:t>cd</w:t></w:r></w:ins></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(TextEditError, match="run-container boundary"):
        replace_text_in_document(path, "bc", "X")


def test_match_fully_inside_one_hyperlink_spanning_two_runs_is_allowed(tmp_path: Path) -> None:
    path = tmp_path / "inside_hyperlink.docx"
    body = "<w:p><w:hyperlink><w:r><w:t>cl</w:t></w:r><w:r><w:t>ick</w:t></w:r></w:hyperlink></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    count = replace_text_in_document(path, "click", "tap")

    assert count == 1
    assert _paragraph_texts(path) == ["tap"]


# --- Whole-paragraph deletion -----------------------------------------------------


def test_replace_entire_paragraph_text_with_empty_string(run_split_docx: Path) -> None:
    full_text = _paragraph_texts(run_split_docx)[1]

    count = replace_text_in_document(
        run_split_docx, full_text, "", location=TextLocation(1, 0, len(full_text))
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[1] == ""
    paragraph = _paragraphs(run_split_docx)[1]
    assert _runs(paragraph) == []
    # The document must still be a valid, re-readable .docx afterward.
    assert find_text_matches(run_split_docx, "receive")  # paragraph 0 unaffected


# --- Case sensitivity edge cases --------------------------------------------------


def test_global_replace_case_insensitive_replaces_all_casings(tmp_path: Path) -> None:
    path = tmp_path / "mixed_case.docx"
    body = "<w:p><w:r><w:t>Test test TEST tESt</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    count = replace_text_in_document(path, "test", "exam", case_sensitive=False)

    assert count == 4
    assert _paragraph_texts(path) == ["exam exam exam exam"]


def test_find_case_insensitive_each_match_keeps_its_own_original_casing(
    tmp_path: Path,
) -> None:
    path = tmp_path / "mixed_case_find.docx"
    body = "<w:p><w:r><w:t>Test test TEST</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    matches = find_text_matches(path, "test", case_sensitive=False)

    assert [m.matched_text for m in matches] == ["Test", "test", "TEST"]


def test_replace_stale_check_honors_case_sensitive_true(run_split_docx: Path) -> None:
    """A location whose text matches only case-insensitively must still be
    rejected as stale when `case_sensitive=True` (the default)."""
    location = TextLocation(0, 7, 14)  # "receive", lowercase, in run_split.docx

    with pytest.raises(TextEditError, match="stale"):
        replace_text_in_document(
            run_split_docx, "RECEIVE", "X", location=location, case_sensitive=True
        )


def test_replace_stale_check_honors_case_sensitive_false(run_split_docx: Path) -> None:
    count = replace_text_in_document(
        run_split_docx, "RECEIVE", "X", location=TextLocation(0, 7, 14), case_sensitive=False
    )

    assert count == 1
    assert _paragraph_texts(run_split_docx)[0] == "Please X this message."


# --- Overlap / Unicode edge cases --------------------------------------------------


def test_find_non_overlapping_matches_do_not_double_count(tmp_path: Path) -> None:
    """Searching "aa" in "aaaa" must find two matches ([0,2) and [2,4)), not
    three overlapping ones - matching plain find/replace semantics."""
    path = tmp_path / "overlap.docx"
    body = "<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    matches = find_text_matches(path, "aa")

    assert [m.location for m in matches] == [TextLocation(0, 0, 2), TextLocation(0, 2, 4)]


def test_find_case_insensitive_length_changing_casefold_does_not_desync_later_offsets(
    tmp_path: Path,
) -> None:
    """A length-changing lowercase mapping (Turkish dotted capital I,
    `'İ'.lower() == 'i̇'`, two characters) earlier in the paragraph must not
    shift the offsets computed for a later, unrelated match - proof that
    `_find_spans` never lower-cases the whole string up front (module
    docstring, `_find_spans`)."""
    path = tmp_path / "turkish_i.docx"
    body = "<w:p><w:r><w:t>İstanbul is here</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    text = "İstanbul is here"

    matches = find_text_matches(path, "here", case_sensitive=False)

    expected_start = text.index("here")
    assert matches[0].location == TextLocation(0, expected_start, expected_start + 4)


def test_find_case_insensitive_does_not_match_a_length_changing_casefold_occurrence(
    tmp_path: Path,
) -> None:
    """The flip side of the offset-safety test above, and the documented
    trade-off (specs/find_text.md §5): case-insensitive search for
    "istanbul" does not find "İstanbul" (Turkish dotted capital I, U+0130),
    since `'İ'.lower()` is two characters and `_find_spans` only ever
    compares fixed, `len(search_text)`-sized windows - never a whole-string
    case-fold that could desync offsets. It still finds every occurrence
    whose casefold is length-preserving."""
    path = tmp_path / "turkish_i_miss.docx"
    body = "<w:p><w:r><w:t>İstanbul ISTANBUL istanbul</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    matches = find_text_matches(path, "istanbul", case_sensitive=False)

    assert [m.matched_text for m in matches] == ["ISTANBUL", "istanbul"]


def test_replace_astral_character_in_search_and_replacement_text(tmp_path: Path) -> None:
    path = tmp_path / "astral.docx"
    body = "<w:p><w:r><w:t>Great job \U0001f600 today</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    count = replace_text_in_document(path, "\U0001f600", "\U0001f389")

    assert count == 1
    assert _paragraph_texts(path) == ["Great job \U0001f389 today"]


def test_replace_match_at_absolute_start_and_end_of_paragraph(tmp_path: Path) -> None:
    path = tmp_path / "start_end.docx"
    body = "<w:p><w:r><w:t>START-middle-END</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    replace_text_in_document(path, "START", "BEGIN", location=TextLocation(0, 0, 5))
    replace_text_in_document(path, "END", "FINISH", location=TextLocation(0, 13, 16))

    assert _paragraph_texts(path) == ["BEGIN-middle-FINISH"]


# --- Formatting independence (deep-copy, not aliasing) ---------------------------


def test_replacement_rpr_is_an_independent_deep_copy_not_shared_with_original(
    run_split_docx: Path,
) -> None:
    replace_text_in_document(run_split_docx, "receive", "REPLIED", location=TextLocation(0, 7, 14))

    paragraph = _paragraphs(run_split_docx)[0]
    runs = _runs(paragraph)
    replacement_rpr = runs[1].find("w:rPr", namespaces=NSMAP)
    replacement_rpr.append(etree.SubElement(replacement_rpr, f"{{{WORD_NS}}}i"))

    # Mutating the replacement run's rPr after the fact must not be possible
    # to have retroactively affected any other run - each rPr is its own
    # element, never the same object referenced twice in the tree.
    other_rprs = [r.find("w:rPr", namespaces=NSMAP) for r in runs if r is not runs[1]]
    for rpr in other_rprs:
        if rpr is not None:
            assert rpr.find(_W_I) is None


# --- Sequential edits (round-trip stability) --------------------------------------


def test_multiple_sequential_edits_round_trip_correctly(run_split_docx: Path) -> None:
    """Three successive `replace_text_in_document` calls against the same
    file, each re-opening and re-writing it, must all apply correctly and
    leave a document the next call (and every read tool) can still parse."""
    replace_text_in_document(run_split_docx, "receive", "REPLIED", location=TextLocation(0, 7, 14))
    replace_text_in_document(run_split_docx, "simple", "basic", location=TextLocation(1, 10, 16))
    count = replace_text_in_document(run_split_docx, "test", "exam")

    assert count == 4
    texts = _paragraph_texts(run_split_docx)
    assert texts[0] == "Please REPLIED this message."
    assert texts[1] == "This is a basic example sentence."
    assert texts[3] == "This paragraph mentions exam, exam, and exam again."
    assert texts[4] == "One more exam paragraph for global replace."


# --- Crash simulation at the replace_text_in_document level ---------------------


def test_replace_simulated_crash_leaves_original_file_untouched(
    run_split_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A crash during the underlying atomic write must leave the document
    byte-for-byte unchanged, proven at the `replace_text_in_document` level
    directly - not only for the shared `atomic_write_part` helper in
    isolation (docs/security-model.md §5)."""
    original_bytes = run_split_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml_module.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        replace_text_in_document(run_split_docx, "receive", "REPLIED")

    assert run_split_docx.read_bytes() == original_bytes
