"""Tests for `docx_mcp.paragraph_edit`, from `specs/insert_paragraph.md`/
`specs/delete_paragraph.md`.

Covers the functional, formatting-fidelity, and error/edge test categories
required by [docs/testing.md §2](../docs/testing.md#2-test-types-per-mcp-tool)
and [docs/security-model.md §5](../docs/security-model.md#5-testing-obligations),
following the same synthetic-XML-fixture convention `tests/test_text_edit.py`
established for edge cases that need a specific hand-crafted `w:pPr`, plus the
real `tests/fixtures/paragraph_edits.docx` fixture (Phase 5) for the ordinary
functional happy paths and the existing `structured.docx`/`footnotes.docx`
fixtures for the custom-heading-style boundary and footnote-survival cases.
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
from docx_mcp.paragraph_edit import (
    ParagraphEditError,
    delete_paragraph_from_document,
    insert_paragraph_in_document,
)

_W_PPR = f"{{{WORD_NS}}}pPr"
_W_SECTPR = f"{{{WORD_NS}}}sectPr"


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_xml(body_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{body_inner_xml}</w:body></w:document>'
    ).encode()


def _document_root(docx_path: Path) -> etree._Element:
    with validate_and_open(docx_path) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)
    return parse_xml(raw_xml, part_name=DOCUMENT_PART)


def _body(docx_path: Path) -> etree._Element:
    return get_body(_document_root(docx_path))


def _paragraphs(docx_path: Path) -> list[etree._Element]:
    return _body(docx_path).findall("w:p", namespaces=NSMAP)


def _paragraph_texts(docx_path: Path) -> list[str]:
    return [paragraph_plain_text(p) for p in _paragraphs(docx_path)]


def _ppr_xml(paragraph: etree._Element) -> bytes | None:
    ppr = paragraph.find("w:pPr", namespaces=NSMAP)
    return etree.tostring(ppr) if ppr is not None else None


# --- Functional: insert_paragraph_in_document, style inheritance --------------


def test_insert_after_heading_does_not_become_a_heading(paragraph_edits_docx: Path) -> None:
    """Anchor (paragraph 0, "Introduction") is a heading; inserting plain
    body text after it must NOT copy the heading's own pStyle - the specific
    bug class ADR-0006 reasons through explicitly."""
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "New body text.", after_paragraph_index=0
    )

    assert new_index == 1
    paragraphs = _paragraphs(paragraph_edits_docx)
    assert paragraph_plain_text(paragraphs[1]) == "New body text."
    assert _ppr_xml(paragraphs[1]) is None


def test_insert_after_body_paragraph_copies_ppr_verbatim(paragraph_edits_docx: Path) -> None:
    """Anchor (paragraph 1, centered + left-indented) is not a heading; the
    new paragraph must carry the *entire* w:pPr, not just a shared style id."""
    before_ppr = _ppr_xml(_paragraphs(paragraph_edits_docx)[1])
    assert before_ppr is not None  # sanity: the fixture really has local formatting here

    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Inherits formatting.", after_paragraph_index=1
    )

    assert new_index == 2
    paragraphs = _paragraphs(paragraph_edits_docx)
    assert paragraph_plain_text(paragraphs[2]) == "Inherits formatting."
    assert _ppr_xml(paragraphs[2]) == before_ppr
    # Untouched: the anchor itself keeps its own, unmodified pPr.
    assert _ppr_xml(paragraphs[1]) == before_ppr


def test_insert_ppr_is_independent_deep_copy_not_shared_with_anchor(
    paragraph_edits_docx: Path,
) -> None:
    insert_paragraph_in_document(
        paragraph_edits_docx, "Inherits formatting.", after_paragraph_index=1
    )

    paragraphs = _paragraphs(paragraph_edits_docx)
    anchor_ppr = paragraphs[1].find("w:pPr", namespaces=NSMAP)
    new_ppr = paragraphs[2].find("w:pPr", namespaces=NSMAP)
    assert anchor_ppr is not new_ppr
    new_ppr.append(etree.SubElement(new_ppr, f"{{{WORD_NS}}}keepNext"))
    assert anchor_ppr.find(f"{{{WORD_NS}}}keepNext") is None


def test_insert_at_end_of_document(paragraph_edits_docx: Path) -> None:
    last_index = len(_paragraphs(paragraph_edits_docx)) - 1

    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Appended at the end.", after_paragraph_index=last_index
    )

    assert new_index == last_index + 1
    texts = _paragraph_texts(paragraph_edits_docx)
    assert texts[-1] == "Appended at the end."
    # Still a valid, re-readable document: the body's sectPr stays its last child.
    body = _body(paragraph_edits_docx)
    assert etree.QName(list(body)[-1]).localname == "sectPr"


def test_insert_at_very_start_non_empty_document(paragraph_edits_docx: Path) -> None:
    """after_paragraph_index=None: the current paragraph 0 ("Introduction", a
    heading) becomes the context; per the anchor-is-a-heading rule, the new
    first paragraph gets no pPr, and the old paragraph 0 shifts to index 1."""
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "New first paragraph.", after_paragraph_index=None
    )

    assert new_index == 0
    paragraphs = _paragraphs(paragraph_edits_docx)
    assert paragraph_plain_text(paragraphs[0]) == "New first paragraph."
    assert _ppr_xml(paragraphs[0]) is None
    assert paragraph_plain_text(paragraphs[1]) == "Introduction"


def test_insert_into_completely_empty_document(tmp_path: Path) -> None:
    path = tmp_path / "empty_body.docx"
    _write_zip(path, {"word/document.xml": _document_xml("")})

    new_index = insert_paragraph_in_document(path, "Only paragraph.", after_paragraph_index=None)

    assert new_index == 0
    paragraphs = _paragraphs(path)
    assert len(paragraphs) == 1
    assert paragraph_plain_text(paragraphs[0]) == "Only paragraph."
    assert _ppr_xml(paragraphs[0]) is None


@pytest.mark.parametrize("level", range(1, 10))
def test_insert_heading_at_each_level(paragraph_edits_docx: Path, level: int) -> None:
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "A new heading.", after_paragraph_index=3, heading_level=level
    )

    paragraph = _paragraphs(paragraph_edits_docx)[new_index]
    ppr = paragraph.find("w:pPr", namespaces=NSMAP)
    assert ppr is not None
    assert len(ppr) == 1
    style = ppr.find("w:pStyle", namespaces=NSMAP)
    assert style is not None
    assert style.get(f"{{{WORD_NS}}}val") == f"Heading{level}"


def test_insert_heading_does_not_inherit_anchor_formatting(paragraph_edits_docx: Path) -> None:
    """Anchor (paragraph 1) has centered/indented local formatting; a
    heading_level insert after it must get a fresh pPr, not that formatting."""
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Fresh heading.", after_paragraph_index=1, heading_level=2
    )

    ppr = _paragraphs(paragraph_edits_docx)[new_index].find("w:pPr", namespaces=NSMAP)
    assert len(ppr) == 1
    assert ppr.find("w:ind", namespaces=NSMAP) is None
    assert ppr.find("w:jc", namespaces=NSMAP) is None


def test_insert_empty_string_paragraph_has_no_runs(paragraph_edits_docx: Path) -> None:
    new_index = insert_paragraph_in_document(paragraph_edits_docx, "", after_paragraph_index=3)

    paragraph = _paragraphs(paragraph_edits_docx)[new_index]
    assert paragraph_plain_text(paragraph) == ""
    assert paragraph.findall("w:r", namespaces=NSMAP) == []


def test_insert_after_numbered_list_item_copies_numpr(tmp_path: Path) -> None:
    path = tmp_path / "numbered.docx"
    body = (
        "<w:p>"
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        "<w:r><w:t>First list item.</w:t></w:r>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    new_index = insert_paragraph_in_document(path, "Second list item.", after_paragraph_index=0)

    new_ppr = _paragraphs(path)[new_index].find("w:pPr", namespaces=NSMAP)
    assert new_ppr is not None
    num_id = new_ppr.find("w:numPr/w:numId", namespaces=NSMAP)
    assert num_id.get(f"{{{WORD_NS}}}val") == "1"


def test_insert_after_anchor_carrying_sectpr_does_not_duplicate_it(tmp_path: Path) -> None:
    """Regression test for a bug found during the post-implementation
    verification pass (ADR-0006 §Decision.8): the anchor is not a heading,
    so the naive rule would copy its *entire* pPr, including w:sectPr -
    duplicating the document's section properties onto the new paragraph
    and silently creating an unintended extra section break. Only the
    anchor's other formatting (w:jc here) must be inherited; w:sectPr must
    never be copied, and the anchor itself must keep its own untouched."""
    path = tmp_path / "section_with_formatting.docx"
    body = (
        "<w:p><w:r><w:t>End of first section.</w:t></w:r>"
        '<w:pPr><w:jc w:val="center"/>'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>'
        "</w:p>"
        "<w:p><w:r><w:t>Second section content.</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    new_index = insert_paragraph_in_document(path, "New text.", after_paragraph_index=0)

    paragraphs = _paragraphs(path)
    new_ppr = paragraphs[new_index].find("w:pPr", namespaces=NSMAP)
    assert new_ppr is not None
    assert new_ppr.find("w:jc", namespaces=NSMAP) is not None
    assert new_ppr.find("w:sectPr", namespaces=NSMAP) is None
    # The anchor itself keeps its own w:sectPr, untouched.
    anchor_ppr = paragraphs[0].find("w:pPr", namespaces=NSMAP)
    assert anchor_ppr.find("w:sectPr", namespaces=NSMAP) is not None


def test_insert_after_anchor_whose_ppr_is_only_sectpr_gets_no_ppr_at_all(
    tmp_path: Path,
) -> None:
    """When stripping w:sectPr leaves nothing else in the copied pPr, no
    pPr at all is attached to the new paragraph - not an empty <w:pPr/>."""
    path = tmp_path / "section_only.docx"
    body = (
        "<w:p><w:r><w:t>End of first section.</w:t></w:r>"
        "<w:pPr><w:sectPr/></w:pPr></w:p>"
        "<w:p><w:r><w:t>Second section content.</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    new_index = insert_paragraph_in_document(path, "New text.", after_paragraph_index=0)

    assert _ppr_xml(_paragraphs(path)[new_index]) is None
    assert _paragraphs(path)[0].find("w:pPr/w:sectPr", namespaces=NSMAP) is not None


def test_insert_after_custom_named_heading_style_is_not_recognized_as_a_heading(
    structured_docx: Path,
) -> None:
    """`structured.docx`'s paragraph 3 ("Custom styled heading text") only
    resolves to heading level 2 through its styles.xml basedOn chain
    (get_structure's fuller resolution) - this tool deliberately uses only
    the built-in-id heuristic, so it must treat it as an ordinary body
    paragraph and copy its pPr verbatim, per ADR-0006 §Decision.4."""
    anchor_before = _ppr_xml(_paragraphs(structured_docx)[3])

    new_index = insert_paragraph_in_document(
        structured_docx, "Plain follow-up text.", after_paragraph_index=3
    )

    assert _ppr_xml(_paragraphs(structured_docx)[new_index]) == anchor_before


def test_insert_does_not_alter_other_paragraphs_or_zip_parts(paragraph_edits_docx: Path) -> None:
    before_paragraphs = _paragraphs(paragraph_edits_docx)
    untouched_before = {i: etree.tostring(p) for i, p in enumerate(before_paragraphs)}
    with zipfile.ZipFile(paragraph_edits_docx) as archive:
        other_parts_before = {
            name: archive.read(name) for name in archive.namelist() if name != DOCUMENT_PART
        }

    insert_paragraph_in_document(paragraph_edits_docx, "New text.", after_paragraph_index=1)

    after_paragraphs = _paragraphs(paragraph_edits_docx)
    for i, xml_before in untouched_before.items():
        index_after = i if i <= 1 else i + 1
        assert etree.tostring(after_paragraphs[index_after]) == xml_before

    with zipfile.ZipFile(paragraph_edits_docx) as archive:
        other_parts_after = {
            name: archive.read(name) for name in archive.namelist() if name != DOCUMENT_PART
        }
    assert other_parts_after == other_parts_before


def test_insert_returned_index_addresses_exactly_the_new_paragraph(
    paragraph_edits_docx: Path,
) -> None:
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Findable new text.", after_paragraph_index=2
    )

    assert _paragraph_texts(paragraph_edits_docx)[new_index] == "Findable new text."


# --- Error / edge: insert_paragraph_in_document --------------------------------


def test_insert_after_paragraph_index_negative_is_rejected(paragraph_edits_docx: Path) -> None:
    with pytest.raises(ParagraphEditError, match="after_paragraph_index"):
        insert_paragraph_in_document(paragraph_edits_docx, "x", after_paragraph_index=-1)


def test_insert_after_paragraph_index_equal_to_count_is_rejected(
    paragraph_edits_docx: Path,
) -> None:
    count = len(_paragraphs(paragraph_edits_docx))
    with pytest.raises(ParagraphEditError, match="after_paragraph_index"):
        insert_paragraph_in_document(paragraph_edits_docx, "x", after_paragraph_index=count)


def test_insert_after_paragraph_index_far_out_of_range_is_rejected(
    paragraph_edits_docx: Path,
) -> None:
    with pytest.raises(ParagraphEditError, match="after_paragraph_index"):
        insert_paragraph_in_document(paragraph_edits_docx, "x", after_paragraph_index=999)


@pytest.mark.parametrize("level", [0, 10, -1])
def test_insert_heading_level_out_of_range_is_rejected(
    paragraph_edits_docx: Path, level: int
) -> None:
    with pytest.raises(ParagraphEditError, match="heading_level"):
        insert_paragraph_in_document(
            paragraph_edits_docx, "x", after_paragraph_index=0, heading_level=level
        )


def test_insert_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"
    with pytest.raises(InvalidDocumentError):
        insert_paragraph_in_document(missing, "x")


def test_insert_not_a_valid_zip_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not_a_zip.docx"
    path.write_bytes(b"this is not a zip file")

    with pytest.raises(InvalidDocumentError):
        insert_paragraph_in_document(path, "x")


def test_insert_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_document_xml.docx"
    _write_zip(path, {"word/styles.xml": b"<styles/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        insert_paragraph_in_document(path, "x")


def test_insert_malformed_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed.docx"
    _write_zip(path, {"word/document.xml": b"<w:document><unclosed>"})

    with pytest.raises(InvalidDocumentError):
        insert_paragraph_in_document(path, "x")


def test_insert_missing_body_element_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_body.docx"
    xml = f'<w:document xmlns:w="{WORD_NS}"></w:document>'.encode()
    _write_zip(path, {"word/document.xml": xml})

    with pytest.raises(InvalidDocumentError, match="body"):
        insert_paragraph_in_document(path, "x")


def test_insert_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    _write_zip(
        path, {"word/document.xml": _document_xml("<w:p><w:r><w:t>Small.</w:t></w:r></w:p>")}
    )
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        insert_paragraph_in_document(path, "x", max_size_bytes=actual_size - 1)


def test_insert_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    _write_zip(
        path, {"word/document.xml": _document_xml("<w:p><w:r><w:t>Small.</w:t></w:r></w:p>")}
    )
    actual_size = path.stat().st_size

    insert_paragraph_in_document(path, "x", after_paragraph_index=0, max_size_bytes=actual_size)
    assert _paragraph_texts(path)[1] == "x"


def test_insert_xxe_entity_is_never_resolved_and_insert_still_succeeds(tmp_path: Path) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe.docx"
    malicious_xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:document [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>'
        "<w:p><w:r><w:t>Safe text.</w:t></w:r></w:p>"
        "</w:body></w:document>"
    ).encode()
    _write_zip(path, {"word/document.xml": malicious_xml})

    new_index = insert_paragraph_in_document(path, "New text.", after_paragraph_index=0)

    with zipfile.ZipFile(path) as archive:
        raw = archive.read(DOCUMENT_PART)
    assert new_index == 1
    assert b"New text." in raw
    assert b"TOP-SECRET-CONTENT" not in raw


def test_insert_simulated_crash_leaves_original_file_untouched(
    paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original_bytes = paragraph_edits_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml_module.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        insert_paragraph_in_document(paragraph_edits_docx, "x", after_paragraph_index=0)

    assert paragraph_edits_docx.read_bytes() == original_bytes


# --- Functional: delete_paragraph_from_document --------------------------------


def test_delete_middle_body_paragraph(paragraph_edits_docx: Path) -> None:
    before = _paragraph_texts(paragraph_edits_docx)

    deleted_text = delete_paragraph_from_document(paragraph_edits_docx, 1)

    assert deleted_text == "First body paragraph."
    after = _paragraph_texts(paragraph_edits_docx)
    assert after == before[:1] + before[2:]


def test_delete_first_paragraph(paragraph_edits_docx: Path) -> None:
    deleted_text = delete_paragraph_from_document(paragraph_edits_docx, 0)

    assert deleted_text == "Introduction"
    assert _paragraph_texts(paragraph_edits_docx)[0] == "First body paragraph."


def test_delete_last_paragraph(paragraph_edits_docx: Path) -> None:
    last_index = len(_paragraphs(paragraph_edits_docx)) - 1

    deleted_text = delete_paragraph_from_document(paragraph_edits_docx, last_index)

    assert deleted_text == "Last paragraph."
    body = _body(paragraph_edits_docx)
    assert etree.QName(list(body)[-1]).localname == "sectPr"


def test_delete_heading_paragraph_returns_marker_free_text(paragraph_edits_docx: Path) -> None:
    deleted_text = delete_paragraph_from_document(paragraph_edits_docx, 0)

    assert deleted_text == "Introduction"  # no "# " prefix - marker-free, per get_structure


def test_delete_with_matching_expected_text_succeeds(paragraph_edits_docx: Path) -> None:
    deleted_text = delete_paragraph_from_document(
        paragraph_edits_docx, 1, expected_text="First body paragraph."
    )

    assert deleted_text == "First body paragraph."


def test_delete_without_expected_text_succeeds_unconditionally(
    paragraph_edits_docx: Path,
) -> None:
    deleted_text = delete_paragraph_from_document(paragraph_edits_docx, 1, expected_text=None)

    assert deleted_text == "First body paragraph."


def test_delete_does_not_alter_other_paragraphs_or_zip_parts(paragraph_edits_docx: Path) -> None:
    before_paragraphs = _paragraphs(paragraph_edits_docx)
    untouched_before = {i: etree.tostring(p) for i, p in enumerate(before_paragraphs) if i != 1}
    with zipfile.ZipFile(paragraph_edits_docx) as archive:
        other_parts_before = {
            name: archive.read(name) for name in archive.namelist() if name != DOCUMENT_PART
        }

    delete_paragraph_from_document(paragraph_edits_docx, 1)

    after_paragraphs = _paragraphs(paragraph_edits_docx)
    for i, xml_before in untouched_before.items():
        index_after = i if i < 1 else i - 1
        assert etree.tostring(after_paragraphs[index_after]) == xml_before

    with zipfile.ZipFile(paragraph_edits_docx) as archive:
        other_parts_after = {
            name: archive.read(name) for name in archive.namelist() if name != DOCUMENT_PART
        }
    assert other_parts_after == other_parts_before


def test_delete_paragraph_anchoring_a_footnote_leaves_footnotes_xml_untouched(
    footnotes_docx: Path,
) -> None:
    """`footnotes.docx` paragraph 2 anchors footnote ids 1 and 2; paragraph 4
    anchors id 3. Deleting paragraph 2 must remove both anchors from
    get_footnotes' output while leaving a *different* paragraph's footnote
    (id 3) correctly resolved, and word/footnotes.xml byte-identical -
    the orphaned entries for ids 1/2 are deliberately left in place
    (ADR-0006 §Decision.5)."""
    with zipfile.ZipFile(footnotes_docx) as archive:
        footnotes_xml_before = archive.read("word/footnotes.xml")

    deleted_text = delete_paragraph_from_document(footnotes_docx, 2)

    assert deleted_text == "This paragraph has two footnotes."
    remaining = get_document_footnotes(footnotes_docx)
    assert [(f.id, f.content) for f in remaining] == [("3", " Third footnote.")]
    with zipfile.ZipFile(footnotes_docx) as archive:
        footnotes_xml_after = archive.read("word/footnotes.xml")
    assert footnotes_xml_after == footnotes_xml_before


# --- Section-properties guard ---------------------------------------------------


def test_delete_paragraph_carrying_sectpr_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "section_break.docx"
    body = (
        "<w:p><w:r><w:t>End of first section.</w:t></w:r>"
        "<w:pPr><w:sectPr/></w:pPr></w:p>"
        "<w:p><w:r><w:t>Second section content.</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(ParagraphEditError, match="sectPr|section"):
        delete_paragraph_from_document(path, 0)

    assert _paragraph_texts(path) == ["End of first section.", "Second section content."]


def test_delete_other_paragraph_in_a_multi_section_document_still_succeeds(
    tmp_path: Path,
) -> None:
    path = tmp_path / "section_break.docx"
    body = (
        "<w:p><w:r><w:t>End of first section.</w:t></w:r>"
        "<w:pPr><w:sectPr/></w:pPr></w:p>"
        "<w:p><w:r><w:t>Second section content.</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    deleted_text = delete_paragraph_from_document(path, 1)

    assert deleted_text == "Second section content."
    assert _paragraph_texts(path) == ["End of first section."]


# --- Error / edge: delete_paragraph_from_document ------------------------------


def test_delete_paragraph_index_negative_is_rejected(paragraph_edits_docx: Path) -> None:
    with pytest.raises(ParagraphEditError, match="paragraph_index"):
        delete_paragraph_from_document(paragraph_edits_docx, -1)


def test_delete_paragraph_index_equal_to_count_is_rejected(paragraph_edits_docx: Path) -> None:
    count = len(_paragraphs(paragraph_edits_docx))
    with pytest.raises(ParagraphEditError, match="paragraph_index"):
        delete_paragraph_from_document(paragraph_edits_docx, count)


def test_delete_paragraph_index_far_out_of_range_is_rejected(paragraph_edits_docx: Path) -> None:
    with pytest.raises(ParagraphEditError, match="paragraph_index"):
        delete_paragraph_from_document(paragraph_edits_docx, 999)


def test_delete_mismatched_expected_text_is_rejected_as_stale(paragraph_edits_docx: Path) -> None:
    with pytest.raises(ParagraphEditError, match="stale"):
        delete_paragraph_from_document(paragraph_edits_docx, 1, expected_text="Wrong text.")

    assert _paragraph_texts(paragraph_edits_docx)[1] == "First body paragraph."


def test_delete_case_different_expected_text_is_rejected_as_stale(
    paragraph_edits_docx: Path,
) -> None:
    """No case_sensitive option: expected_text is exact-match only."""
    with pytest.raises(ParagraphEditError, match="stale"):
        delete_paragraph_from_document(
            paragraph_edits_docx, 1, expected_text="first body paragraph."
        )


def test_delete_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"
    with pytest.raises(InvalidDocumentError):
        delete_paragraph_from_document(missing, 0)


def test_delete_not_a_valid_zip_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not_a_zip.docx"
    path.write_bytes(b"this is not a zip file")

    with pytest.raises(InvalidDocumentError):
        delete_paragraph_from_document(path, 0)


def test_delete_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_document_xml.docx"
    _write_zip(path, {"word/styles.xml": b"<styles/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        delete_paragraph_from_document(path, 0)


def test_delete_malformed_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed.docx"
    _write_zip(path, {"word/document.xml": b"<w:document><unclosed>"})

    with pytest.raises(InvalidDocumentError):
        delete_paragraph_from_document(path, 0)


def test_delete_missing_body_element_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_body.docx"
    xml = f'<w:document xmlns:w="{WORD_NS}"></w:document>'.encode()
    _write_zip(path, {"word/document.xml": xml})

    with pytest.raises(InvalidDocumentError, match="body"):
        delete_paragraph_from_document(path, 0)


def test_delete_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(
                "<w:p><w:r><w:t>Small.</w:t></w:r></w:p><w:p><w:r><w:t>Other.</w:t></w:r></w:p>"
            )
        },
    )
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        delete_paragraph_from_document(path, 0, max_size_bytes=actual_size - 1)


def test_delete_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(
                "<w:p><w:r><w:t>Small.</w:t></w:r></w:p><w:p><w:r><w:t>Other.</w:t></w:r></w:p>"
            )
        },
    )
    actual_size = path.stat().st_size

    delete_paragraph_from_document(path, 0, max_size_bytes=actual_size)
    assert _paragraph_texts(path) == ["Other."]


def test_delete_xxe_entity_is_never_resolved_and_delete_still_succeeds(tmp_path: Path) -> None:
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

    deleted_text = delete_paragraph_from_document(path, 0)

    with zipfile.ZipFile(path) as archive:
        raw = archive.read(DOCUMENT_PART)
    assert deleted_text == "Safe text."
    assert b"TOP-SECRET-CONTENT" not in raw


def test_delete_simulated_crash_leaves_original_file_untouched(
    paragraph_edits_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original_bytes = paragraph_edits_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml_module.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        delete_paragraph_from_document(paragraph_edits_docx, 0)

    assert paragraph_edits_docx.read_bytes() == original_bytes


# --- Sequential edits (insert then delete round-trips) -------------------------


def test_insert_then_delete_the_same_paragraph_restores_original_state(
    paragraph_edits_docx: Path,
) -> None:
    before = _paragraph_texts(paragraph_edits_docx)

    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Temporary.", after_paragraph_index=2
    )
    delete_paragraph_from_document(paragraph_edits_docx, new_index)

    assert _paragraph_texts(paragraph_edits_docx) == before


def test_multiple_sequential_inserts_round_trip_correctly(paragraph_edits_docx: Path) -> None:
    """Three successive `insert_paragraph_in_document` calls against the same
    file, each re-opening and re-writing it, must all apply correctly and
    leave a document the next call (and every read tool) can still parse -
    mirrors `test_text_edit.py`'s sequential-edit coverage for `replace_text`."""
    first_index = insert_paragraph_in_document(
        paragraph_edits_docx, "One.", after_paragraph_index=0
    )
    second_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Two.", after_paragraph_index=first_index
    )
    third_index = insert_paragraph_in_document(
        paragraph_edits_docx, "Three.", after_paragraph_index=second_index
    )

    texts = _paragraph_texts(paragraph_edits_docx)
    assert texts[first_index : third_index + 1] == ["One.", "Two.", "Three."]


# --- Additional adversarial / edge-case coverage (post-implementation pass) ----


def test_insert_with_after_paragraph_index_zero_into_an_empty_document_is_rejected(
    tmp_path: Path,
) -> None:
    """An empty document has no valid index at all - only `None` may be used
    to insert into it; `0` must still be rejected as out of range, not
    silently treated the same as `None`."""
    path = tmp_path / "empty_body.docx"
    _write_zip(path, {"word/document.xml": _document_xml("")})

    with pytest.raises(ParagraphEditError, match="after_paragraph_index"):
        insert_paragraph_in_document(path, "x", after_paragraph_index=0)


def test_insert_heading_with_empty_text_has_style_but_no_runs(paragraph_edits_docx: Path) -> None:
    new_index = insert_paragraph_in_document(
        paragraph_edits_docx, "", after_paragraph_index=1, heading_level=3
    )

    paragraph = _paragraphs(paragraph_edits_docx)[new_index]
    assert paragraph_plain_text(paragraph) == ""
    assert paragraph.findall("w:r", namespaces=NSMAP) == []
    style = paragraph.find("w:pPr/w:pStyle", namespaces=NSMAP)
    assert style.get(f"{{{WORD_NS}}}val") == "Heading3"


def test_insert_astral_character_in_text_round_trips(paragraph_edits_docx: Path) -> None:
    text = "Great job \U0001f600 today"

    new_index = insert_paragraph_in_document(paragraph_edits_docx, text, after_paragraph_index=1)

    assert _paragraph_texts(paragraph_edits_docx)[new_index] == text


def test_insert_special_xml_characters_are_safely_escaped(paragraph_edits_docx: Path) -> None:
    """lxml serializes `.text` content safely regardless of what characters
    it contains - this pins that `insert_paragraph` never builds XML by
    string concatenation, which would be injectable."""
    text = "<script>alert(1)</script> & \"quotes\" & 'apostrophes'"

    new_index = insert_paragraph_in_document(paragraph_edits_docx, text, after_paragraph_index=1)

    assert _paragraph_texts(paragraph_edits_docx)[new_index] == text
    # The document must still be a single, well-formed w:t text node - not
    # split across accidental sibling elements from an unescaped "<".
    paragraph = _paragraphs(paragraph_edits_docx)[new_index]
    assert len(paragraph.findall("w:r", namespaces=NSMAP)) == 1


def test_delete_the_only_paragraph_in_a_single_paragraph_document(tmp_path: Path) -> None:
    path = tmp_path / "single_paragraph.docx"
    _write_zip(
        path, {"word/document.xml": _document_xml("<w:p><w:r><w:t>Only one.</w:t></w:r></w:p>")}
    )

    deleted_text = delete_paragraph_from_document(path, 0)

    assert deleted_text == "Only one."
    assert _paragraphs(path) == []


def test_delete_rejected_for_sectpr_even_when_expected_text_matches(tmp_path: Path) -> None:
    """The section-properties guard applies regardless of whether the
    staleness check would otherwise have passed - it is not bypassable by
    supplying the correct expected_text."""
    path = tmp_path / "section_break.docx"
    body = (
        "<w:p><w:r><w:t>End of first section.</w:t></w:r>"
        "<w:pPr><w:sectPr/></w:pPr></w:p>"
        "<w:p><w:r><w:t>Second section content.</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(ParagraphEditError, match="sectPr|section"):
        delete_paragraph_from_document(path, 0, expected_text="End of first section.")

    assert _paragraph_texts(path) == ["End of first section.", "Second section content."]
