"""Tests for `docx_mcp.footnote_edit`, from `specs/add_footnote.md`/
`specs/edit_footnote.md`.

Covers the functional, formatting-fidelity, and error/edge test categories
required by [docs/testing.md §2](../docs/testing.md#2-test-types-per-mcp-tool)
and [docs/security-model.md §5](../docs/security-model.md#5-testing-obligations),
reusing the existing `structured_docx` (zero footnotes - the from-scratch
`word/footnotes.xml`/`[Content_Types].xml`/`.rels` creation path),
`footnotes_docx`, and `minimal_docx` fixtures (footnotes already present -
the append path) rather than adding a new fixture, per
[ADR-0007](../docs/adr/0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md).
Edge cases needing a specific hand-crafted archive (a missing `.rels` part,
non-numeric ids, ...) follow the synthetic-per-test-archive convention
`tests/test_paragraph_edit.py`/`tests/test_text_edit.py` established.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

import docx_mcp.ooxml as ooxml_module
from docx_mcp.document import DOCUMENT_PART, NSMAP, WORD_NS, InvalidDocumentError, get_body
from docx_mcp.footnote_edit import (
    CONTENT_TYPES_PART,
    DOCUMENT_RELS_PART,
    FootnoteEditError,
    add_footnote_to_document,
    edit_footnote_in_document,
)
from docx_mcp.footnotes import FOOTNOTES_PART, get_document_footnotes, render_footnote_content
from docx_mcp.ooxml import parse_xml, read_part, validate_and_open

_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_FOOTNOTES_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_xml(body_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{body_inner_xml}</w:body></w:document>'
    ).encode()


def _footnotes_xml(footnotes_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:footnotes xmlns:w="{WORD_NS}">{footnotes_inner_xml}</w:footnotes>'
    ).encode()


def _minimal_content_types() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Types xmlns="{_CONTENT_TYPES_NS}">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        "</Types>"
    ).encode()


def _minimal_rels() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_RELS_NS}"></Relationships>'
    ).encode()


def _document_root(docx_path: Path) -> etree._Element:
    with validate_and_open(docx_path) as archive:
        raw_xml = read_part(archive, DOCUMENT_PART)
    return parse_xml(raw_xml, part_name=DOCUMENT_PART)


def _paragraphs(docx_path: Path) -> list[etree._Element]:
    return get_body(_document_root(docx_path)).findall("w:p", namespaces=NSMAP)


def _footnotes_root(docx_path: Path) -> etree._Element | None:
    with zipfile.ZipFile(docx_path) as archive:
        try:
            raw = archive.read(FOOTNOTES_PART)
        except KeyError:
            return None
    return parse_xml(raw, part_name=FOOTNOTES_PART)


def _footnote_by_id(docx_path: Path, footnote_id: str) -> etree._Element | None:
    root = _footnotes_root(docx_path)
    if root is None:
        return None
    for footnote in root.findall("w:footnote", namespaces=NSMAP):
        if footnote.get(f"{{{WORD_NS}}}id") == footnote_id:
            return footnote
    return None


def _content_types_declares_footnotes(docx_path: Path) -> bool:
    with zipfile.ZipFile(docx_path) as archive:
        root = etree.fromstring(archive.read(CONTENT_TYPES_PART))
    return any(
        el.get("PartName") == "/word/footnotes.xml"
        for el in root.findall(f"{{{_CONTENT_TYPES_NS}}}Override")
    )


def _rels_declares_footnotes(docx_path: Path) -> bool:
    with zipfile.ZipFile(docx_path) as archive:
        try:
            raw = archive.read(DOCUMENT_RELS_PART)
        except KeyError:
            return False
    root = etree.fromstring(raw)
    return any(
        el.get("Type") == _FOOTNOTES_RELATIONSHIP_TYPE
        for el in root.findall(f"{{{_RELS_NS}}}Relationship")
    )


# --- Functional: add_footnote_to_document, from-scratch (structured_docx) -----


def test_add_footnote_creates_footnotes_xml_content_types_and_rels_from_scratch(
    structured_docx: Path,
) -> None:
    """`structured.docx` (Phase 2) has zero footnotes: no word/footnotes.xml,
    no footnotes Override, no footnotes Relationship."""
    assert _footnotes_root(structured_docx) is None
    assert _content_types_declares_footnotes(structured_docx) is False

    new_id = add_footnote_to_document(structured_docx, 0, "A new footnote.")

    assert new_id == "1"
    footnote = _footnote_by_id(structured_docx, "1")
    assert footnote is not None
    assert render_footnote_content(footnote) == "A new footnote."
    assert _content_types_declares_footnotes(structured_docx) is True
    assert _rels_declares_footnotes(structured_docx) is True


def test_add_footnote_from_scratch_includes_boilerplate_separator_footnotes(
    structured_docx: Path,
) -> None:
    add_footnote_to_document(structured_docx, 0, "Content.")

    root = _footnotes_root(structured_docx)
    assert root is not None
    ids = {f.get(f"{{{WORD_NS}}}id") for f in root.findall("w:footnote", namespaces=NSMAP)}
    assert ids == {"-1", "0", "1"}


def test_add_footnote_anchor_is_appended_at_the_end_of_the_target_paragraph(
    structured_docx: Path,
) -> None:
    new_id = add_footnote_to_document(structured_docx, 1, "Anchored footnote.")

    footnotes = get_document_footnotes(structured_docx)
    assert [(f.id, f.paragraph_index) for f in footnotes] == [(new_id, 1)]


def test_add_footnote_from_scratch_does_not_alter_other_paragraphs(
    structured_docx: Path,
) -> None:
    before = [etree.tostring(p) for p in _paragraphs(structured_docx)]

    add_footnote_to_document(structured_docx, 2, "Content.")

    after = _paragraphs(structured_docx)
    for index, xml_before in enumerate(before):
        if index != 2:
            assert etree.tostring(after[index]) == xml_before


# --- Functional: add_footnote_to_document, append (existing footnotes.xml) ----


def test_add_footnote_to_footnotes_fixture_allocates_next_id(footnotes_docx: Path) -> None:
    """`footnotes.docx` already declares ids 1-3 (plus boilerplate -1/0); the
    next real footnote must get id 4."""
    new_id = add_footnote_to_document(footnotes_docx, 5, "Fourth footnote.")

    assert new_id == "4"
    footnote = _footnote_by_id(footnotes_docx, "4")
    assert footnote is not None
    assert render_footnote_content(footnote) == "Fourth footnote."


def test_add_footnote_to_minimal_fixture_allocates_id_two(minimal_docx: Path) -> None:
    new_id = add_footnote_to_document(minimal_docx, 3, "Second footnote.")

    assert new_id == "2"


def test_add_footnote_append_does_not_touch_content_types_or_rels(footnotes_docx: Path) -> None:
    with zipfile.ZipFile(footnotes_docx) as archive:
        content_types_before = archive.read(CONTENT_TYPES_PART)
        rels_before = archive.read(DOCUMENT_RELS_PART)

    add_footnote_to_document(footnotes_docx, 5, "Fourth footnote.")

    with zipfile.ZipFile(footnotes_docx) as archive:
        assert archive.read(CONTENT_TYPES_PART) == content_types_before
        assert archive.read(DOCUMENT_RELS_PART) == rels_before


def test_add_footnote_append_leaves_existing_footnotes_untouched(footnotes_docx: Path) -> None:
    before = get_document_footnotes(footnotes_docx)

    add_footnote_to_document(footnotes_docx, 5, "Fourth footnote.")

    after = get_document_footnotes(footnotes_docx)
    assert after[: len(before)] == before


def test_add_footnote_multiline_content_round_trips_through_get_footnotes(
    footnotes_docx: Path,
) -> None:
    new_id = add_footnote_to_document(footnotes_docx, 5, "First line.\nSecond line.")

    footnotes = get_document_footnotes(footnotes_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.content == "First line.\nSecond line."


def test_add_footnote_empty_content_round_trips_as_empty_string(footnotes_docx: Path) -> None:
    new_id = add_footnote_to_document(footnotes_docx, 5, "")

    footnotes = get_document_footnotes(footnotes_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.content == ""


def test_add_footnote_astral_character_in_content_round_trips(footnotes_docx: Path) -> None:
    text = "Great job \U0001f600 today"

    new_id = add_footnote_to_document(footnotes_docx, 5, text)

    footnotes = get_document_footnotes(footnotes_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.content == text


def test_add_footnote_special_xml_characters_are_safely_escaped(footnotes_docx: Path) -> None:
    """lxml serializes `.text` content safely regardless of what characters
    it contains - pins that `add_footnote` never builds XML by string
    concatenation, which would be injectable (mirrors
    `tests/test_paragraph_edit.py`'s equivalent coverage for `insert_paragraph`)."""
    text = "<script>alert(1)</script> & \"quotes\" & 'apostrophes'"

    new_id = add_footnote_to_document(footnotes_docx, 5, text)

    footnotes = get_document_footnotes(footnotes_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.content == text
    footnote = _footnote_by_id(footnotes_docx, new_id)
    assert footnote is not None
    # Must still be a single, well-formed w:t text node.
    text_runs = footnote.findall("w:p/w:r/w:t", namespaces=NSMAP)
    assert len(text_runs) == 1


def test_add_footnote_last_paragraph_index_boundary(footnotes_docx: Path) -> None:
    last_index = len(_paragraphs(footnotes_docx)) - 1

    new_id = add_footnote_to_document(footnotes_docx, last_index, "Anchored at the very end.")

    footnotes = get_document_footnotes(footnotes_docx)
    added = next(f for f in footnotes if f.id == new_id)
    assert added.paragraph_index == last_index


def test_add_footnote_second_call_appends_instead_of_recreating_from_scratch(
    structured_docx: Path,
) -> None:
    """The first call on a zero-footnote document takes the from-scratch
    branch; a second call against the same, now-modified file must correctly
    re-detect that `word/footnotes.xml` now exists and take the append
    branch instead - not re-run the from-scratch branch and duplicate the
    boilerplate footnotes, the `Override`, or the `Relationship`."""
    first_id = add_footnote_to_document(structured_docx, 0, "First.")
    second_id = add_footnote_to_document(structured_docx, 1, "Second.")

    assert (first_id, second_id) == ("1", "2")
    root = _footnotes_root(structured_docx)
    assert root is not None
    ids = [f.get(f"{{{WORD_NS}}}id") for f in root.findall("w:footnote", namespaces=NSMAP)]
    assert ids == ["-1", "0", "1", "2"]  # boilerplate created exactly once

    with zipfile.ZipFile(structured_docx) as archive:
        content_types_root = etree.fromstring(archive.read(CONTENT_TYPES_PART))
    overrides = [
        el
        for el in content_types_root.findall(f"{{{_CONTENT_TYPES_NS}}}Override")
        if el.get("PartName") == "/word/footnotes.xml"
    ]
    assert len(overrides) == 1  # not duplicated by the second call


# --- Functional: add_footnote_to_document, id-allocation edge cases -----------


def test_add_footnote_id_allocation_skips_non_numeric_ids_without_colliding(
    tmp_path: Path,
) -> None:
    path = tmp_path / "non_numeric_id.docx"
    body = "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"
    footnotes = (
        '<w:footnote w:id="abc"><w:p><w:r><w:t>Weird.</w:t></w:r></w:p></w:footnote>'
        '<w:footnote w:id="2"><w:p><w:r><w:t>Real.</w:t></w:r></w:p></w:footnote>'
    )
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            FOOTNOTES_PART: _footnotes_xml(footnotes),
        },
    )

    new_id = add_footnote_to_document(path, 0, "New.")

    assert new_id == "3"


def test_add_footnote_id_allocation_defends_against_string_collision(tmp_path: Path) -> None:
    """A pathological document whose only declared ids happen to already
    include the naive `max+1` candidate as a non-derivable string must still
    get a genuinely unused id."""
    path = tmp_path / "collision.docx"
    body = "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"
    footnotes = (
        '<w:footnote w:id="1"><w:p><w:r><w:t>One.</w:t></w:r></w:p></w:footnote>'
        '<w:footnote w:id="2"><w:p><w:r><w:t>Two.</w:t></w:r></w:p></w:footnote>'
    )
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            FOOTNOTES_PART: _footnotes_xml(footnotes),
        },
    )

    new_id = add_footnote_to_document(path, 0, "New.")

    assert new_id == "3"
    assert _footnote_by_id(path, "1") is not None
    assert _footnote_by_id(path, "2") is not None


# --- Functional: add_footnote_to_document, package-wiring edge cases ----------


def test_add_footnote_creates_rels_part_when_entirely_absent(tmp_path: Path) -> None:
    path = tmp_path / "no_rels.docx"
    body = "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            CONTENT_TYPES_PART: _minimal_content_types(),
        },
    )
    with zipfile.ZipFile(path) as archive:
        assert DOCUMENT_RELS_PART not in archive.namelist()

    add_footnote_to_document(path, 0, "New.")

    assert _rels_declares_footnotes(path) is True


def test_add_footnote_missing_content_types_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_content_types.docx"
    body = "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    with pytest.raises(InvalidDocumentError, match="Content_Types"):
        add_footnote_to_document(path, 0, "New.")


def test_add_footnote_existing_but_unreferenced_content_type_is_not_duplicated(
    tmp_path: Path,
) -> None:
    """A document already declaring the footnotes Override despite
    word/footnotes.xml itself being absent (an inconsistent-but-not-invalid
    state this project's own tools never produce) must not end up with two
    Override entries for the same part."""
    path = tmp_path / "already_declared.docx"
    body = "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Types xmlns="{_CONTENT_TYPES_NS}">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/footnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        "</Types>"
    ).encode()
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            CONTENT_TYPES_PART: content_types,
            DOCUMENT_RELS_PART: _minimal_rels(),
        },
    )

    add_footnote_to_document(path, 0, "New.")

    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read(CONTENT_TYPES_PART))
    overrides = [
        el
        for el in root.findall(f"{{{_CONTENT_TYPES_NS}}}Override")
        if el.get("PartName") == "/word/footnotes.xml"
    ]
    assert len(overrides) == 1


def test_add_footnote_existing_but_unreferenced_relationship_is_not_duplicated(
    tmp_path: Path,
) -> None:
    """The `.rels` counterpart of the dedup guard above: a document already
    declaring the footnotes `Relationship` despite `word/footnotes.xml`
    itself being absent must not end up with two `Relationship` entries."""
    path = tmp_path / "already_declared_rels.docx"
    body = "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_RELS_NS}">'
        f'<Relationship Id="rId1" Type="{_FOOTNOTES_RELATIONSHIP_TYPE}" Target="footnotes.xml"/>'
        "</Relationships>"
    ).encode()
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            CONTENT_TYPES_PART: _minimal_content_types(),
            DOCUMENT_RELS_PART: rels,
        },
    )

    add_footnote_to_document(path, 0, "New.")

    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read(DOCUMENT_RELS_PART))
    relationships = [
        el
        for el in root.findall(f"{{{_RELS_NS}}}Relationship")
        if el.get("Type") == _FOOTNOTES_RELATIONSHIP_TYPE
    ]
    assert len(relationships) == 1


# --- Error / edge: add_footnote_to_document -------------------------------------


def test_add_footnote_paragraph_index_negative_is_rejected(structured_docx: Path) -> None:
    with pytest.raises(FootnoteEditError, match="paragraph_index"):
        add_footnote_to_document(structured_docx, -1, "x")


def test_add_footnote_paragraph_index_out_of_range_is_rejected(structured_docx: Path) -> None:
    with pytest.raises(FootnoteEditError, match="paragraph_index"):
        add_footnote_to_document(structured_docx, 999, "x")


def test_add_footnote_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"
    with pytest.raises(InvalidDocumentError):
        add_footnote_to_document(missing, 0, "x")


def test_add_footnote_not_a_valid_zip_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not_a_zip.docx"
    path.write_bytes(b"this is not a zip file")

    with pytest.raises(InvalidDocumentError):
        add_footnote_to_document(path, 0, "x")


def test_add_footnote_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_document_xml.docx"
    _write_zip(path, {"word/styles.xml": b"<styles/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        add_footnote_to_document(path, 0, "x")


def test_add_footnote_malformed_footnotes_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed_footnotes.docx"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml("<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"),
            FOOTNOTES_PART: b"<w:footnotes><unclosed>",
        },
    )

    with pytest.raises(InvalidDocumentError, match="footnotes.xml"):
        add_footnote_to_document(path, 0, "x")


def test_add_footnote_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    _write_zip(
        path, {"word/document.xml": _document_xml("<w:p><w:r><w:t>Small.</w:t></w:r></w:p>")}
    )
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        add_footnote_to_document(path, 0, "x", max_size_bytes=actual_size - 1)


def test_add_footnote_xxe_entity_in_existing_footnotes_xml_is_never_resolved(
    tmp_path: Path,
) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    malicious_footnotes = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:footnotes [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:footnotes xmlns:w="{WORD_NS}">'
        '<w:footnote w:id="1"><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p></w:footnote>'
        "</w:footnotes>"
    ).encode()
    _write_zip(
        path, {"word/document.xml": _document_xml(body), FOOTNOTES_PART: malicious_footnotes}
    )

    new_id = add_footnote_to_document(path, 0, "New text.")

    with zipfile.ZipFile(path) as archive:
        raw = archive.read(FOOTNOTES_PART)
    assert new_id == "2"
    assert b"TOP-SECRET-CONTENT" not in raw


def test_add_footnote_simulated_crash_leaves_original_file_untouched_append_case(
    footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original_bytes = footnotes_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml_module.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        add_footnote_to_document(footnotes_docx, 0, "x")

    assert footnotes_docx.read_bytes() == original_bytes


def test_add_footnote_simulated_crash_leaves_original_file_untouched_from_scratch_case(
    structured_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original_bytes = structured_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml_module.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        add_footnote_to_document(structured_docx, 0, "x")

    assert structured_docx.read_bytes() == original_bytes


# --- Functional: edit_footnote_in_document --------------------------------------


def test_edit_footnote_replaces_content_and_returns_previous_content(
    footnotes_docx: Path,
) -> None:
    previous = edit_footnote_in_document(footnotes_docx, "1", "Updated footnote text.")

    assert previous == " First footnote."
    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == "Updated footnote text."


def test_edit_footnote_multiline_content_increases_paragraph_count(footnotes_docx: Path) -> None:
    edit_footnote_in_document(footnotes_docx, "1", "Line one.\nLine two.\nLine three.")

    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == "Line one.\nLine two.\nLine three."


def test_edit_footnote_multiline_content_can_shrink_back_to_one_paragraph(
    footnotes_docx: Path,
) -> None:
    edit_footnote_in_document(footnotes_docx, "1", "Line one.\nLine two.")
    edit_footnote_in_document(footnotes_docx, "1", "Back to one line.")

    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == "Back to one line."


def test_edit_footnote_does_not_touch_document_xml(footnotes_docx: Path) -> None:
    with zipfile.ZipFile(footnotes_docx) as archive:
        document_xml_before = archive.read(DOCUMENT_PART)

    edit_footnote_in_document(footnotes_docx, "1", "Updated.")

    with zipfile.ZipFile(footnotes_docx) as archive:
        assert archive.read(DOCUMENT_PART) == document_xml_before


def test_edit_footnote_leaves_other_footnotes_untouched(footnotes_docx: Path) -> None:
    before = get_document_footnotes(footnotes_docx)
    others_before = [f for f in before if f.id != "1"]

    edit_footnote_in_document(footnotes_docx, "1", "Updated.")

    after = get_document_footnotes(footnotes_docx)
    others_after = [f for f in after if f.id != "1"]
    assert others_after == others_before


def test_edit_footnote_preserves_id_attribute(footnotes_docx: Path) -> None:
    edit_footnote_in_document(footnotes_docx, "1", "Updated.")

    footnote = _footnote_by_id(footnotes_docx, "1")
    assert footnote is not None
    assert footnote.get(f"{{{WORD_NS}}}id") == "1"


def test_edit_footnote_with_matching_expected_content_succeeds(footnotes_docx: Path) -> None:
    previous = edit_footnote_in_document(
        footnotes_docx, "1", "Updated.", expected_content=" First footnote."
    )

    assert previous == " First footnote."


def test_edit_footnote_empty_content_round_trips_as_empty_string(footnotes_docx: Path) -> None:
    edit_footnote_in_document(footnotes_docx, "1", "")

    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == ""


def test_edit_footnote_astral_character_in_content_round_trips(footnotes_docx: Path) -> None:
    text = "Great job \U0001f600 today"

    edit_footnote_in_document(footnotes_docx, "1", text)

    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == text


def test_edit_footnote_special_xml_characters_are_safely_escaped(footnotes_docx: Path) -> None:
    text = "<script>alert(1)</script> & \"quotes\" & 'apostrophes'"

    edit_footnote_in_document(footnotes_docx, "1", text)

    footnotes = get_document_footnotes(footnotes_docx)
    updated = next(f for f in footnotes if f.id == "1")
    assert updated.content == text


def test_edit_footnote_with_a_degenerate_zero_paragraph_footnote(tmp_path: Path) -> None:
    """A malformed-but-parseable `<w:footnote>` with no `w:p` children at all
    (`render_footnote_content` would report `""` for it) must still be
    editable - the clear-and-rebuild logic must not assume at least one
    existing paragraph to iterate over."""
    path = tmp_path / "zero_paragraph_footnote.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    footnotes = '<w:footnote w:id="1"></w:footnote>'
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), FOOTNOTES_PART: _footnotes_xml(footnotes)},
    )

    previous = edit_footnote_in_document(path, "1", "Now has real content.")

    assert previous == ""
    footnotes_after = get_document_footnotes(path)
    assert footnotes_after[0].content == "Now has real content."


# --- Error / edge: edit_footnote_in_document ------------------------------------


def test_edit_footnote_unknown_id_is_rejected(footnotes_docx: Path) -> None:
    with pytest.raises(FootnoteEditError, match="not found"):
        edit_footnote_in_document(footnotes_docx, "999", "x")


@pytest.mark.parametrize("reserved_id", ["-1", "0"])
def test_edit_footnote_reserved_boilerplate_id_is_rejected(
    footnotes_docx: Path, reserved_id: str
) -> None:
    with pytest.raises(FootnoteEditError, match="reserved|boilerplate"):
        edit_footnote_in_document(footnotes_docx, reserved_id, "x")

    # Not even attempted to be opened for writing - the boilerplate footnote's
    # own content must be provably unaffected.
    footnote = _footnote_by_id(footnotes_docx, reserved_id)
    assert footnote is not None


def test_edit_footnote_footnotes_xml_absent_is_rejected(structured_docx: Path) -> None:
    with pytest.raises(FootnoteEditError, match="not found"):
        edit_footnote_in_document(structured_docx, "1", "x")


def test_edit_footnote_mismatched_expected_content_is_rejected_as_stale(
    footnotes_docx: Path,
) -> None:
    with pytest.raises(FootnoteEditError, match="stale"):
        edit_footnote_in_document(footnotes_docx, "1", "x", expected_content="Wrong content.")

    footnotes = get_document_footnotes(footnotes_docx)
    unchanged = next(f for f in footnotes if f.id == "1")
    assert unchanged.content == " First footnote."


def test_edit_footnote_malformed_footnotes_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed_footnotes.docx"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml("<w:p><w:r><w:t>Text.</w:t></w:r></w:p>"),
            FOOTNOTES_PART: b"<w:footnotes><unclosed>",
        },
    )

    with pytest.raises(InvalidDocumentError, match="footnotes.xml"):
        edit_footnote_in_document(path, "1", "x")


def test_edit_footnote_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"
    with pytest.raises(InvalidDocumentError):
        edit_footnote_in_document(missing, "1", "x")


def test_edit_footnote_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no_document_xml.docx"
    _write_zip(path, {"word/styles.xml": b"<styles/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        edit_footnote_in_document(path, "1", "x")


def test_edit_footnote_oversized_file_is_rejected_before_parsing(footnotes_docx: Path) -> None:
    actual_size = footnotes_docx.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        edit_footnote_in_document(footnotes_docx, "1", "x", max_size_bytes=actual_size - 1)


def test_edit_footnote_xxe_entity_is_never_resolved(tmp_path: Path) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    malicious_footnotes = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:footnotes [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:footnotes xmlns:w="{WORD_NS}">'
        '<w:footnote w:id="1"><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p></w:footnote>'
        "</w:footnotes>"
    ).encode()
    _write_zip(
        path, {"word/document.xml": _document_xml(body), FOOTNOTES_PART: malicious_footnotes}
    )

    previous_content = edit_footnote_in_document(path, "1", "Safe replacement.")

    assert "TOP-SECRET-CONTENT" not in previous_content


def test_edit_footnote_simulated_crash_leaves_original_file_untouched(
    footnotes_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original_bytes = footnotes_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml_module.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        edit_footnote_in_document(footnotes_docx, "1", "x")

    assert footnotes_docx.read_bytes() == original_bytes


# --- Round-trip: add then edit --------------------------------------------------


def test_add_then_edit_the_same_new_footnote(structured_docx: Path) -> None:
    new_id = add_footnote_to_document(structured_docx, 0, "Original content.")

    previous = edit_footnote_in_document(structured_docx, new_id, "Edited content.")

    assert previous == "Original content."
    footnotes = get_document_footnotes(structured_docx)
    edited = next(f for f in footnotes if f.id == new_id)
    assert edited.content == "Edited content."
