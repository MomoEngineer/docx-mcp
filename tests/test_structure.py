"""Tests for `docx_mcp.structure.get_document_structure`, from `specs/get_structure.md`.

Covers the functional, error/edge, and XXE-hardening test categories required by
[docs/testing.md §2](../docs/testing.md#2-test-types-per-mcp-tool) and
[docs/security-model.md §5](../docs/security-model.md#5-testing-obligations),
following the same synthetic-XML-fixture convention `tests/test_document.py`
established in Phase 1.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp.document import InvalidDocumentError
from docx_mcp.structure import get_document_structure

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_xml(body_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{body_inner_xml}</w:body></w:document>'
    ).encode()


def _styles_xml(styles_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{WORD_NS}">{styles_inner_xml}</w:styles>'
    ).encode()


# --- Functional: against the real fixtures -------------------------------------


def test_structured_fixture_paragraphs(structured_docx: Path) -> None:
    result = get_document_structure(structured_docx)

    actual = [(p.paragraph_index, p.text, p.style_id, p.heading_level) for p in result.paragraphs]
    assert actual == [
        (0, "Introduction", "Heading1", 1),
        (1, "This is a plain paragraph under the introduction.", None, None),
        (2, "Background", "Heading2", 2),
        (3, "Custom styled heading text", "MySectionHeading", 2),
        (4, "Deep Dive", "Heading3", 3),
        (5, "Final plain paragraph.", None, None),
    ]


def test_structured_fixture_toc_is_derived_from_headings(structured_docx: Path) -> None:
    result = get_document_structure(structured_docx)

    actual = [(t.paragraph_index, t.level, t.text) for t in result.toc]
    assert actual == [
        (0, 1, "Introduction"),
        (2, 2, "Background"),
        (3, 2, "Custom styled heading text"),
        (4, 3, "Deep Dive"),
    ]


def test_structured_fixture_table_with_horizontal_merge(structured_docx: Path) -> None:
    result = get_document_structure(structured_docx)

    assert len(result.tables) == 1
    table = result.tables[0]
    assert table.table_index == 0
    assert table.rows == (
        ("Header A", "Header B"),
        ("Value 1", "Value 2"),
        ("Merged Row",),
    )


def test_structured_fixture_has_no_footnotes(structured_docx: Path) -> None:
    result = get_document_structure(structured_docx)

    assert result.footnotes == ()


def test_minimal_fixture_footnote_anchor_index(minimal_docx: Path) -> None:
    """Reuses Phase 1's `minimal.docx` (one footnote reference at a known paragraph)."""
    result = get_document_structure(minimal_docx)

    assert [(f.id, f.paragraph_index) for f in result.footnotes] == [("1", 2)]


def test_minimal_fixture_has_no_tables(minimal_docx: Path) -> None:
    result = get_document_structure(minimal_docx)

    assert result.tables == ()


# --- Heading resolution: direct outlineLvl --------------------------------------


def test_heading_level_via_direct_outline_lvl_on_paragraph(tmp_path: Path) -> None:
    path = tmp_path / "direct.docx"
    body = '<w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 3


def test_outline_lvl_nine_means_body_text_not_a_heading(tmp_path: Path) -> None:
    path = tmp_path / "body_text.docx"
    body = '<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Not a heading</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level is None


@pytest.mark.parametrize("raw_val", ["-1", "10", "100", "abc", ""])
def test_out_of_range_or_malformed_direct_outline_lvl_is_not_a_heading(
    tmp_path: Path, raw_val: str
) -> None:
    """A `w:outlineLvl` value outside OOXML's valid `0..9` range, or not an
    integer at all, must never produce a nonsensical `heading_level` (e.g. `0`
    from `val="-1"`, or `101` from `val="100"`) - it is treated as if the
    element were absent (falls through, here to "not a heading" since there
    is no pStyle to fall back to)."""
    path = tmp_path / "bad_outline.docx"
    body = f'<w:p><w:pPr><w:outlineLvl w:val="{raw_val}"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level is None


def test_malformed_direct_outline_lvl_falls_through_to_digit_suffix_heuristic(
    tmp_path: Path,
) -> None:
    """A malformed direct override must fall through to the next resolution
    steps (spec §3), not just resolve to null - here, all the way to the
    Phase 1 `HeadingN` heuristic since there is no `styles.xml`."""
    path = tmp_path / "bad_outline_with_heading_style.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="-1"/></w:pPr>'
        "<w:r><w:t>Still a heading</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 2


def test_malformed_style_outline_lvl_falls_through_to_based_on_ancestor(
    tmp_path: Path,
) -> None:
    """A style whose own `w:outlineLvl` is malformed must not stop the
    `w:basedOn` chain walk - it continues to the ancestor, same as if that
    style had declared no `w:outlineLvl` at all."""
    path = tmp_path / "bad_style_outline.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="ChildStyle"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>'
    styles = (
        '<w:style w:styleId="ParentStyle"><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>'
        '<w:style w:styleId="ChildStyle">'
        '<w:basedOn w:val="ParentStyle"/><w:pPr><w:outlineLvl w:val="oops"/></w:pPr>'
        "</w:style>"
    )
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), "word/styles.xml": _styles_xml(styles)},
    )

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 3


def test_direct_outline_lvl_overrides_heading_style(tmp_path: Path) -> None:
    """A Heading1-styled paragraph with a direct outlineLvl=9 override is not a
    heading - direct/local formatting wins over the style (spec §3, step 1)."""
    path = tmp_path / "override.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/></w:pPr>'
        "<w:r><w:t>Looks like a heading</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level is None
    assert result.paragraphs[0].style_id == "Heading1"


# --- Heading resolution: styles.xml ---------------------------------------------


def test_heading_level_resolved_via_style_own_outline_lvl(tmp_path: Path) -> None:
    path = tmp_path / "style_direct.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="MyHeading"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>'
    styles = '<w:style w:styleId="MyHeading"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), "word/styles.xml": _styles_xml(styles)},
    )

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 2


def test_heading_level_resolved_via_based_on_chain(tmp_path: Path) -> None:
    path = tmp_path / "based_on.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="ChildStyle"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>'
    styles = (
        '<w:style w:styleId="ParentStyle"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
        '<w:style w:styleId="ChildStyle"><w:basedOn w:val="ParentStyle"/></w:style>'
    )
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), "word/styles.xml": _styles_xml(styles)},
    )

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 1


def test_circular_based_on_chain_does_not_hang_or_crash(tmp_path: Path) -> None:
    path = tmp_path / "circular.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
    styles = (
        '<w:style w:styleId="A"><w:basedOn w:val="B"/></w:style>'
        '<w:style w:styleId="B"><w:basedOn w:val="A"/></w:style>'
    )
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), "word/styles.xml": _styles_xml(styles)},
    )

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level is None


def test_fallback_to_digit_suffix_when_style_not_found_in_styles_xml(tmp_path: Path) -> None:
    path = tmp_path / "fallback.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>'
    styles = (
        '<w:style w:styleId="SomeOtherStyle"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
    )
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), "word/styles.xml": _styles_xml(styles)},
    )

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 3


def test_missing_styles_xml_falls_back_to_digit_suffix_heuristic(tmp_path: Path) -> None:
    path = tmp_path / "no_styles.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 2


def test_non_heading_style_without_outline_lvl_is_not_a_heading(tmp_path: Path) -> None:
    path = tmp_path / "quote.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>Not a heading.</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level is None


def test_malformed_styles_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed_styles.docx"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"),
            "word/styles.xml": b"<w:styles><unclosed>",
        },
    )

    with pytest.raises(InvalidDocumentError, match="styles.xml"):
        get_document_structure(path)


# --- Tables ----------------------------------------------------------------------


def test_table_cell_text_concatenates_multiple_paragraphs_with_newline(tmp_path: Path) -> None:
    path = tmp_path / "multi_para_cell.docx"
    body = (
        "<w:tbl><w:tr><w:tc>"
        "<w:p><w:r><w:t>First line</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Second line</w:t></w:r></w:p>"
        "</w:tc></w:tr></w:tbl>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.tables[0].rows == (("First line\nSecond line",),)


def test_vertically_merged_continuation_cell_renders_as_empty_string(tmp_path: Path) -> None:
    path = tmp_path / "vmerge.docx"
    body = (
        "<w:tbl>"
        '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
        "<w:p><w:r><w:t>Master</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>"
        "</w:tbl>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.tables[0].rows == (("Master",), ("",))


def test_footnote_reference_inside_table_cell_is_not_indexed_but_still_rendered(
    tmp_path: Path,
) -> None:
    """`paragraph_index`/`footnotes` only cover top-level body paragraphs (spec §5);
    the marker still appears inline in the cell's own text, like `read_document`."""
    path = tmp_path / "footnote_in_table.docx"
    body = (
        "<w:p><w:r><w:t>Body paragraph</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:tc></w:tr></w:tbl>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.footnotes == ()
    assert result.tables[0].rows == (("Cell[^1]",),)


def test_nested_table_is_not_extracted(tmp_path: Path) -> None:
    path = tmp_path / "nested_table.docx"
    body = (
        "<w:tbl><w:tr><w:tc>"
        "<w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        "</w:tc></w:tr></w:tbl>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert len(result.tables) == 1
    assert result.tables[0].rows == (("Outer cell",),)


def test_multiple_tables_are_indexed_independently_of_paragraphs(tmp_path: Path) -> None:
    path = tmp_path / "multi_table.docx"
    body = (
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>T0</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        "<w:p><w:r><w:t>Between</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>T1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert [t.table_index for t in result.tables] == [0, 1]
    assert result.tables[0].rows == (("T0",),)
    assert result.tables[1].rows == (("T1",),)
    # paragraph_index is a separate 0-based sequence over w:p only - unaffected
    # by the interleaved tables (spec §5).
    assert [p.paragraph_index for p in result.paragraphs] == [0, 1]


def test_table_with_zero_rows_is_an_empty_rows_tuple(tmp_path: Path) -> None:
    path = tmp_path / "empty_table.docx"
    body = "<w:tbl><w:tblPr/></w:tbl>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert len(result.tables) == 1
    assert result.tables[0].rows == ()


def test_unicode_and_astral_characters_in_headings_and_table_cells(tmp_path: Path) -> None:
    path = tmp_path / "unicode.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        "<w:r><w:t>Bericht \U0001f600</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>café</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_structure(path)

    assert result.paragraphs[0].text == "Bericht \U0001f600"
    assert result.toc[0].text == "Bericht \U0001f600"
    assert result.tables[0].rows == (("café",),)


def test_deep_non_circular_based_on_chain_resolves_correctly(tmp_path: Path) -> None:
    """A long but acyclic `w:basedOn` chain (unlike the malformed-circular
    case) must resolve normally, all the way to the ancestor that declares
    `outlineLvl` - proving the 64-hop guard doesn't reject legitimate chains
    well under that limit."""
    path = tmp_path / "deep_chain.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="Level5"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>'
    styles = (
        '<w:style w:styleId="Level1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
        '<w:style w:styleId="Level2"><w:basedOn w:val="Level1"/></w:style>'
        '<w:style w:styleId="Level3"><w:basedOn w:val="Level2"/></w:style>'
        '<w:style w:styleId="Level4"><w:basedOn w:val="Level3"/></w:style>'
        '<w:style w:styleId="Level5"><w:basedOn w:val="Level4"/></w:style>'
    )
    _write_zip(
        path,
        {"word/document.xml": _document_xml(body), "word/styles.xml": _styles_xml(styles)},
    )

    result = get_document_structure(path)

    assert result.paragraphs[0].heading_level == 1


def test_empty_document_has_empty_structure(tmp_path: Path) -> None:
    path = tmp_path / "empty.docx"
    _write_zip(path, {"word/document.xml": _document_xml("")})

    result = get_document_structure(path)

    assert result.paragraphs == ()
    assert result.toc == ()
    assert result.tables == ()
    assert result.footnotes == ()


# --- Error / edge (shared validation, proven wired through structure.py too) ----


def test_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"

    with pytest.raises(InvalidDocumentError, match="not found"):
        get_document_structure(missing)


def test_non_zip_file_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not-a-zip.docx"
    path.write_bytes(b"this is not a zip archive")

    with pytest.raises(InvalidDocumentError, match="ZIP"):
        get_document_structure(path)


def test_zip_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no-document-xml.docx"
    _write_zip(path, {"word/other.xml": b"<empty/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        get_document_structure(path)


def test_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        get_document_structure(path, max_size_bytes=actual_size - 1)


def test_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    result = get_document_structure(path, max_size_bytes=actual_size)

    assert result.paragraphs[0].text == "Small content."


# --- XXE / entity-expansion hardening for styles.xml (security-model.md §5) -----


def test_external_entity_in_styles_xml_is_not_resolved(tmp_path: Path) -> None:
    """`structure.py` never reads a style's `w:name` text into `DocumentStructure`
    at all, so a "secret not in the result" assertion alone wouldn't discriminate
    a hardened parser from an unhardened one here (unlike `word/document.xml`,
    whose `w:t` content does flow straight into the result). What *is*
    meaningful: parsing must still succeed (no crash/hang on the DOCTYPE) and
    heading resolution for the rest of the document must still be correct -
    proving the parse completed normally despite the external-entity
    declaration, without ever needing to dereference it."""
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe_styles.docx"
    styles_xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:styles [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:styles xmlns:w="{WORD_NS}">'
        '<w:style w:styleId="Heading1"><w:name>&xxe;</w:name></w:style>'
        '<w:style w:styleId="Heading2"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
        "</w:styles>"
    ).encode()
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body), "word/styles.xml": styles_xml})

    result = get_document_structure(path)

    assert "TOP-SECRET-CONTENT" not in str(result)
    # Heading1 has no own outlineLvl -> falls back to the digit-suffix heuristic;
    # Heading2's own outlineLvl resolves normally - both prove the parse
    # completed and styles.xml was fully, correctly walked.
    assert result.paragraphs[0].heading_level == 1
    assert result.paragraphs[1].heading_level == 2
