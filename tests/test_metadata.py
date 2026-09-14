"""Tests for `docx_mcp.metadata.get_document_metadata`, from `specs/get_metadata.md`.

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
from docx_mcp.metadata import get_document_metadata

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_xml(body_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{body_inner_xml}</w:body></w:document>'
    ).encode()


def _core_xml(inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<cp:coreProperties xmlns:cp="{CORE_NS}" xmlns:dc="{DC_NS}" '
        f'xmlns:dcterms="{DCTERMS_NS}" xmlns:xsi="{XSI_NS}">{inner_xml}</cp:coreProperties>'
    ).encode()


# --- Functional: against the real fixtures -------------------------------------


def test_structured_fixture_core_properties(structured_docx: Path) -> None:
    result = get_document_metadata(structured_docx)

    assert result.title == "Structured Test Document"
    assert result.author == "Test Author"
    assert result.created == "2024-01-01T00:00:00Z"
    assert result.modified == "2024-06-15T12:30:00Z"


def test_structured_fixture_word_count_includes_table_cells(structured_docx: Path) -> None:
    result = get_document_metadata(structured_docx)

    # 19 words across paragraphs + 10 words across the 5 table cells (see
    # tests/fixtures/generate_fixtures.py's generate_structured_docx docstring).
    assert result.word_count == 29


def test_minimal_fixture_word_count(minimal_docx: Path) -> None:
    """ "Introduction"(1) + first paragraph(8) + footnote paragraph(6) +
    "Background"(1) + empty(0) + "Final paragraph."(2) = 18; no markers counted."""
    result = get_document_metadata(minimal_docx)

    assert result.word_count == 18


def test_minimal_fixture_has_no_title_but_has_python_docx_template_defaults(
    minimal_docx: Path,
) -> None:
    """`minimal.docx` never explicitly sets core properties (unlike `structured.docx`);
    `dc:title` is empty (-> `None`), but `dc:creator`/`dcterms:created`/`dcterms:modified`
    come from python-docx's own bundled default template, not from docx-mcp -
    a real (if incidental) "properties were never explicitly set by the author"
    fixture, distinct from `structured.docx`'s deliberately-set ones."""
    result = get_document_metadata(minimal_docx)

    assert result.title is None
    assert result.author == "python-docx"


# --- Core properties: missing / empty / malformed -------------------------------


def test_missing_docprops_core_xml_yields_null_properties_but_word_count_still_works(
    tmp_path: Path,
) -> None:
    path = tmp_path / "no_core.docx"
    body = "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_metadata(path)

    assert result.title is None
    assert result.author is None
    assert result.created is None
    assert result.modified is None
    assert result.word_count == 2


def test_malformed_docprops_core_xml_degrades_to_null_instead_of_raising(tmp_path: Path) -> None:
    """Unlike a malformed `word/styles.xml` in `get_structure` (a hard error), a
    broken *optional* `docProps/core.xml` must not fail the whole call - see
    specs/get_metadata.md §7 for the documented asymmetry."""
    path = tmp_path / "malformed_core.docx"
    body = "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "docProps/core.xml": b"<cp:coreProperties><unclosed>",
        },
    )

    result = get_document_metadata(path)

    assert result.title is None
    assert result.word_count == 2


def test_empty_title_and_creator_elements_are_null_not_empty_string(tmp_path: Path) -> None:
    path = tmp_path / "empty_props.docx"
    body = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    core = _core_xml("<dc:title/><dc:creator></dc:creator>")
    _write_zip(path, {"word/document.xml": _document_xml(body), "docProps/core.xml": core})

    result = get_document_metadata(path)

    assert result.title is None
    assert result.author is None


def test_created_and_modified_are_returned_verbatim_not_reformatted(tmp_path: Path) -> None:
    path = tmp_path / "dates.docx"
    body = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    core = _core_xml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">2020-05-01T09:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2021-11-30T23:59:59Z</dcterms:modified>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body), "docProps/core.xml": core})

    result = get_document_metadata(path)

    assert result.created == "2020-05-01T09:00:00Z"
    assert result.modified == "2021-11-30T23:59:59Z"


def test_unicode_and_astral_characters_in_core_properties_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "unicode_props.docx"
    body = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    core = _core_xml("<dc:title>Bericht \U0001f600</dc:title><dc:creator>Müller, José</dc:creator>")
    _write_zip(path, {"word/document.xml": _document_xml(body), "docProps/core.xml": core})

    result = get_document_metadata(path)

    assert result.title == "Bericht \U0001f600"
    assert result.author == "Müller, José"


def test_xml_predefined_entities_in_core_properties_decode_normally(tmp_path: Path) -> None:
    """`&amp;`/`&lt;`/`&gt;`/`&quot;` are ordinary XML escaping, unrelated to
    the XXE hardening tested separately - ampersands in a title/author
    ("Smith & Sons") are common and must decode correctly."""
    path = tmp_path / "escaped_props.docx"
    body = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    core = _core_xml('<dc:title>Smith &amp; Sons: &lt;Draft&gt; "v2"</dc:title>')
    _write_zip(path, {"word/document.xml": _document_xml(body), "docProps/core.xml": core})

    result = get_document_metadata(path)

    assert result.title == 'Smith & Sons: <Draft> "v2"'


def test_last_modified_by_is_not_surfaced(tmp_path: Path) -> None:
    """`cp:lastModifiedBy` is a documented non-goal (spec §3) - only dc:creator
    ("author") is surfaced, never confused with the last-editor property."""
    path = tmp_path / "last_modified_by.docx"
    body = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    core = _core_xml(
        "<dc:creator>Original Author</dc:creator>"
        "<cp:lastModifiedBy>Someone Else</cp:lastModifiedBy>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body), "docProps/core.xml": core})

    result = get_document_metadata(path)

    assert result.author == "Original Author"


# --- Word count: run-splitting and marker exclusion ------------------------------


def test_word_count_does_not_split_a_word_across_runs(tmp_path: Path) -> None:
    path = tmp_path / "run_split.docx"
    body = "<w:p><w:r><w:t>wor</w:t></w:r><w:r><w:t>d</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_metadata(path)

    assert result.word_count == 1


def test_word_count_joins_paragraphs_with_real_whitespace(tmp_path: Path) -> None:
    """Two adjacent paragraphs must not have their edge words fused together."""
    path = tmp_path / "two_paragraphs.docx"
    body = "<w:p><w:r><w:t>End</w:t></w:r></w:p><w:p><w:r><w:t>Start</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_metadata(path)

    assert result.word_count == 2


def test_word_count_excludes_heading_and_footnote_markers(tmp_path: Path) -> None:
    """word_count must reflect real content, not read_document's own '#'/'[^N]'
    rendering markers - a one-word heading with a footnote reference is 1 word."""
    path = tmp_path / "markers.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        "<w:r><w:t>Title</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_metadata(path)

    assert result.word_count == 1


def test_word_count_of_empty_document_is_zero(tmp_path: Path) -> None:
    path = tmp_path / "empty.docx"
    _write_zip(path, {"word/document.xml": _document_xml("")})

    result = get_document_metadata(path)

    assert result.word_count == 0


# --- Error / edge (shared validation, proven wired through metadata.py too) -----


def test_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"

    with pytest.raises(InvalidDocumentError, match="not found"):
        get_document_metadata(missing)


def test_non_zip_file_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not-a-zip.docx"
    path.write_bytes(b"this is not a zip archive")

    with pytest.raises(InvalidDocumentError, match="ZIP"):
        get_document_metadata(path)


def test_zip_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no-document-xml.docx"
    _write_zip(path, {"word/other.xml": b"<empty/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        get_document_metadata(path)


def test_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        get_document_metadata(path, max_size_bytes=actual_size - 1)


def test_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    result = get_document_metadata(path, max_size_bytes=actual_size)

    assert result.word_count == 2


# --- XXE / entity-expansion hardening for docProps/core.xml ---------------------


def test_external_entity_in_core_xml_is_not_resolved(tmp_path: Path) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe_core.docx"
    core_xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE cp:coreProperties [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<cp:coreProperties xmlns:cp="{CORE_NS}" xmlns:dc="{DC_NS}">'
        "<dc:title>&xxe;</dc:title></cp:coreProperties>"
    ).encode()
    body = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body), "docProps/core.xml": core_xml})

    result = get_document_metadata(path)

    assert "TOP-SECRET-CONTENT" not in (result.title or "")
