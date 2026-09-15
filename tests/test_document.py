"""Tests for `docx_mcp.document.extract_text`, from `specs/read_document.md`.

Covers the functional, error/edge, and XXE-hardening test categories required
by [docs/testing.md §2](../docs/testing.md#2-test-types-per-mcp-tool) and
[docs/security-model.md §5](../docs/security-model.md#5-testing-obligations).
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp.document import InvalidDocumentError, extract_text

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_xml(body_inner_xml: str, *, with_relationships_ns: bool = False) -> bytes:
    namespaces = f'xmlns:w="{WORD_NS}"'
    if with_relationships_ns:
        namespaces += f' xmlns:r="{REL_NS}"'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {namespaces}><w:body>{body_inner_xml}</w:body></w:document>"
    ).encode()


# --- Functional -------------------------------------------------------------


def test_minimal_fixture_produces_expected_text(minimal_docx: Path) -> None:
    expected = (
        "# Introduction\n"
        "This is the first paragraph of the document.\n"
        "This paragraph has a footnote reference.[^1]\n"
        "## Background\n"
        "\n"
        "Final paragraph."
    )

    assert extract_text(minimal_docx) == expected


def test_plain_paragraphs_without_headings_or_footnotes(tmp_path: Path) -> None:
    path = tmp_path / "plain.docx"
    body = "<w:p><w:r><w:t>First.</w:t></w:r></w:p><w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "First.\nSecond."


def test_heading_level_is_taken_from_style_id_digits(tmp_path: Path) -> None:
    path = tmp_path / "headings.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Deep heading</w:t></w:r></w:p>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "### Deep heading"


def test_non_heading_style_is_rendered_as_plain_text(tmp_path: Path) -> None:
    path = tmp_path / "styled.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>Not a heading.</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Not a heading."


def test_empty_paragraph_is_an_empty_line(tmp_path: Path) -> None:
    path = tmp_path / "empty_paragraph.docx"
    body = "<w:p/>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == ""


def test_multiple_runs_in_one_paragraph_are_concatenated_in_order(tmp_path: Path) -> None:
    path = tmp_path / "runs.docx"
    body = "<w:p><w:r><w:t>Hello, </w:t></w:r><w:r><w:t>world.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Hello, world."


def test_tab_is_rendered_as_literal_tab_character(tmp_path: Path) -> None:
    path = tmp_path / "tabs.docx"
    body = "<w:p><w:r><w:t>Name:</w:t><w:tab/><w:t>Value</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Name:\tValue"


def test_footnote_reference_is_rendered_inline_at_its_exact_position(tmp_path: Path) -> None:
    path = tmp_path / "footnote.docx"
    body = (
        "<w:p><w:r><w:t>Before</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="7"/></w:r>'
        "<w:r><w:t>After</w:t></w:r></w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Before[^7]After"


def test_tracked_change_deletion_is_excluded(tmp_path: Path) -> None:
    """`w:delText` (inside `w:del`) must never contribute to the extracted text."""
    path = tmp_path / "tracked_deletion.docx"
    body = (
        "<w:p>"
        "<w:r><w:t>Kept text.</w:t></w:r>"
        '<w:del w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:delText>Deleted text.</w:delText></w:r>"
        "</w:del>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Kept text."


def test_tracked_change_insertion_is_included(tmp_path: Path) -> None:
    """Ordinary `w:t` inside `w:ins` is current content and must be included."""
    path = tmp_path / "tracked_insertion.docx"
    body = (
        "<w:p>"
        "<w:r><w:t>Kept text. </w:t></w:r>"
        '<w:ins w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:t>Inserted text.</w:t></w:r>"
        "</w:ins>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Kept text. Inserted text."


def test_text_inside_a_hyperlink_is_included(tmp_path: Path) -> None:
    """`w:t` nested inside `w:hyperlink` is still a descendant of the paragraph."""
    path = tmp_path / "hyperlink.docx"
    body = '<w:p><w:hyperlink r:id="rId1"><w:r><w:t>Click here</w:t></w:r></w:hyperlink></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body, with_relationships_ns=True)})

    assert extract_text(path) == "Click here"


def test_field_instruction_text_is_excluded_but_cached_result_is_included(tmp_path: Path) -> None:
    """`w:instrText` (a field's code, e.g. a TOC field) must never appear; the
    cached display result after `w:fldChar type="separate"` is ordinary `w:t`
    and must appear, exactly like a Word-generated table of contents entry."""
    path = tmp_path / "field.docx"
    body = (
        "<w:p>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>Introduction</w:t></w:r>"
        "<w:r><w:tab/></w:r>"
        "<w:r><w:t>1</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = extract_text(path)

    assert "TOC" not in result
    assert result == "Introduction\t1"


def test_xml_space_preserve_whitespace_is_kept_verbatim(tmp_path: Path) -> None:
    path = tmp_path / "space.docx"
    body = '<w:p><w:r><w:t xml:space="preserve">  leading and trailing  </w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "  leading and trailing  "


def test_multiple_footnotes_in_the_same_paragraph_both_appear_in_order(tmp_path: Path) -> None:
    """The edge case Roadmap Phase 3 calls out for `get_footnotes` must already
    render correctly here, since `read_document` only places inline markers."""
    path = tmp_path / "multi_footnote.docx"
    body = (
        "<w:p>"
        "<w:r><w:t>A</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r>'
        "<w:r><w:t>B</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="2"/></w:r>'
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "A[^1]B[^2]"


def test_heading_and_footnote_in_the_same_paragraph_combine_correctly(tmp_path: Path) -> None:
    path = tmp_path / "heading_footnote.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        "<w:r><w:t>Title</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "# Title[^1]"


def test_heading_style_id_with_two_digit_level_is_not_misparsed(tmp_path: Path) -> None:
    """`Heading10` is not a real Word built-in style, but the digit-suffix rule
    must still behave predictably (ten '#' characters) rather than crash or
    silently truncate to a single digit."""
    path = tmp_path / "heading10.docx"
    body = '<w:p><w:pPr><w:pStyle w:val="Heading10"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "#" * 10 + " Deep"


def test_bold_and_italic_run_formatting_does_not_leak_into_text(tmp_path: Path) -> None:
    path = tmp_path / "formatting.docx"
    body = (
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>"
        "<w:r><w:t> plain </w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r>"
        "</w:p>"
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "Bold plain Italic"


def test_xml_predefined_entities_are_decoded_normally(tmp_path: Path) -> None:
    """`&amp;`/`&lt;`/`&gt;` are ordinary XML escaping, unrelated to the XXE
    hardening below - they must decode exactly like any other XML content."""
    path = tmp_path / "escaped.docx"
    body = '<w:p><w:r><w:t>A &amp; B &lt;tag&gt; "quoted"</w:t></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == 'A & B <tag> "quoted"'


def test_unicode_text_including_astral_characters_round_trips(tmp_path: Path) -> None:
    path = tmp_path / "unicode.docx"
    body = "<w:p><w:r><w:t>café \U0001f600</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    assert extract_text(path) == "café \U0001f600"


def test_multiple_paragraphs_with_mixed_headings_footnotes_and_tabs(tmp_path: Path) -> None:
    """A denser document than the fixture, closer to a real multi-section paper."""
    path = tmp_path / "dense.docx"
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Section 1</w:t></w:r></w:p>'
        "<w:p><w:r><w:t>Intro text</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Subsection</w:t></w:r></w:p>'
        "<w:p><w:r><w:t>Key:</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Value</w:t></w:r></w:p>"
        "<w:p/>"
        '<w:p><w:r><w:t>End</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>'
    )
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    expected = "# Section 1\nIntro text[^1]\n## Subsection\nKey:\tValue\n\nEnd[^2]"
    assert extract_text(path) == expected


# --- Error / edge cases -------------------------------------------------------


def test_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"

    with pytest.raises(InvalidDocumentError, match="not found"):
        extract_text(missing)


def test_directory_instead_of_file_is_rejected(tmp_path: Path) -> None:
    directory = tmp_path / "a-directory.docx"
    directory.mkdir()

    with pytest.raises(InvalidDocumentError):
        extract_text(directory)


def test_non_zip_file_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not-a-zip.docx"
    path.write_bytes(b"this is not a zip archive")

    with pytest.raises(InvalidDocumentError, match="ZIP"):
        extract_text(path)


def test_zip_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no-document-xml.docx"
    _write_zip(path, {"word/other.xml": b"<empty/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        extract_text(path)


def test_malformed_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed.docx"
    _write_zip(path, {"word/document.xml": b"<w:document><unclosed>"})

    with pytest.raises(InvalidDocumentError, match="malformed XML"):
        extract_text(path)


def test_odt_shaped_archive_is_rejected(tmp_path: Path) -> None:
    """An OpenDocument Text file is a ZIP too, but its main part is
    `content.xml`, not `word/document.xml` - it fails the same structural
    check as any other ZIP missing that part (see spec §4, Assumptions)."""
    path = tmp_path / "looks-like.odt"
    _write_zip(path, {"content.xml": b"<office:document-content/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        extract_text(path)


def test_dotx_extension_with_valid_wordprocessingml_is_accepted(minimal_docx: Path) -> None:
    """Validation is structural (spec §4): a `.dotx` template sharing the
    exact same `word/document.xml` shape as a `.docx` is accepted - this
    tool never inspects the file extension or `[Content_Types].xml`'s
    document-vs-template declaration."""
    dotx_path = minimal_docx.with_suffix(".dotx")
    dotx_path.write_bytes(minimal_docx.read_bytes())

    assert extract_text(dotx_path) == extract_text(minimal_docx)


def test_docm_extension_with_valid_wordprocessingml_is_accepted(minimal_docx: Path) -> None:
    """Validation is structural (spec §4): a `.docm` macro-enabled document
    sharing the exact same `word/document.xml` shape as a `.docx` is
    accepted - this tool never inspects the file extension or
    `[Content_Types].xml`'s document-vs-macro-enabled declaration."""
    docm_path = minimal_docx.with_suffix(".docm")
    docm_path.write_bytes(minimal_docx.read_bytes())

    assert extract_text(docm_path) == extract_text(minimal_docx)


def test_document_xml_without_body_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no-body.docx"
    xml = f'<w:document xmlns:w="{WORD_NS}"></w:document>'.encode()
    _write_zip(path, {"word/document.xml": xml})

    with pytest.raises(InvalidDocumentError, match="body"):
        extract_text(path)


def test_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        extract_text(path, max_size_bytes=actual_size - 1)


def test_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    assert extract_text(path, max_size_bytes=actual_size) == "Small content."


# --- XXE / entity-expansion hardening (docs/security-model.md §4) -----------


def test_external_entity_is_not_resolved_and_does_not_leak_secret_content(tmp_path: Path) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe.docx"
    xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:document [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>'
        "<w:p><w:r><w:t>&xxe;</w:t></w:r></w:p>"
        "</w:body></w:document>"
    ).encode()
    _write_zip(path, {"word/document.xml": xml})

    result = extract_text(path)

    assert "TOP-SECRET-CONTENT" not in result


def test_internal_entity_expansion_does_not_hang_or_crash(tmp_path: Path) -> None:
    """A small "billion laughs"-shaped internal entity chain must not be expanded."""
    path = tmp_path / "entity_expansion.docx"
    doctype = (
        "<!DOCTYPE w:document [\n"
        '<!ENTITY a "spam">\n'
        '<!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">\n'
        '<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">\n'
        "]>"
    )
    xml = (
        '<?xml version="1.0"?>\n'
        f"{doctype}\n"
        f'<w:document xmlns:w="{WORD_NS}"><w:body>'
        "<w:p><w:r><w:t>&c;</w:t></w:r></w:p>"
        "</w:body></w:document>"
    ).encode()
    _write_zip(path, {"word/document.xml": xml})

    result = extract_text(path)

    assert "spam" not in result
    assert len(result) < 1000
