"""Tests for `docx_mcp.footnotes.get_document_footnotes`, from `specs/get_footnotes.md`.

Covers the functional, error/edge, and XXE-hardening test categories required by
[docs/testing.md §2](../docs/testing.md#2-test-types-per-mcp-tool) and
[docs/security-model.md §5](../docs/security-model.md#5-testing-obligations),
following the same synthetic-XML-fixture convention `tests/test_document.py`/
`tests/test_structure.py` established in Phases 1-2 for the error/edge matrix,
plus the real `tests/fixtures/footnotes.docx` fixture (Phase 3) for the
functional happy path.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp.document import InvalidDocumentError
from docx_mcp.footnotes import get_document_footnotes

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


def _footnotes_xml(footnotes_inner_xml: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:footnotes xmlns:w="{WORD_NS}">{footnotes_inner_xml}</w:footnotes>'
    ).encode()


# --- Functional: against the real fixtures -------------------------------------


def test_footnotes_fixture_three_footnotes_two_sharing_a_paragraph(
    footnotes_docx: Path,
) -> None:
    """`footnotes.docx` (Phase 3) is built exactly for this: ids 1 and 2 both
    anchored in paragraph 2, id 3 anchored separately in paragraph 4 - the
    edge case Roadmap Phase 3 names as most likely to break a naive
    anchor-to-paragraph mapping."""
    result = get_document_footnotes(footnotes_docx)

    assert [(f.id, f.paragraph_index, f.content) for f in result] == [
        ("1", 2, " First footnote."),
        ("2", 2, " Second footnote."),
        ("3", 4, " Third footnote."),
    ]


def test_minimal_fixture_single_footnote(minimal_docx: Path) -> None:
    """Reuses Phase 1's `minimal.docx` (one footnote reference at a known paragraph).
    Content carries the leading space Word writes after the auto-number,
    preserved verbatim (xml:space="preserve") rather than trimmed."""
    result = get_document_footnotes(minimal_docx)

    assert [(f.id, f.paragraph_index, f.content) for f in result] == [
        ("1", 2, " This is a footnote."),
    ]


# --- Functional: content rendering (synthetic) ----------------------------------


def test_multi_paragraph_footnote_content_is_joined_by_newline(tmp_path: Path) -> None:
    path = tmp_path / "multi_para_footnote.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    footnotes = (
        '<w:footnote w:id="1">'
        "<w:p><w:r><w:t>First line.</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Second line.</w:t></w:r></w:p>"
        "</w:footnote>"
    )
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "word/footnotes.xml": _footnotes_xml(footnotes),
        },
    )

    result = get_document_footnotes(path)

    assert result[0].content == "First line.\nSecond line."


def test_footnote_content_uses_plain_text_without_heading_or_footnote_markers(
    tmp_path: Path,
) -> None:
    """`content` uses `paragraph_plain_text`, not `render_paragraph` - a
    footnote paragraph that happens to carry a Heading-prefixed `pStyle` must
    not leak a `"# "` marker into `content` (spec §3)."""
    path = tmp_path / "heading_style_footnote.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    footnotes = (
        '<w:footnote w:id="1">'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        "<w:r><w:t>Looks like a heading</w:t></w:r></w:p>"
        "</w:footnote>"
    )
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "word/footnotes.xml": _footnotes_xml(footnotes),
        },
    )

    result = get_document_footnotes(path)

    assert result[0].content == "Looks like a heading"


def test_boilerplate_separator_footnotes_never_appear_in_the_output(tmp_path: Path) -> None:
    """Word's `id="-1"`/`id="0"` separator/continuationSeparator footnotes are
    never referenced by a body `w:footnoteReference`, so the anchor-driven
    population (spec §3, ADR-0004) excludes them without any explicit
    `w:type` filtering."""
    path = tmp_path / "boilerplate.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    footnotes = (
        '<w:footnote w:type="separator" w:id="-1">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        '<w:footnote w:type="continuationSeparator" w:id="0">'
        "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
        '<w:footnote w:id="1"><w:p><w:r><w:t>Real footnote.</w:t></w:r></w:p></w:footnote>'
    )
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "word/footnotes.xml": _footnotes_xml(footnotes),
        },
    )

    result = get_document_footnotes(path)

    assert [f.id for f in result] == ["1"]


# --- Functional: the central multi-anchor edge case (synthetic, unit-level) -----


def test_two_footnotes_anchored_in_the_same_paragraph_are_two_separate_entries(
    tmp_path: Path,
) -> None:
    """The output is a flat list, not a mapping keyed by `paragraph_index` -
    two different ids anchored in one paragraph must both survive, in order,
    never collapsing into a single entry (spec §3)."""
    path = tmp_path / "same_paragraph.docx"
    body = (
        "<w:p>"
        "<w:r><w:t>A</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r>'
        "<w:r><w:t>B</w:t></w:r>"
        '<w:r><w:footnoteReference w:id="2"/></w:r>'
        "</w:p>"
    )
    footnotes = (
        '<w:footnote w:id="1"><w:p><w:r><w:t>Footnote one.</w:t></w:r></w:p></w:footnote>'
        '<w:footnote w:id="2"><w:p><w:r><w:t>Footnote two.</w:t></w:r></w:p></w:footnote>'
    )
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "word/footnotes.xml": _footnotes_xml(footnotes),
        },
    )

    result = get_document_footnotes(path)

    assert [(f.id, f.paragraph_index, f.content) for f in result] == [
        ("1", 0, "Footnote one."),
        ("2", 0, "Footnote two."),
    ]


# --- Error / edge: footnotes.xml-specific ---------------------------------------


def test_missing_footnotes_xml_with_no_anchors_is_an_empty_tuple(tmp_path: Path) -> None:
    path = tmp_path / "no_footnotes.docx"
    body = "<w:p><w:r><w:t>No footnotes here.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_footnotes(path)

    assert result == ()


def test_missing_footnotes_xml_with_a_dangling_anchor_gives_null_content(tmp_path: Path) -> None:
    """A body anchor with no `word/footnotes.xml` part at all (a malformed/
    inconsistent document) is not an error - the entry still appears, with
    `content: None` (spec §7)."""
    path = tmp_path / "dangling_no_part.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    _write_zip(path, {"word/document.xml": _document_xml(body)})

    result = get_document_footnotes(path)

    assert [(f.id, f.paragraph_index, f.content) for f in result] == [("1", 0, None)]


def test_dangling_anchor_id_not_declared_in_valid_footnotes_xml_gives_null_content(
    tmp_path: Path,
) -> None:
    """A `word/footnotes.xml` that is present and well-formed but simply
    doesn't declare the referenced id degrades that one entry's `content` to
    `None` rather than failing the whole call (spec §7)."""
    path = tmp_path / "dangling_id.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="99"/></w:r></w:p>'
    footnotes = '<w:footnote w:id="1"><w:p><w:r><w:t>Unrelated.</w:t></w:r></w:p></w:footnote>'
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml(body),
            "word/footnotes.xml": _footnotes_xml(footnotes),
        },
    )

    result = get_document_footnotes(path)

    assert [(f.id, f.paragraph_index, f.content) for f in result] == [("99", 0, None)]


def test_malformed_footnotes_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed_footnotes.docx"
    _write_zip(
        path,
        {
            "word/document.xml": _document_xml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"),
            "word/footnotes.xml": b"<w:footnotes><unclosed>",
        },
    )

    with pytest.raises(InvalidDocumentError, match="footnotes.xml"):
        get_document_footnotes(path)


# --- Error / edge: shared validation (proven wired through footnotes.py too) ----


def test_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"

    with pytest.raises(InvalidDocumentError, match="not found"):
        get_document_footnotes(missing)


def test_non_zip_file_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "not-a-zip.docx"
    path.write_bytes(b"this is not a zip archive")

    with pytest.raises(InvalidDocumentError, match="ZIP"):
        get_document_footnotes(path)


def test_zip_missing_document_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "no-document-xml.docx"
    _write_zip(path, {"word/other.xml": b"<empty/>"})

    with pytest.raises(InvalidDocumentError, match="document.xml"):
        get_document_footnotes(path)


def test_oversized_file_is_rejected_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "oversized.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    with pytest.raises(InvalidDocumentError, match="maximum allowed size"):
        get_document_footnotes(path, max_size_bytes=actual_size - 1)


def test_file_at_exactly_the_size_limit_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "exact_size.docx"
    body = "<w:p><w:r><w:t>Small content.</w:t></w:r></w:p>"
    _write_zip(path, {"word/document.xml": _document_xml(body)})
    actual_size = path.stat().st_size

    result = get_document_footnotes(path, max_size_bytes=actual_size)

    assert result == ()


# --- XXE / entity-expansion hardening for footnotes.xml (security-model.md §5) --


def test_external_entity_in_footnotes_xml_is_not_resolved_and_does_not_leak_secret_content(
    tmp_path: Path,
) -> None:
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET-CONTENT")
    secret_uri = secret.resolve().as_uri()

    path = tmp_path / "xxe_footnotes.docx"
    body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    footnotes_xml = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE w:footnotes [<!ENTITY xxe SYSTEM "{secret_uri}">]>\n'
        f'<w:footnotes xmlns:w="{WORD_NS}">'
        '<w:footnote w:id="1"><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p></w:footnote>'
        "</w:footnotes>"
    ).encode()
    _write_zip(
        path, {"word/document.xml": _document_xml(body), "word/footnotes.xml": footnotes_xml}
    )

    result = get_document_footnotes(path)

    assert result[0].content is not None
    assert "TOP-SECRET-CONTENT" not in result[0].content
