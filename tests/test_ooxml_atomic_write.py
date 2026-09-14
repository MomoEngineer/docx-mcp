"""Tests for `docx_mcp.ooxml.atomic_write_part`, the project's first write-path plumbing.

Covers the atomic-write contract from
[docs/security-model.md §3](../docs/security-model.md#3-atomic-write-contract)
and its testing obligation in
[docs/security-model.md §5](../docs/security-model.md#5-testing-obligations):
a simulated-crash test proving the original file is left byte-for-byte
intact. See [ADR-0005](../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md).
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp import ooxml
from docx_mcp.ooxml import InvalidDocumentError, atomic_write_part


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


@pytest.fixture
def sample_docx(tmp_path: Path) -> Path:
    path = tmp_path / "sample.docx"
    _write_zip(
        path,
        {
            "word/document.xml": b"<original/>",
            "word/styles.xml": b"<styles/>",
            "[Content_Types].xml": b"<types/>",
        },
    )
    return path


def test_target_part_content_is_replaced(sample_docx: Path) -> None:
    atomic_write_part(sample_docx, "word/document.xml", b"<edited/>")

    with zipfile.ZipFile(sample_docx) as archive:
        assert archive.read("word/document.xml") == b"<edited/>"


def test_untouched_parts_keep_identical_content_and_metadata(sample_docx: Path) -> None:
    with zipfile.ZipFile(sample_docx) as archive:
        before_info = archive.getinfo("word/styles.xml")
        before = {
            "content": archive.read("word/styles.xml"),
            "compress_type": before_info.compress_type,
            "date_time": before_info.date_time,
            "external_attr": before_info.external_attr,
        }

    atomic_write_part(sample_docx, "word/document.xml", b"<edited/>")

    with zipfile.ZipFile(sample_docx) as archive:
        after_info = archive.getinfo("word/styles.xml")
        assert archive.read("word/styles.xml") == before["content"]
        assert after_info.compress_type == before["compress_type"]
        assert after_info.date_time == before["date_time"]
        assert after_info.external_attr == before["external_attr"]


def test_target_part_can_shrink_or_grow_in_size(sample_docx: Path) -> None:
    """The target part's new content need not match the original's length -
    `writestr` must recompute size/CRC for the new bytes, not reuse the
    original `ZipInfo`'s stale values."""
    atomic_write_part(sample_docx, "word/document.xml", b"<a/>")
    with zipfile.ZipFile(sample_docx) as archive:
        assert archive.read("word/document.xml") == b"<a/>"

    much_larger = b"<edited>" + b"x" * 10_000 + b"</edited>"
    atomic_write_part(sample_docx, "word/document.xml", much_larger)
    with zipfile.ZipFile(sample_docx) as archive:
        assert archive.read("word/document.xml") == much_larger
        assert archive.read("word/styles.xml") == b"<styles/>"


def test_zip_entry_order_is_preserved(sample_docx: Path) -> None:
    with zipfile.ZipFile(sample_docx) as archive:
        names_before = archive.namelist()

    atomic_write_part(sample_docx, "word/document.xml", b"<edited/>")

    with zipfile.ZipFile(sample_docx) as archive:
        assert archive.namelist() == names_before


def test_missing_part_is_rejected(sample_docx: Path) -> None:
    with pytest.raises(InvalidDocumentError, match="word/footnotes.xml"):
        atomic_write_part(sample_docx, "word/footnotes.xml", b"<footnotes/>")

    with zipfile.ZipFile(sample_docx) as archive:
        assert archive.read("word/document.xml") == b"<original/>"


def test_missing_file_is_rejected(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist.docx"
    with pytest.raises(InvalidDocumentError):
        atomic_write_part(missing, "word/document.xml", b"<edited/>")


def test_oversized_file_is_rejected_before_any_temp_file_is_created(sample_docx: Path) -> None:
    actual_size = sample_docx.stat().st_size
    with pytest.raises(InvalidDocumentError):
        atomic_write_part(
            sample_docx, "word/document.xml", b"<edited/>", max_size_bytes=actual_size - 1
        )

    leftover = [p for p in sample_docx.parent.iterdir() if p.name.startswith(".docx-mcp-")]
    assert leftover == []


def test_simulated_crash_before_replace_leaves_original_byte_for_byte_intact(
    sample_docx: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Monkeypatching `os.replace` to fail simulates a crash/interruption at
    the last possible moment - after the new archive is fully assembled in
    the temp file, but before the atomic swap. The original must be
    untouched, and the stray temp file must not linger."""
    original_bytes = sample_docx.read_bytes()

    def _boom(_src: str, _dst: str) -> None:
        raise OSError("simulated crash")

    monkeypatch.setattr(ooxml.os, "replace", _boom)

    with pytest.raises(OSError, match="simulated crash"):
        atomic_write_part(sample_docx, "word/document.xml", b"<edited/>")

    assert sample_docx.read_bytes() == original_bytes
    leftover = [p for p in sample_docx.parent.iterdir() if p.name.startswith(".docx-mcp-")]
    assert leftover == []
