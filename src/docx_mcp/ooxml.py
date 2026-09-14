"""Generic OOXML/zip plumbing shared by every part-parsing module.

Split out of `document.py` in Phase 2 ([ADR-0003](../../docs/adr/0003-phase-2-module-layout.md)),
per the trigger [ADR-0001](../../docs/adr/0001-ooxml-library-and-module-layout.md)
named up front: a second read path (`structure.py`, `metadata.py`) needing the
same hardened-parsing/zip-opening logic `document.py`'s `extract_text` already
had. Deliberately has no knowledge of any WordprocessingML tag (`w:p`, `w:tbl`,
...) - that rendering logic stays in `document.py`, the one module that owns
"how a paragraph's text is assembled from its runs".
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

MAX_DOCX_SIZE_BYTES = 50 * 1024 * 1024


class InvalidDocumentError(Exception):
    """Raised when a file is not a readable, valid `.docx` document, or a
    required/declared internal part is missing or malformed."""


def hardened_parser() -> etree.XMLParser:
    """Build an `lxml` parser that never resolves entities or touches the network.

    `resolve_entities=False` and `load_dtd=False` together mean a crafted
    external entity (XXE) in a `.docx`'s XML is never substituted and never
    read; `no_network=True` blocks any network-based entity/DTD fetch as
    defense in depth; `huge_tree=False` (the default) keeps libxml2's built-in
    entity-expansion ("billion laughs") limits active. See
    [ADR-0001](../../docs/adr/0001-ooxml-library-and-module-layout.md).
    """
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        huge_tree=False,
        dtd_validation=False,
        load_dtd=False,
    )


def validate_and_open(
    docx_path: Path, *, max_size_bytes: int = MAX_DOCX_SIZE_BYTES
) -> zipfile.ZipFile:
    """Validate `docx_path` (exists, within the size limit, a valid ZIP) and open it.

    Args:
        docx_path: Path to the `.docx` file. Callers must have already
            validated this path against the allowed roots (see
            `docx_mcp.security.resolve_safe_path`); this function does not
            perform any sandboxing itself.
        max_size_bytes: Reject the file if it is larger than this, before it
            is opened as a ZIP archive.

    Returns:
        An open `zipfile.ZipFile`. Callers are responsible for closing it
        (typically via a `with` statement).

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, or is not a valid ZIP archive.
    """
    if not docx_path.is_file():
        raise InvalidDocumentError(f"file not found: {docx_path}")

    size = docx_path.stat().st_size
    if size > max_size_bytes:
        raise InvalidDocumentError(
            f"file exceeds the maximum allowed size ({size} > {max_size_bytes} bytes)"
        )

    try:
        return zipfile.ZipFile(docx_path)
    except zipfile.BadZipFile as exc:
        raise InvalidDocumentError("not a valid .docx file: not a ZIP archive") from exc


def read_part(archive: zipfile.ZipFile, part_name: str) -> bytes:
    """Read a required internal part from an already-opened `.docx` archive.

    Raises:
        InvalidDocumentError: If `part_name` is not present in `archive`.
    """
    try:
        return archive.read(part_name)
    except KeyError as exc:
        raise InvalidDocumentError(f"not a valid .docx file: missing {part_name}") from exc


def read_optional_part(archive: zipfile.ZipFile, part_name: str) -> bytes | None:
    """Read an internal part from an already-opened `.docx` archive, or `None` if absent."""
    try:
        return archive.read(part_name)
    except KeyError:
        return None


def parse_xml(raw_xml: bytes, *, part_name: str) -> etree._Element:
    """Parse `raw_xml` with the hardened parser.

    Args:
        raw_xml: The raw bytes of one internal `.docx` part.
        part_name: The part's path within the archive (e.g. `word/document.xml`),
            used only to produce a clear error message.

    Raises:
        InvalidDocumentError: If `raw_xml` is not well-formed XML.
    """
    try:
        return etree.fromstring(raw_xml, parser=hardened_parser())
    except etree.XMLSyntaxError as exc:
        raise InvalidDocumentError(
            f"not a valid .docx file: malformed XML in {part_name}: {exc}"
        ) from exc
