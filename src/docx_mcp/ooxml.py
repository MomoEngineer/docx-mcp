"""Generic OOXML/zip plumbing shared by every part-parsing module.

Split out of `document.py` in Phase 2 ([ADR-0003](../../docs/adr/0003-phase-2-module-layout.md)),
per the trigger [ADR-0001](../../docs/adr/0001-ooxml-library-and-module-layout.md)
named up front: a second read path (`structure.py`, `metadata.py`) needing the
same hardened-parsing/zip-opening logic `document.py`'s `extract_text` already
had. Deliberately has no knowledge of any WordprocessingML tag (`w:p`, `w:tbl`,
...) - that rendering logic stays in `document.py`, the one module that owns
"how a paragraph's text is assembled from its runs".

As of Phase 4 ([ADR-0005](../../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md)),
this module also owns `atomic_write_part` - the project's first write-path
plumbing, implementing the atomic-write contract from
[docs/security-model.md §3](../../docs/security-model.md#3-atomic-write-contract).
Like everything else here, it has no WordprocessingML knowledge: it rewrites
one named zip part's bytes, generically, for any future write tool
(`insert_paragraph`/`delete_paragraph` in Phase 5, `add_footnote`/`edit_footnote`
in Phase 6) to reuse rather than reimplement.

As of Phase 6 (ADR-0007, "Phase 6 Module Layout: `footnote_edit.py`, Multi-Part
Atomic Writes, and Anchor/Content Rules for `add_footnote`/`edit_footnote`"),
this module also owns `atomic_write_parts` - a more permissive sibling of
`atomic_write_part` (mirroring this module's existing `read_part`/
`read_optional_part` strict-vs-permissive split, now for writes) that writes
several named parts, any of which may be brand new to the archive, in one
atomic step. `atomic_write_part` itself, including its "the part must
already exist" validation, is unchanged; both functions share one private
helper (`_atomic_replace_archive`) for the actual temp-file-assemble-and-
replace sequence.
"""

from __future__ import annotations

import os
import tempfile
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


def _atomic_replace_archive(
    docx_path: Path,
    infos: list[zipfile.ZipInfo],
    contents: dict[str, bytes],
    overrides: dict[str, bytes],
) -> None:
    """Assemble a new zip archive in a temp file next to `docx_path`, then
    atomically replace it - the shared tail end of `atomic_write_part` and
    `atomic_write_parts`. `infos`/`contents` are the source archive's existing
    entries (in their original order); every name in `overrides` gets its
    given bytes instead of `contents`, whether or not that name was already
    in `infos` - a name in `overrides` but not in `infos` is appended as a
    brand new entry, in `overrides`' iteration order.

    Raises:
        OSError: If the temp file cannot be created or the final replace
            fails (e.g. a permissions error); the stray temp file is removed
            before the exception propagates.
    """
    existing_names = {info.filename for info in infos}
    new_names = [name for name in overrides if name not in existing_names]

    fd, tmp_name = tempfile.mkstemp(dir=docx_path.parent, prefix=".docx-mcp-", suffix=".tmp")
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out:
            for info in infos:
                data = overrides.get(info.filename, contents[info.filename])
                out.writestr(info, data)
            for name in new_names:
                out.writestr(name, overrides[name])
        os.replace(tmp_path, docx_path)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise


def atomic_write_part(
    docx_path: Path,
    part_name: str,
    new_content: bytes,
    *,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> None:
    """Rewrite `docx_path`, replacing `part_name`'s bytes with `new_content`, atomically.

    Implements [docs/security-model.md §3](../../docs/security-model.md#3-atomic-write-contract):
    the new archive is assembled in a temp file created in `docx_path.parent`
    (guaranteeing the final rename shares a filesystem/volume with the
    target), then atomically replaces the original via `os.replace`. The
    original is only ever touched by that single, final replace - a crash or
    interruption at any earlier point leaves it byte-for-byte intact, and the
    stray temp file is best-effort removed.

    Every part other than `part_name` is copied through with its
    **decompressed content and `zipfile.ZipInfo` metadata** (compression
    type, timestamp, external attributes) unchanged - not a claim of
    bit-identical *compressed* byte streams, which would need fragile
    low-level `zipfile` access; see
    [ADR-0005](../../docs/adr/0005-phase-4-text-edit-module-and-atomic-write.md)
    for why this is the guarantee actually made and tested.

    Args:
        docx_path: Path to the `.docx` file to rewrite. Callers must have
            already validated this path against the allowed roots (see
            `docx_mcp.security.resolve_safe_path`); this function does not
            perform any sandboxing itself.
        part_name: The internal part to replace (e.g. `word/document.xml`).
            Must already exist in the archive. Use `atomic_write_parts` if
            the part may not exist yet.
        new_content: The part's new raw bytes.
        max_size_bytes: Reject the current on-disk file if it exceeds this,
            before any temp file is created.

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, is not a valid ZIP archive, or does not already
            contain `part_name`.
        OSError: If the temp file cannot be created or the final replace
            fails (e.g. a permissions error); the stray temp file is removed
            before the exception propagates.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as source:
        read_part(source, part_name)  # validates part_name exists; raises otherwise
        infos = source.infolist()
        contents = {info.filename: source.read(info.filename) for info in infos}

    _atomic_replace_archive(docx_path, infos, contents, {part_name: new_content})


def atomic_write_parts(
    docx_path: Path,
    parts: dict[str, bytes],
    *,
    max_size_bytes: int = MAX_DOCX_SIZE_BYTES,
) -> None:
    """Rewrite `docx_path`, setting every part named in `parts` to its given bytes, atomically.

    A more permissive sibling of `atomic_write_part` - the same relationship
    `read_optional_part` has to `read_part` in this module, now for writes.
    Each name in `parts` is **replaced** if it already exists in the archive,
    or **added** as a brand new entry if it doesn't - added in `parts`'
    iteration order, after every existing entry. This is what `add_footnote`
    needs and `atomic_write_part` deliberately does not provide: adding
    `word/footnotes.xml` (plus `[Content_Types].xml`/`word/_rels/document.xml.rels`
    updates) together with the `word/document.xml` edit that references it,
    in one atomic step, so a crash between them can never happen (see
    [ADR-0007](../../docs/adr/0007-phase-6-footnote-edit-module-layout-and-multi-part-atomic-write.md)).
    `atomic_write_part`'s own behavior, including its "the part must already
    exist" validation, is completely unaffected by this function's existence.

    Every part not named in `parts` is copied through with its decompressed
    content and `zipfile.ZipInfo` metadata unchanged, using the same
    temp-file-in-`docx_path.parent`-then-atomic-`os.replace` sequence
    `atomic_write_part` uses (see
    [docs/security-model.md §3](../../docs/security-model.md#3-atomic-write-contract)).

    Args:
        docx_path: Path to the `.docx` file to rewrite. Callers must have
            already validated this path against the allowed roots; this
            function does not perform any sandboxing itself.
        parts: Mapping of internal part name to its new raw bytes. No name
            in `parts` is required to already exist in the archive.
        max_size_bytes: Reject the current on-disk file if it exceeds this,
            before any temp file is created.

    Raises:
        InvalidDocumentError: If the file does not exist, exceeds
            `max_size_bytes`, or is not a valid ZIP archive.
        OSError: If the temp file cannot be created or the final replace
            fails; the stray temp file is removed before the exception
            propagates.
    """
    with validate_and_open(docx_path, max_size_bytes=max_size_bytes) as source:
        infos = source.infolist()
        contents = {info.filename: source.read(info.filename) for info in infos}

    _atomic_replace_archive(docx_path, infos, contents, parts)
