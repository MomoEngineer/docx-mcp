"""Path-sandboxing tests for `docx_mcp.security.resolve_safe_path`.

Required by [docs/security-model.md §5](../docs/security-model.md#5-testing-obligations)
and [CONTRIBUTING.md §5](../CONTRIBUTING.md#5-definition-of-done): every
filesystem-facing tool must reject a path that escapes its allowed root(s),
checked before any file is opened.
"""

from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

from docx_mcp.security import PathAccessError, resolve_safe_path

UNC_NETWORK_TIMEOUT_BUDGET_SECONDS = 1.0
"""A UNC path to an unreachable host was observed to block `Path.resolve()`
for ~8s doing real SMB/DNS resolution (see `security.py`'s module docstring).
The fix must reject it long before any such network attempt starts."""


def test_path_inside_allowed_root_is_accepted(allowed_root: Path) -> None:
    target = allowed_root / "doc.docx"
    target.write_bytes(b"")

    resolved = resolve_safe_path(str(target), (allowed_root,))

    assert resolved == target.resolve()


def test_relative_traversal_outside_root_is_rejected(allowed_root: Path) -> None:
    traversal = str(allowed_root / ".." / "secret.docx")

    with pytest.raises(PathAccessError):
        resolve_safe_path(traversal, (allowed_root,))


def test_relative_traversal_that_stays_inside_root_is_accepted(allowed_root: Path) -> None:
    nested = allowed_root / "nested"
    nested.mkdir()
    traversal = str(nested / ".." / "doc.docx")

    resolved = resolve_safe_path(traversal, (allowed_root,))

    assert resolved == (allowed_root / "doc.docx").resolve()


def test_absolute_path_outside_every_root_is_rejected(
    allowed_root: Path, outside_root: Path
) -> None:
    outside_file = outside_root / "secret.docx"

    with pytest.raises(PathAccessError):
        resolve_safe_path(str(outside_file), (allowed_root,))


def test_empty_allowed_roots_denies_every_path(allowed_root: Path) -> None:
    target = allowed_root / "doc.docx"

    with pytest.raises(PathAccessError):
        resolve_safe_path(str(target), ())


def test_empty_path_is_rejected(allowed_root: Path) -> None:
    with pytest.raises(PathAccessError):
        resolve_safe_path("", (allowed_root,))


def test_path_matching_one_of_several_roots_is_accepted(
    allowed_root: Path, outside_root: Path
) -> None:
    target = outside_root / "doc.docx"

    resolved = resolve_safe_path(str(target), (allowed_root, outside_root))

    assert resolved == target.resolve()


def test_root_itself_is_accepted(allowed_root: Path) -> None:
    resolved = resolve_safe_path(str(allowed_root), (allowed_root,))

    assert resolved == allowed_root.resolve()


def test_sibling_directory_sharing_a_name_prefix_is_rejected(tmp_path: Path) -> None:
    """`/allowed-root-evil` must not be accepted as inside `/allowed-root`."""
    root = tmp_path / "allowed-root"
    root.mkdir()
    sibling = tmp_path / "allowed-root-evil"
    sibling.mkdir()
    target = sibling / "doc.docx"

    with pytest.raises(PathAccessError):
        resolve_safe_path(str(target), (root,))


def test_symlink_escaping_the_allowed_root_is_rejected(
    allowed_root: Path, outside_root: Path
) -> None:
    secret = outside_root / "secret.docx"
    secret.write_bytes(b"")
    link = allowed_root / "link.docx"
    try:
        link.symlink_to(secret)
    except OSError:
        pytest.skip("creating symlinks is not permitted in this environment")

    with pytest.raises(PathAccessError):
        resolve_safe_path(str(link), (allowed_root,))


def test_symlinked_directory_escaping_the_allowed_root_is_rejected(
    allowed_root: Path, outside_root: Path
) -> None:
    (outside_root / "doc.docx").write_bytes(b"")
    link_dir = allowed_root / "escape"
    try:
        link_dir.symlink_to(outside_root, target_is_directory=True)
    except OSError:
        pytest.skip("creating symlinks is not permitted in this environment")

    with pytest.raises(PathAccessError):
        resolve_safe_path(str(link_dir / "doc.docx"), (allowed_root,))


def test_relative_input_path_is_resolved_against_cwd(
    allowed_root: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.chdir(allowed_root)
    (allowed_root / "doc.docx").write_bytes(b"")

    resolved = resolve_safe_path("doc.docx", (allowed_root,))

    assert resolved == (allowed_root / "doc.docx").resolve()


def test_nonexistent_path_inside_root_resolves_without_error(allowed_root: Path) -> None:
    """Resolution must not require the file to exist yet - existence is `document.py`'s job."""
    target = allowed_root / "does-not-exist.docx"

    resolved = resolve_safe_path(str(target), (allowed_root,))

    assert resolved == target.resolve()


def test_windows_drive_relative_traversal_is_rejected_or_kept_inside(allowed_root: Path) -> None:
    """A path using this OS's separator conventions never accidentally escapes."""
    traversal = str(allowed_root) + os.sep + ".." + os.sep + ".." + os.sep + "secret.docx"

    with pytest.raises(PathAccessError):
        resolve_safe_path(traversal, (allowed_root,))


# --- UNC network-path hardening ----------------------------------------------
#
# Found during a verification pass: on Windows, canonicalizing a UNC path
# (`\\host\share\...`) makes `Path.resolve()` perform a real SMB/DNS network
# round-trip to `host` - observed to block ~8s against an unreachable
# hostname - before the path sandboxing check even runs. That is a
# caller-controlled network access and denial-of-service vector, entirely
# apart from the path ending up rejected as outside the allowed roots. These
# tests pin down the fix: a UNC path is rejected immediately, before any
# resolution is attempted.


def test_unc_network_path_is_rejected_fast_not_after_a_network_timeout(
    allowed_root: Path,
) -> None:
    unreachable_unc_path = r"\\nonexistent-host-xyz-12345\share\file.docx"

    start = time.monotonic()
    with pytest.raises(PathAccessError, match="UNC"):
        resolve_safe_path(unreachable_unc_path, (allowed_root,))
    elapsed = time.monotonic() - start

    assert elapsed < UNC_NETWORK_TIMEOUT_BUDGET_SECONDS


def test_unc_network_path_with_forward_slashes_is_also_rejected_fast(
    allowed_root: Path,
) -> None:
    unreachable_unc_path = "//nonexistent-host-xyz-12345/share/file.docx"

    start = time.monotonic()
    with pytest.raises(PathAccessError, match="UNC"):
        resolve_safe_path(unreachable_unc_path, (allowed_root,))
    elapsed = time.monotonic() - start

    assert elapsed < UNC_NETWORK_TIMEOUT_BUDGET_SECONDS


def test_unc_path_to_a_reachable_host_is_still_rejected(allowed_root: Path) -> None:
    """Even a UNC path that *would* resolve immediately (loopback) must be
    rejected: the rule is "no UNC paths", not "no UNC paths that time out"."""
    with pytest.raises(PathAccessError, match="UNC"):
        resolve_safe_path(r"\\127.0.0.1\share\file.docx", (allowed_root,))


def test_extended_length_local_path_prefix_is_not_treated_as_unc(
    allowed_root: Path,
) -> None:
    """`\\\\?\\C:\\...` shares the leading double backslash but is a local
    path (the long-path-bypass prefix), not a network path - it must not be
    rejected by the UNC check."""
    target = allowed_root / "doc.docx"
    target.write_bytes(b"")
    extended_length_path = "\\\\?\\" + str(target)

    try:
        resolve_safe_path(extended_length_path, (allowed_root,))
    except PathAccessError as exc:
        assert "UNC" not in str(exc)


# --- Windows-specific path-comparison edge cases -----------------------------


@pytest.mark.skipif(os.name != "nt", reason="case-insensitive path semantics are Windows-only")
def test_differently_cased_path_still_matches_the_allowed_root(allowed_root: Path) -> None:
    target = allowed_root / "doc.docx"
    target.write_bytes(b"")
    uppercased = str(target).upper()

    resolved = resolve_safe_path(uppercased, (allowed_root,))

    assert resolved == target.resolve()


def test_allowed_root_with_trailing_separator_still_matches(allowed_root: Path) -> None:
    target = allowed_root / "doc.docx"
    target.write_bytes(b"")
    root_with_trailing_sep = Path(str(allowed_root) + os.sep)

    resolved = resolve_safe_path(str(target), (root_with_trailing_sep,))

    assert resolved == target.resolve()


def test_deeply_nested_path_inside_root_is_accepted(allowed_root: Path) -> None:
    nested = allowed_root
    for name in ("a", "b", "c", "d", "e"):
        nested = nested / name
        nested.mkdir()
    target = nested / "doc.docx"
    target.write_bytes(b"")

    resolved = resolve_safe_path(str(target), (allowed_root,))

    assert resolved == target.resolve()
