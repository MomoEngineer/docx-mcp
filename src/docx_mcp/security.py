"""Path sandboxing: every filesystem-facing tool resolves its path through here first.

Implements [docs/security-model.md §2](../../docs/security-model.md#2-path-sandboxing-rules):
a path is canonicalized (`..` segments and symlinks resolved) and checked
against an explicit allow-listed set of roots *before* any file is opened.
An empty allow-list denies every path.

A UNC path (`\\\\host\\share\\...`) is rejected *before* that canonicalization
step: on Windows, `Path.resolve()` canonicalizing such a path triggers a real
SMB/DNS network round-trip to `host` - confirmed to block for several seconds
against an unreachable host - which would violate the "local-first, no network
access anywhere in the server" principle
([CONTRIBUTING.md §1](../../CONTRIBUTING.md#1-guiding-principles), item 6) and
is a caller-controlled denial-of-service vector, entirely apart from whether
the path would end up inside an allowed root. See
[docs/security-model.md §1](../../docs/security-model.md#1-threat-model).
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path, PureWindowsPath


class PathAccessError(Exception):
    """Raised when a path cannot be resolved or escapes every allowed root."""


def _is_unc_network_path(raw_path: str) -> bool:
    """True for a `\\\\host\\share\\...` (or `//host/share/...`) UNC path.

    Deliberately excludes the `\\\\?\\` (extended-length local path) and
    `\\\\.\\` (device path) prefixes, which share the leading double
    backslash but never cause network resolution. Checked with
    `PureWindowsPath` (pure lexical parsing, no filesystem or network access)
    regardless of the host OS, so the same protective behavior is tested and
    enforced consistently everywhere, not only on Windows where the
    underlying `Path.resolve()` behavior actually triggers a network call.
    """
    drive = PureWindowsPath(raw_path).drive
    return drive.startswith("\\\\") and not drive.startswith(("\\\\?\\", "\\\\.\\"))


def resolve_safe_path(raw_path: str, allowed_roots: Sequence[Path]) -> Path:
    """Resolve `raw_path` and verify it falls inside one of `allowed_roots`.

    Args:
        raw_path: The path argument as supplied by the calling agent, absolute
            or relative.
        allowed_roots: Canonicalized roots the resolved path must fall inside
            (see `docx_mcp.config.load_config`).

    Returns:
        The canonicalized, absolute path, guaranteed to be inside at least one
        allowed root.

    Raises:
        PathAccessError: If `allowed_roots` is empty, `raw_path` is a UNC
            network path, `raw_path` cannot be resolved, or the resolved path
            does not fall inside any allowed root. Raised before any attempt
            to open the file - and, for a UNC path, before any attempt to
            even canonicalize it - per
            [docs/security-model.md §2](../../docs/security-model.md#2-path-sandboxing-rules).
    """
    if not allowed_roots:
        raise PathAccessError("access denied: no allowed roots are configured")

    if not raw_path or not raw_path.strip():
        raise PathAccessError("path must not be empty")

    if _is_unc_network_path(raw_path):
        raise PathAccessError("UNC network paths are not allowed")

    try:
        resolved = Path(raw_path).resolve(strict=False)
    except OSError as exc:
        raise PathAccessError(f"path could not be resolved: {exc}") from exc

    for root in allowed_roots:
        if resolved == root or root in resolved.parents:
            return resolved

    raise PathAccessError("path is outside the allowed roots")
