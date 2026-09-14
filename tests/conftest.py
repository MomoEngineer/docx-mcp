"""Shared pytest fixtures.

Per [docs/testing.md §5](../docs/testing.md#5-principles-for-good-tests),
every test that touches a fixture `.docx` works on its own copy under
`tmp_path`, so no test can mutate `tests/fixtures/minimal.docx` on disk.
"""

from __future__ import annotations

import os
import shutil
from collections.abc import Iterator
from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"

# Sanitize ambient DOCX_MCP_* variables before any test module is imported:
# `docx_mcp.server` builds its module-level `mcp` singleton (which now reads
# DOCX_MCP_LOG_LEVEL - see tests/test_server.py) at import time, which happens
# during collection, before the `_clean_docx_mcp_env` fixture below would run.
os.environ.pop("DOCX_MCP_ALLOWED_ROOTS", None)
os.environ.pop("DOCX_MCP_LOG_LEVEL", None)


@pytest.fixture
def allowed_root(tmp_path: Path) -> Path:
    """An empty directory that tests treat as the single allowed root."""
    root = tmp_path / "workspace"
    root.mkdir()
    return root


@pytest.fixture
def minimal_docx(allowed_root: Path) -> Path:
    """A private copy of `tests/fixtures/minimal.docx`, inside `allowed_root`."""
    destination = allowed_root / "minimal.docx"
    shutil.copyfile(FIXTURES_DIR / "minimal.docx", destination)
    return destination


@pytest.fixture
def structured_docx(allowed_root: Path) -> Path:
    """A private copy of `tests/fixtures/structured.docx`, inside `allowed_root`."""
    destination = allowed_root / "structured.docx"
    shutil.copyfile(FIXTURES_DIR / "structured.docx", destination)
    return destination


@pytest.fixture
def footnotes_docx(allowed_root: Path) -> Path:
    """A private copy of `tests/fixtures/footnotes.docx`, inside `allowed_root`."""
    destination = allowed_root / "footnotes.docx"
    shutil.copyfile(FIXTURES_DIR / "footnotes.docx", destination)
    return destination


@pytest.fixture
def run_split_docx(allowed_root: Path) -> Path:
    """A private copy of `tests/fixtures/run_split.docx`, inside `allowed_root`."""
    destination = allowed_root / "run_split.docx"
    shutil.copyfile(FIXTURES_DIR / "run_split.docx", destination)
    return destination


@pytest.fixture
def outside_root(tmp_path: Path) -> Path:
    """A directory that is deliberately *not* an allowed root."""
    outside = tmp_path / "outside"
    outside.mkdir()
    return outside


@pytest.fixture(autouse=True)
def _clean_docx_mcp_env(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    """Ensure no ambient `DOCX_MCP_*` variable from the developer's shell leaks into a test."""
    monkeypatch.delenv("DOCX_MCP_ALLOWED_ROOTS", raising=False)
    monkeypatch.delenv("DOCX_MCP_LOG_LEVEL", raising=False)
    yield


@pytest.fixture
def anyio_backend() -> str:
    """Run `@pytest.mark.anyio` tests on asyncio only (trio is not a project dependency)."""
    return "asyncio"
