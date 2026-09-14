"""Tests for `docx_mcp.config`.

See `docs/documentation-standards.md` §5 (Configuration and secrets) for the
naming convention and default values these tests pin down.
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from docx_mcp.config import DEFAULT_LOG_LEVEL, load_config


def test_empty_environment_denies_every_path() -> None:
    config = load_config({})
    assert config.allowed_roots == ()


def test_unset_allowed_roots_defaults_to_empty() -> None:
    config = load_config({"SOME_OTHER_VAR": "1"})
    assert config.allowed_roots == ()


def test_single_allowed_root_is_resolved_absolute(tmp_path: Path) -> None:
    root = tmp_path / "workspace"
    root.mkdir()
    config = load_config({"DOCX_MCP_ALLOWED_ROOTS": str(root)})
    assert config.allowed_roots == (root.resolve(),)


def test_multiple_allowed_roots_are_split_on_pathsep(tmp_path: Path) -> None:
    first = tmp_path / "a"
    second = tmp_path / "b"
    first.mkdir()
    second.mkdir()
    raw = os.pathsep.join([str(first), str(second)])
    config = load_config({"DOCX_MCP_ALLOWED_ROOTS": raw})
    assert config.allowed_roots == (first.resolve(), second.resolve())


def test_blank_segments_between_separators_are_ignored(tmp_path: Path) -> None:
    root = tmp_path / "workspace"
    root.mkdir()
    raw = f"{os.pathsep}{root}{os.pathsep}{os.pathsep}"
    config = load_config({"DOCX_MCP_ALLOWED_ROOTS": raw})
    assert config.allowed_roots == (root.resolve(),)


def test_log_level_defaults_to_info() -> None:
    config = load_config({})
    assert config.log_level == DEFAULT_LOG_LEVEL == "INFO"


def test_log_level_is_read_from_environment() -> None:
    config = load_config({"DOCX_MCP_LOG_LEVEL": "DEBUG"})
    assert config.log_level == "DEBUG"


def test_defaults_to_os_environ_when_no_mapping_given(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCX_MCP_LOG_LEVEL", "WARNING")
    config = load_config()
    assert config.log_level == "WARNING"
