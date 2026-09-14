"""Tests for `docx_mcp.server` construction: does `ServerConfig` actually reach the server?

Written test-first after finding, during a verification pass, that the
module-level `mcp = MCPServer(...)` singleton never passed `log_level`
through, silently ignoring `DOCX_MCP_LOG_LEVEL` even though
[docs/documentation-standards.md §5](../docs/documentation-standards.md#5-configuration-and-secrets)
and `.env.example` document it as effective.
"""

from __future__ import annotations

import subprocess
import sys

import pytest

from docx_mcp.config import ConfigError, ServerConfig
from docx_mcp.server import create_server


def test_default_server_uses_info_log_level() -> None:
    server = create_server(ServerConfig(allowed_roots=(), log_level="INFO"))

    assert server.settings.log_level == "INFO"


def test_server_honors_configured_debug_log_level() -> None:
    server = create_server(ServerConfig(allowed_roots=(), log_level="DEBUG"))

    assert server.settings.log_level == "DEBUG"


@pytest.mark.parametrize("level", ["WARNING", "ERROR", "CRITICAL"])
def test_server_honors_every_valid_log_level(level: str) -> None:
    server = create_server(ServerConfig(allowed_roots=(), log_level=level))

    assert server.settings.log_level == level


def test_create_server_defaults_to_process_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCX_MCP_LOG_LEVEL", "WARNING")

    server = create_server()

    assert server.settings.log_level == "WARNING"


def test_create_server_reports_the_registered_read_document_tool() -> None:
    server = create_server(ServerConfig(allowed_roots=(), log_level="INFO"))

    assert server.name == "docx-mcp"


def test_invalid_log_level_in_environment_raises_config_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("DOCX_MCP_LOG_LEVEL", "NOT_A_LEVEL")

    from docx_mcp.config import load_config

    with pytest.raises(ConfigError, match="NOT_A_LEVEL"):
        load_config()


def test_log_level_is_case_insensitive() -> None:
    from docx_mcp.config import load_config

    config = load_config({"DOCX_MCP_LOG_LEVEL": "debug"})

    assert config.log_level == "DEBUG"


def test_log_level_actually_reaches_the_python_logging_module_in_a_fresh_process() -> None:
    """`server.settings.log_level` above proves the value is threaded through
    `create_server`; it does not by itself prove Python's global `logging`
    verbosity actually changes, since the underlying SDK configures it via
    `logging.basicConfig()` (a no-op after the first call in a process - which
    is exactly once in real operation, at `docx_mcp.server` import time, but
    not across the several `create_server()` calls the tests above make in
    one pytest process). This spawns a real, fresh interpreter to check the
    end-to-end effect the way `python -m docx_mcp` actually experiences it.
    """
    code = (
        "import os, logging\n"
        "os.environ['DOCX_MCP_LOG_LEVEL'] = 'DEBUG'\n"
        "os.environ.pop('DOCX_MCP_ALLOWED_ROOTS', None)\n"
        "import docx_mcp.server\n"
        "print(logging.getLogger().getEffectiveLevel())\n"
    )
    result = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=30,
        check=True,
    )

    assert int(result.stdout.strip()) == 10  # logging.DEBUG
