"""Server configuration, read from environment variables.

See `docs/documentation-standards.md` §5 (Configuration and secrets) for the
naming convention and defaults.
"""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

ALLOWED_ROOTS_ENV_VAR = "DOCX_MCP_ALLOWED_ROOTS"
LOG_LEVEL_ENV_VAR = "DOCX_MCP_LOG_LEVEL"
DEFAULT_LOG_LEVEL = "INFO"

LogLevel = Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
VALID_LOG_LEVELS: tuple[LogLevel, ...] = ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL")


class ConfigError(Exception):
    """Raised when an environment variable holds a value that cannot be used."""


@dataclass(frozen=True)
class ServerConfig:
    """Resolved server configuration for one run.

    Attributes:
        allowed_roots: Canonicalized filesystem roots a `path` tool argument
            must resolve inside. Empty means every path is denied (see
            [docs/security-model.md §2](../../docs/security-model.md#2-path-sandboxing-rules)).
        log_level: The configured log level name (e.g. "INFO"), one of `VALID_LOG_LEVELS`.
    """

    allowed_roots: tuple[Path, ...]
    log_level: LogLevel


def _parse_log_level(raw: str) -> LogLevel:
    normalized = raw.strip().upper()
    if normalized not in VALID_LOG_LEVELS:
        valid = ", ".join(VALID_LOG_LEVELS)
        raise ConfigError(f"invalid {LOG_LEVEL_ENV_VAR}={raw!r}: must be one of {valid}")
    return normalized


def load_config(env: Mapping[str, str] | None = None) -> ServerConfig:
    """Build a `ServerConfig` from environment variables.

    Args:
        env: Mapping to read variables from. Defaults to `os.environ`; tests
            pass an explicit mapping instead of mutating process-wide state.

    Returns:
        A `ServerConfig` with canonicalized allowed roots and the configured
        log level (default `"INFO"` if unset; matched case-insensitively).

    Raises:
        ConfigError: If `DOCX_MCP_LOG_LEVEL` is set to something other than
            `DEBUG`, `INFO`, `WARNING`, `ERROR`, or `CRITICAL` (case-insensitive).
    """
    source = env if env is not None else os.environ
    raw_roots = source.get(ALLOWED_ROOTS_ENV_VAR, "")
    roots = tuple(Path(part).resolve() for part in raw_roots.split(os.pathsep) if part.strip())
    log_level = _parse_log_level(source.get(LOG_LEVEL_ENV_VAR, DEFAULT_LOG_LEVEL))
    return ServerConfig(allowed_roots=roots, log_level=log_level)
