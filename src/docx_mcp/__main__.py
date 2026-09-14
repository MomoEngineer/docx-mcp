"""Entry point: `python -m docx_mcp` or the `docx-mcp` console script.

Runs the server over `stdio` only, per
[CONTRIBUTING.md §1](../../CONTRIBUTING.md#1-guiding-principles) (no other
transport, no network access).
"""

from __future__ import annotations

from docx_mcp.server import mcp


def main() -> None:
    """Run the docx-mcp server over stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
