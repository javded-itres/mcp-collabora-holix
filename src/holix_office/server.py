"""stdio MCP server for workspace office documents."""

from __future__ import annotations

from holix_office import tools  # noqa: F401
from holix_office.mcp_app import mcp


def run() -> None:
    mcp.run(transport="stdio")
