"""Shared FastMCP app for workspace office documents."""

from __future__ import annotations

try:
    from mcp.server.fastmcp import FastMCP
except ModuleNotFoundError as exc:  # pragma: no cover - environment pin issue
    raise ModuleNotFoundError(
        "mcp.server.fastmcp is missing. Holix Studio requires the MCP Python SDK "
        "v1.x: pip install 'mcp>=1.2.0,<2'."
    ) from exc

mcp = FastMCP(
    name="mcp-collabora-holix",
    instructions=(
        "Read and edit office documents in the workspace directory: "
        "docx, xlsx, pptx, odt, ods, and odp. Use office_list_tool, "
        "office_read_tool, office_replace_tool, office_format_tool, "
        "office_media_tool, office_slide_tool, office_sheet_tool, "
        "office_chart_tool, office_table_tool, and office_append_tool "
        "(docx and odt paragraphs). Paths are relative to HOLIX_OFFICE_WORKSPACE. "
        "Do not unzip a document or rewrite its XML in the shell. "
        "When Holix Studio is hosting the file, a successful edit reloads the "
        "open editor and drops unsaved keystrokes in that tab."
    ),
)
