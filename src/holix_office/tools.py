"""MCP tools for workspace office documents."""

from __future__ import annotations

import json
from typing import Any

from holix_office.edit import OfficeEditError
from holix_office.mcp_app import mcp
from holix_office.ops import (
    office_append,
    office_chart,
    office_format,
    office_list,
    office_media,
    office_read,
    office_replace,
    office_sheets,
    office_slides,
    office_table,
)


def _dumps(payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False, default=str)


def _call(fn, *args, **kwargs) -> str:
    try:
        return _dumps(fn(*args, **kwargs))
    except OfficeEditError as exc:
        return _dumps({"ok": False, "error": str(exc)})


@mcp.tool()
def office_list_tool() -> str:
    """List docx, xlsx, pptx, odt, ods, and odp files in the profile workspace."""
    return _call(office_list)


@mcp.tool()
def office_read_tool(path: str, max_chars: int = 40000) -> str:
    """Read plain text from a workspace office document.

    path: workspace-relative path, for example reports/plan.docx.
    """
    return _call(office_read, path, max_chars=max_chars)


@mcp.tool()
def office_replace_tool(path: str, old: str, new: str) -> str:
    """Replace text in a workspace office document and reload it in the editor.

    Contiguous text keeps its formatting. Text split across runs is rewritten
    as one run in that paragraph. path is workspace-relative.
    """
    return _call(office_replace, path, old, new)


@mcp.tool()
def office_format_tool(
    path: str,
    text: str,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str = "",
    fill: str = "",
    font: str = "",
    size: str = "",
    style: str = "",
) -> str:
    """Format existing text and reload the file in the editor.

    Works for docx, xlsx, pptx, odt, ods, and odp. path is workspace-relative.
    text is the exact phrase. font is a family such as Calibri. size is points,
    such as 14. color and fill are RRGGBB. style is a named paragraph style
    such as Heading 1, and is available for docx and odt. Omit a property to
    leave it unchanged. Bold, italic, color, font, and size apply to each
    match. fill shades a paragraph only when that paragraph's whole text is
    the phrase. Do not unzip the file or rewrite its XML.
    """
    return _call(
        office_format,
        path,
        text,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
        fill=fill,
        font=font,
        size=size,
        style=style,
    )


@mcp.tool()
def office_media_tool(
    path: str,
    media: str,
    slide: int = 1,
    width_cm: float = 12,
    height_cm: float = 6.8,
) -> str:
    """Insert a workspace image or video into an office document.

    path and media are workspace-relative. Images go into docx, xlsx, pptx,
    odt, ods, and odp. Video plays on pptx and odp slides and in an odt frame.
    A docx stores the video and a link. slide is 1-based for presentations and
    the sheet number for xlsx. width_cm and height_cm set the frame size.
    """
    return _call(
        office_media,
        path,
        media,
        slide=slide,
        width_cm=width_cm,
        height_cm=height_cm,
    )


@mcp.tool()
def office_slide_tool(
    path: str,
    action: str,
    slide: int = 1,
    title: str = "",
    body: str = "",
    layout: str = "title_body",
    notes: str = "",
) -> str:
    """Edit a pptx or odp presentation and reload it in the editor.

    action is list, add, set, notes, or delete. slide is 1-based. layout for
    a new pptx slide is title, title_body, section, two, title_only, or blank.
    Set title and body on add or set. notes replaces the speaker notes on pptx
    and adds a notes box on odp. Use office_format_tool for fonts on slide text,
    office_media_tool for pictures and video, office_table_tool and
    office_chart_tool for tables and charts.
    """
    return _call(
        office_slides,
        path,
        action,
        slide=slide,
        title=title,
        body=body,
        layout=layout,
        notes=notes,
    )


@mcp.tool()
def office_sheet_tool(
    path: str,
    action: str,
    sheet: str = "",
    cell: str = "",
    value: str = "",
    formula: str = "",
    link: str = "",
    range_ref: str = "",
) -> str:
    """Edit an xlsx or ods workbook: sheets, cells, formulas, and links.

    action is list, add, read, set, formula, link, or name. sheet is the sheet
    name. cell looks like A1. formula is Excel-like, including a link to another
    sheet such as =Other!A1. link is a hyperlink, for example https://example.com
    or #'Other'!A1. name creates an xlsx defined name; range_ref looks like
    Sheet1!$A$1:$B$10. list and read return values and formulas.
    """
    return _call(
        office_sheets,
        path,
        action,
        sheet=sheet,
        cell=cell,
        value=value,
        formula=formula,
        link=link,
        range_ref=range_ref,
    )


@mcp.tool()
def office_chart_tool(
    path: str,
    chart_type: str = "bar",
    sheet: str = "",
    data_range: str = "",
    categories_range: str = "",
    anchor: str = "E2",
    title: str = "",
    slide: int = 1,
    categories: str = "",
    series_name: str = "",
    values: str = "",
) -> str:
    """Add a bar, line, or pie chart to an xlsx sheet or a pptx slide.

    For xlsx, data_range is the values including the series title, for example
    B1:B4, and categories_range is the labels, for example A2:A4. anchor is the
    cell where the chart sits, for example E2. For pptx, categories and values
    are comma-separated lists of the same length, such as Jan,Feb and 10,12.
    """
    return _call(
        office_chart,
        path,
        chart_type=chart_type,
        sheet=sheet,
        data_range=data_range,
        categories_range=categories_range,
        anchor=anchor,
        title=title,
        slide=slide,
        categories=categories,
        series_name=series_name,
        values=values,
    )


@mcp.tool()
def office_table_tool(
    path: str,
    action: str,
    rows_json: str = "",
    table: int = 1,
    row: int = 0,
    col: int = 0,
    value: str = "",
    slide: int = 1,
    header: bool = True,
) -> str:
    """Read or write a table in docx, xlsx, pptx, odt, ods, or odp.

    action is read, add, or set. rows_json is a JSON list of rows, for example
    [["Name","Qty"],["A","1"]]. table is 1-based. row and col are 0-based when
    setting one cell. On xlsx, add writes the grid at A1 and creates an Excel
    table. On ods, add writes the grid onto the first sheet. slide is 1-based
    for pptx and odp.
    """
    return _call(
        office_table,
        path,
        action,
        rows_json=rows_json,
        table=table,
        row=row,
        col=col,
        value=value,
        slide=slide,
        header=header,
    )


@mcp.tool()
def office_append_tool(path: str, text: str) -> str:
    """Append a paragraph to a docx or odt file and reload it in the editor."""
    return _call(office_append, path, text)
