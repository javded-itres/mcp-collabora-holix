---
name: office_documents
description: >-
  Read and edit workspace docx, xlsx, pptx, odt, ods, and odp files through
  the holix_office MCP. A successful edit reloads the open Studio editor.
tags:
  - holix-studio
  - mcp
  - office
  - documents
  - docx
  - xlsx
agents:
  - main
user-invocable: true
---

# Workspace office documents

Use this when the user asks to read or change a Word, Excel, PowerPoint, or
LibreOffice file in the workspace (docx, xlsx, pptx, odt, ods, odp).

The tools belong to MCP server `holix_office`. Paths are relative to
`HOLIX_OFFICE_WORKSPACE`. In Holix Studio that directory is the profile
workspace, and a successful edit reloads an open Collabora tab.

If the `office_*` tools are missing, office editing is off. Say so. Do not
round-trip the file through Collabora convert-to. Do not unzip a document or
rewrite `word/document.xml` in the shell: that skips the editor reload.

## Tools

| Goal | Tool |
|------|------|
| Find documents | `office_list_tool` |
| Read plain text | `office_read_tool(path)` |
| Change text | `office_replace_tool(path, old, new)` |
| Font, size, bold, italic, color, shading, named style | `office_format_tool(...)` |
| Insert an image or video | `office_media_tool(path, media, slide, width_cm, height_cm)` |
| Slides: list, add, set, notes, delete | `office_slide_tool(path, action, ...)` |
| Sheets, cells, formulas, links, defined names | `office_sheet_tool(path, action, ...)` |
| Bar, line, or pie chart | `office_chart_tool(...)` |
| Tables | `office_table_tool(path, action, rows_json, ...)` |
| Add a paragraph (docx, odt) | `office_append_tool(path, text)` |

`path` and `media` are workspace-relative, for example `reports/plan.docx`
and `media/cover.png`.

`office_replace_tool` keeps formatting when the old text sits in one run.
Text split across runs is rewritten as one run in that paragraph.

`office_format_tool` works on docx, xlsx, pptx, odt, ods, and odp. `font` is
a family such as `Calibri`. `size` is points, such as `18`. `color` and
`fill` are `RRGGBB`. `style` is a named paragraph style such as `Heading 1`
and applies to docx and odt. Pass only the properties to change. Bold,
italic, color, font, and size apply to the exact phrase. `fill` shades a
paragraph only when that paragraph's whole text is the phrase. On xlsx the
font applies to each matching cell. Example:
`office_format_tool(path, "Title", bold=true, font="Calibri", size="18", color="FFFFFF", fill="000000")`.

`office_media_tool` inserts png, jpg, gif, webp, bmp, and tiff into any of
the six types. mp4, webm, and mov play on a pptx or odp slide and in an odt
frame. A docx stores the video file and a link. `slide` is 1-based. For
xlsx it selects the sheet.

`office_slide_tool` edits pptx and odp. `action` is `list`, `add`, `set`,
`notes`, or `delete`. `layout` for a new pptx slide is `title`,
`title_body`, `section`, `two`, `title_only`, or `blank`.

`office_sheet_tool` edits xlsx and ods. `action` is `list`, `add`, `read`,
`set`, `formula`, `link`, or `name`. `formula` is Excel-like, including a
cross-sheet link such as `=Other!A1`. `link` is a hyperlink. `name` with
`range_ref` such as `Sheet1!$A$1:$B$10` creates an xlsx defined name; pass
that name in `sheet`. `read` returns cell values and formulas. Use this for
numbers, not only `office_read_tool`.

`office_chart_tool` adds a `bar`, `line`, or `pie` chart. On xlsx,
`data_range` is the series including its title (`B1:B4`) and
`categories_range` is the labels (`A2:A4`). `anchor` is the cell, such as
`E2`. On pptx, `categories` and `values` are comma-separated lists of the
same length (`Jan,Feb` and `10,12`). Charts are not written into ods or odp.

`office_table_tool` `action` is `read`, `add`, or `set`. `rows_json` is a
JSON list of rows, for example `[["Name","Qty"],["A","1"]]`. `table` is
1-based. `row` and `col` are 0-based for `set`. On xlsx, `add` writes the
grid at A1 and creates an Excel table. On ods, `add` writes the grid onto
the first sheet.

A reload shows the saved file. Keystrokes that were still unsaved in that
Collabora tab are dropped. The editor also has a Refresh document button for
a file that was changed outside these tools.
