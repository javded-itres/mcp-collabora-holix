# Holix Office tool reference

The server name is `holix-office`. Every path is relative to `HOLIX_OFFICE_WORKSPACE`. A successful edit returns `ok: true` and, when Holix Studio is hosting the file, `editor_refresh: true` or `false`. A refused edit returns `ok: false` and `error`.

Do not unzip the package or rewrite XML in a shell. That skips the Studio reload and can break the document.

## office_list_tool

No arguments. Returns `files` and `count` for `.docx`, `.xlsx`, `.pptx`, `.odt`, `.ods`, and `.odp` under the workspace. Hidden and vendor directories such as `.git`, `node_modules`, and `.venv` are skipped. The list stops at 200 files.

## office_read_tool

| Argument | Meaning |
|---|---|
| `path` | Document, for example `reports/plan.docx` |
| `max_chars` | Plain-text cap. Default 40000 |

Returns the joined text of paragraphs, slide shapes, shared strings, and ODF text nodes. Excel numbers stored only in `<v>` are not in this text. Read those with `office_sheet_tool`.

## office_replace_tool

| Argument | Meaning |
|---|---|
| `path` | Document |
| `old` | Exact text to find |
| `new` | Replacement. May be empty |

A phrase that sits in one run keeps the run's formatting. A phrase split across runs in one paragraph is written back into the first text node of that paragraph. The return field `replaced` is how many matches changed. No match does not refresh the editor.

Editable parts: DOCX document, footnotes, endnotes, comments, headers, and footers; PPTX slides and notes; XLSX shared strings and worksheets; ODT, ODS, and ODP `content.xml`.

## office_format_tool

| Argument | Meaning |
|---|---|
| `path` | Document |
| `text` | Exact phrase |
| `font` | Family, such as `Calibri` |
| `size` | Points, such as `14` or `14pt`. Greater than 0 and at most 500 |
| `bold`, `italic`, `underline` | Set or leave unchanged. Omit a property to leave it |
| `color`, `fill` | `RRGGBB`. `auto` is rejected |
| `style` | Named paragraph style, such as `Heading 1`. DOCX and ODT only |

At least one property is required. Bold, italic, underline, color, font, and size apply to each match of the phrase. `fill` shades a paragraph only when the paragraph's whole stripped text equals the phrase. On XLSX the font and fill apply to each cell whose value or formula contains the phrase. On ODS and ODP the same properties are written as ODF text styles. Named styles on XLSX, PPTX, ODS, or ODP raise an error when style is the only request.

## office_append_tool

| Argument | Meaning |
|---|---|
| `path` | DOCX or ODT |
| `text` | Paragraph text |

Other types raise an error. Use `office_sheet_tool` or `office_slide_tool` to add content there.

## office_media_tool

| Argument | Meaning |
|---|---|
| `path` | Document |
| `media` | Image or video inside the workspace |
| `slide` | 1-based slide, or the sheet number for XLSX |
| `width_cm`, `height_cm` | Frame size. Defaults 12 and 6.8 |

Images: png, jpg, jpeg, gif, webp, bmp, tiff. Video: mp4, m4v, webm, mov, ogg. The largest accepted file is 200 MB.

Video plays on a PPTX slide and an ODP page, and in an ODT frame. DOCX stores the bytes under `word/media/` and adds a hyperlink paragraph. XLSX and ODS reject video. XLSX images are anchored on the chosen sheet.

## office_slide_tool

PPTX and ODP. `slide` is 1-based.

| `action` | Effect |
|---|---|
| `list` | Slides and their text. Does not save |
| `add` | New slide. PPTX `layout` is `title`, `title_body`, `section`, `two`, `title_only`, or `blank` |
| `set` | Replace the title and body of `slide` |
| `notes` | Replace speaker notes. ODP notes are a text box |
| `delete` | Remove `slide` |

`title`, `body`, and `notes` are the text for `add`, `set`, and `notes`. Fonts on existing slide text are `office_format_tool`. Pictures are `office_media_tool`. There is no animation, transition, or slide-master tool.

## office_sheet_tool

XLSX and ODS.

| `action` | Effect |
|---|---|
| `list` | Sheet names |
| `add` | Create `sheet`. Titles are at most 31 characters |
| `read` | Cell values and formulas, capped at 400 cells |
| `set` | Write `value` into `cell` on `sheet`. Numeric strings become numbers |
| `formula` | Write `formula`. `=Other!A1` and `=SUM(A1:A2)` are accepted. ODS stores `of:=[Other.A1]` and `of:=SUM([.A1:.A2])` |
| `link` | Hyperlink in `cell`. `link` is the URL |
| `name` | XLSX defined name. Pass the name in `sheet` and a range such as `Sheet1!$A$1:$B$10` in `range_ref`. ODS defined names are rejected |

`list` and `read` do not refresh the editor.

## office_chart_tool

| Argument | Meaning |
|---|---|
| `chart_type` | `bar`, `column`, `line`, or `pie` |
| `path` | XLSX or PPTX |
| `sheet` | XLSX sheet |
| `data_range` | XLSX series, including the title cell, such as `B1:B4` |
| `categories_range` | XLSX labels, such as `A2:A4` |
| `anchor` | XLSX cell for the chart. Default `E2` |
| `title` | Chart title |
| `slide` | 1-based PPTX slide |
| `categories`, `values` | PPTX comma-separated lists of equal length, such as `Jan,Feb` and `10,12` |
| `series_name` | PPTX series name |

ODS and ODP charts are not written.

## office_table_tool

| Argument | Meaning |
|---|---|
| `action` | `read`, `add`, or `set` |
| `rows_json` | JSON list of rows for `add`, such as `[["Name","Qty"],["A","1"]]` |
| `table` | 1-based table index for `read` and `set` |
| `row`, `col` | 0-based cell for `set` |
| `value` | New cell text for `set` |
| `slide` | 1-based PPTX or ODP slide |
| `header` | First row is a header. Default true |

`add` accepts at most 200 rows and 40 columns. On XLSX, `add` writes the grid at A1 and creates an Excel table. On ODS, `add` writes the grid onto the first sheet. `set` of one cell is DOCX, PPTX, ODT, and ODP. `read` does not refresh the editor.

## Errors you can trust

| Message | What to change |
|---|---|
| Path must stay inside the workspace | Use a relative path under the workspace |
| Path must be a docx, xlsx, pptx, odt, ods, or odp file | The target is not one of the six types |
| Named styles are available for docx and odt | Drop `style` or use a DOCX/ODT file |
| Appending a paragraph supports docx and odt | Use a sheet or slide tool |
| Charts are not written into ods or odp | Use XLSX or PPTX |
