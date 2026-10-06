# MCP Collabora Holix

An MCP server that reads and edits the office files in one folder: Word, Excel, PowerPoint, and LibreOffice.

It talks to the agent over stdio. The agent calls tools. The tools change the real `.docx`, `.xlsx`, `.pptx`, `.odt`, `.ods`, and `.odp` on disk. Nothing is uploaded to a third-party converter.

[Русский](README.ru.md) · [Tool reference](docs/en/reference.md) · [Справочник](docs/ru/reference.md)

## What it does

Point the server at a directory. Ask the agent to find a document, read it, change the wording, restyle a phrase, add a picture, build a sheet formula, or add a slide. The file on disk is the result.

Holix Studio connects this server by itself when office editing is turned on, and an open Collabora tab reloads after a successful edit. Claude, OpenClaw, Hermes, Codex, and Holix use the same server with the snippets below.

## What you can ask for

| | |
|---|---|
| Documents | DOCX, XLSX, PPTX, ODT, ODS, ODP in the workspace folder |
| Text | List files, read text, replace a phrase, append a paragraph to DOCX or ODT |
| Appearance | Font, size, bold, italic, underline, color, and paragraph shading. Named styles such as Heading 1 on DOCX and ODT |
| Media | PNG, JPEG, GIF, WebP, BMP, and TIFF in all six types. MP4, WebM, and MOV play on PPTX and ODP slides and in an ODT frame. A DOCX stores the video and a link |
| Presentations | List, add, retitle, write notes, and delete PPTX and ODP slides |
| Sheets | Add sheets, set cells, formulas, hyperlinks, and Excel defined names. ODS formulas accept Excel-style references such as `=Other!A1` |
| Charts | Bar, line, and pie charts on XLSX and PPTX |
| Tables | Read, add, and edit a cell in DOCX, PPTX, ODT, and ODP. Add a grid to XLSX or ODS |

Paths are relative to the workspace, for example `reports/plan.docx` and `media/cover.png`. The server refuses a path that leaves that directory.

## Honest limits

- Paragraph shading applies only when the paragraph's whole text is the phrase you name.
- Named styles are DOCX and ODT. Charts are XLSX and PPTX. Defined names are XLSX.
- A DOCX does not play video. The file is stored and linked.
- There are no slide animations, transitions, or master slides.
- `office_read_tool` returns text. Use `office_sheet_tool` with `read` for numbers and formulas.
- Appending a paragraph is DOCX and ODT.
- In Holix Studio, a reload replaces unsaved keystrokes in that Collabora tab with the saved file.

The full contracts are in the [tool reference](docs/en/reference.md).

## Install

Python 3.12 or newer.

```bash
pip install "git+https://github.com/javded-itres/mcp-collabora-holix.git"
mcp-collabora-holix --help 2>/dev/null || true
python -m holix_office
```

`python -m holix_office` waits on stdin for an MCP client. That is expected. Set the workspace before the client starts the process:

```bash
export HOLIX_OFFICE_WORKSPACE="$HOME/Documents"
```

Use an absolute path. A new chat is required after you add or change the server, so the client starts a fresh process.

## Connect

### Holix Studio

No snippet. When office editing is enabled, Studio writes the `holix_office` server into the Holix profile, assigns it to the main agent, and installs the `office_documents` skill. The workspace is the profile workspace, not whatever directory the chat happens to use.

Open a new chat after office editing is turned on. The running MCP process does not reload tools by itself.

### Holix

In the profile `config.yaml`:

```yaml
mcp_servers:
  holix_office:
    transport: stdio
    command: mcp-collabora-holix
    args: []
    env:
      HOLIX_OFFICE_WORKSPACE: /absolute/path/to/documents
```

Start a new session. Studio users should leave this block alone: Studio owns the `holix_office` entry it created and will refresh it.

### Claude

Claude Desktop, macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`. Windows: `%APPDATA%\Claude\claude_desktop_config.json`.

```json
{
  "mcpServers": {
    "mcp-collabora-holix": {
      "command": "mcp-collabora-holix",
      "args": [],
      "env": {
        "HOLIX_OFFICE_WORKSPACE": "/absolute/path/to/documents"
      }
    }
  }
}
```

Quit Claude Desktop and open it again.

Claude Code, from the project directory:

```bash
claude mcp add --transport stdio --env HOLIX_OFFICE_WORKSPACE=/absolute/path/to/documents mcp-collabora-holix -- mcp-collabora-holix
```

If `mcp-collabora-holix` is not on `PATH`, use the full path printed by `command -v mcp-collabora-holix`, or `python` with args `["-m", "holix_office"]`.

### OpenClaw

```bash
openclaw mcp add mcp-collabora-holix \
  --command mcp-collabora-holix \
  --env HOLIX_OFFICE_WORKSPACE=/absolute/path/to/documents
openclaw mcp doctor mcp-collabora-holix --probe
```

The same server in OpenClaw config:

```text
mcp:
  servers:
    mcp-collabora-holix:
      command: mcp-collabora-holix
      transport: stdio
      enabled: true
      env:
        HOLIX_OFFICE_WORKSPACE: /absolute/path/to/documents
```

### Hermes

`~/.hermes/config.yaml`:

```yaml
mcp_servers:
  mcp-collabora-holix:
    command: mcp-collabora-holix
    args: []
    env:
      HOLIX_OFFICE_WORKSPACE: /absolute/path/to/documents
```

Then `hermes chat`. The tools show up as `mcp__mcp-collabora-holix__office_list_tool` and the other `office_*` tools.

### Codex

`~/.codex/config.toml`, or `.codex/config.toml` in a trusted project:

```toml
[mcp_servers.mcp-collabora-holix]
command = "mcp-collabora-holix"
args = []

[mcp_servers.mcp-collabora-holix.env]
HOLIX_OFFICE_WORKSPACE = "/absolute/path/to/documents"
```

Or:

```bash
codex mcp add mcp-collabora-holix --env HOLIX_OFFICE_WORKSPACE=/absolute/path/to/documents -- mcp-collabora-holix
```

In the Codex TUI, `/mcp` lists the server.

## Workspace

| Variable | Meaning |
|---|---|
| `HOLIX_OFFICE_WORKSPACE` | Absolute path of the folder the tools may touch |
| `HOLIX_STUDIO_WORKSPACE_ROOT` | Set by Holix Studio. Used when the public variable is absent |
| `HOLIX_PROFILE` | Studio profile name, used only to refresh an open editor |

If neither path variable is set, the process working directory is the workspace.

## License

MIT. Copyright 2026 Pavel Lukyanov and ITRES contributors.
