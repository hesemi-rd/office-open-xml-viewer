# ooxml-mcp-server

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that lets AI agents read Excel, Word, and PowerPoint files without any additional code.

---

## Easiest: VS Code extension (recommended for VS Code users)

Install the [Office Viewer extension](https://marketplace.visualstudio.com/items?itemName=silurus.office-open-xml-viewer). Open a workspace that contains an `.xlsx`, `.docx`, or `.pptx` file and accept the prompt — the extension downloads a prebuilt binary (~5 MB, SHA256-verified) and registers the MCP server with VS Code automatically. Copilot Agent mode and any other MCP-aware agent in VS Code picks it up with no further configuration.

If you don't use VS Code, or want to wire this into Claude Code / Codex CLI / a different editor, follow the manual install below.

---

## Manual install: prebuilt binaries

Each release ships prebuilt binaries on the [Releases page](https://github.com/yukiyokotani/office-open-xml-viewer/releases/latest):

| Platform | Asset name |
|----------|------------|
| macOS (Apple Silicon) | `ooxml-mcp-server-aarch64-apple-darwin` |
| macOS (Intel) | `ooxml-mcp-server-x86_64-apple-darwin` |
| Linux x64 | `ooxml-mcp-server-x86_64-unknown-linux-gnu` |
| Linux arm64 | `ooxml-mcp-server-aarch64-unknown-linux-gnu` |
| Windows x64 | `ooxml-mcp-server-x86_64-pc-windows-msvc.exe` |

Each asset has an accompanying `.sha256` file. Download, verify, mark executable, and place anywhere on your `PATH`:

```bash
TAG=v0.51.0   # replace with the latest tag from the Releases page
ASSET=ooxml-mcp-server-aarch64-apple-darwin   # pick your platform
curl -L -o ooxml-mcp-server  "https://github.com/yukiyokotani/office-open-xml-viewer/releases/download/${TAG}/${ASSET}"
curl -L -o sums.txt         "https://github.com/yukiyokotani/office-open-xml-viewer/releases/download/${TAG}/${ASSET}.sha256"
shasum -a 256 -c sums.txt   # must print "OK"
chmod +x ooxml-mcp-server
mv ooxml-mcp-server /usr/local/bin/
```

---

## Manual install: build from source

Skip this section unless you want to build from source.

```bash
# Install Rust if needed: https://rustup.rs
cargo install --git https://github.com/yukiyokotani/office-open-xml-viewer.git \
  --package ooxml-mcp-server
```

The binary is placed in `~/.cargo/bin/ooxml-mcp-server`. Make sure `~/.cargo/bin` is on your `PATH`.

---

## Configure your AI client

Pick the client you use and follow the instructions below.

### Claude Code

Create `.mcp.json` in your project root (or `~/.claude.json` for all projects):

```json
{
  "mcpServers": {
    "ooxml": {
      "type": "stdio",
      "command": "ooxml-mcp-server"
    }
  }
}
```

Start Claude Code in that directory and run `/mcp` to confirm the server shows as connected.

**Try it:**

```
> What sheets are in /Users/me/Documents/budget.xlsx?
```

---

### GitHub Copilot (VS Code)

Create `.vscode/mcp.json` in your workspace root:

```json
{
  "servers": {
    "ooxml": {
      "type": "stdio",
      "command": "ooxml-mcp-server"
    }
  }
}
```

Open the Command Palette (`⇧⌘P`) → **MCP: List Servers** to confirm the server is running.

> MCP tools are only available in **Agent mode**. In the Copilot Chat panel, click the mode selector and choose **Agent** before asking a question.

**Try it:**

```
Extract all text from /Users/me/Documents/deck.pptx
```

---

### Codex CLI (OpenAI)

Add to `~/.codex/config.toml`:

```toml
[mcp_servers.ooxml]
command = "ooxml-mcp-server"
args = []
```

Restart Codex, then run `codex mcp list` to verify registration.

**Try it:**

```bash
codex "Show me all formulas in Sheet1 of /Users/me/Documents/model.xlsx"
```

---

## Troubleshooting: command not found in MCP config

Some launchers (VS Code, Codex) do not inherit your shell `PATH`. If the server fails to start with a "command not found" error, use the full path to the binary instead of just the name:

```bash
# Find the full path
echo ~/.cargo/bin/ooxml-mcp-server
```

Then in your config:

```json
"command": "/Users/you/.cargo/bin/ooxml-mcp-server"
```

```toml
command = "/Users/you/.cargo/bin/ooxml-mcp-server"
```

---

## Available tools

### xlsx (Excel)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `xlsx_parse` | `path` | All sheet names and IDs |
| `xlsx_get_sheet_names` | `path` | Sheet name list |
| `xlsx_get_sheet_dimensions` | `path`, `sheet` | Number of rows and columns |
| `xlsx_get_cell_range` | `path`, `sheet`, `range` | Cell values and formulas for a range like `"A1:C10"` |
| `xlsx_get_formulas` | `path`, `sheet` | Every formula cell with its cached value |
| `xlsx_search_cells` | `path`, `query`, `sheet?` | Cells whose value or formula contains the query string |

`sheet` can be a name (`"Sheet1"`) or a 0-based index (`"0"`). For `xlsx_search_cells`, omitting `sheet` searches all sheets.

### docx (Word)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `docx_extract_text` | `path` | All text as plain string |
| `docx_get_structure` | `path` | Paragraph and table structure with style info |
| `docx_get_tables` | `path` | All tables with each cell's text |
| `docx_search_text` | `path`, `query` | Matching paragraphs and table cells with their position |

### pptx (PowerPoint)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `pptx_get_slides` | `path` | Slide count and each slide's title |
| `pptx_extract_text` | `path`, `slide_index?` | Text from all slides, or one slide (0-based index) |
| `pptx_get_slide_structure` | `path`, `slide_index` | All shapes with position, size, and text |
| `pptx_search_text` | `path`, `query` | Matching slide numbers and text snippets |

All `path` parameters require absolute paths (e.g. `/Users/me/Documents/file.xlsx`).  
All search tools use **case-insensitive substring matching**.
