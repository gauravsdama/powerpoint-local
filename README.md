# PowerPoint Local

`powerpoint-local` is a macOS stdio MCP server for turning local source documents into PowerPoint deck plans without making any network calls. It uses only the Python standard library, local files, `osascript`, and `screencapture`.

## What it does

- Finds local `.pptx`, `.pptm`, `.potx`, and `.potm` files in Office template and cache locations.
- Reads local PowerPoint MRU/backstage cache metadata. SharePoint and Oracle hits are reported as references only; the server never opens a URL or downloads a template.
- Reads local `.md`, `.txt`, and `.docx` files and creates deterministic deck plans.
- Inspects template layouts and maps plan roles to locally available layouts, preferring Oracle-branded templates.
- Suggests Mermaid source for process, architecture, and timeline slides.
- Calls the local Mermaid Studio MCP over stdio to turn approved Mermaid source into a local PNG artifact for a planned slide.
- Creates a local editable template copy plus a deck manifest for guided authoring, opens PowerPoint locally, exports slide PNGs, and captures a local screenshot for visual review.

## Local-only boundary

The server does not import networking libraries, invoke a shell, call HTTP endpoints, fetch packages, search online, or connect to cloud services. A SharePoint result is metadata only. Save an Oracle template locally before using it for strict local-only generation.

The Mermaid collaboration is also local: PowerPoint Local starts the fixed Mermaid Studio MCP command over stdio. Mermaid Studio must already be running locally (`npm run start` in `/Users/gauravsdama/git/mermaid-studio`) because its renderer owns the loopback gateway.

## Install / run

The Codex global server registration is:

```toml
[mcp_servers.powerpoint_local]
command = "/usr/bin/python3"
args = ["/Users/gauravsdama/git/powerpoint-local/src/powerpoint_local/server.py"]
```

Run the tests without installing anything:

```bash
/usr/bin/python3 -m unittest discover -s tests -v
```

## Permission fixes

If a tool reports an automation or capture failure, enable these for the process that starts Codex (normally Codex or ChatGPT):

1. **System Settings → Privacy & Security → Automation**: allow it to control Microsoft PowerPoint.
2. **System Settings → Privacy & Security → Screen Recording**: allow it to capture a review screenshot.
3. Quit and reopen Codex after changing either permission.

`powerpoint_local_status` exposes the PowerPoint version and the exact local roots searched.

On some PowerPoint builds the AppleScript PNG exporter can return without producing files. The MCP detects this and reports it as an error; use `powerpoint_local_capture_screen` for a local visual-review capture while that Office limitation is present.

## Upstream references

The `upstream/` directory contains pinned, unmodified clones used for design study:

- `jongalloway/pptx-tools` at `1ff4fd10e41a9bb33853eb7b8ca5a61675571673`
- `GongRzhe/Office-PowerPoint-MCP-Server` at `3631ba2ec0c24504476f78bf74d329c9be11caaa`

They are not runtime dependencies.
