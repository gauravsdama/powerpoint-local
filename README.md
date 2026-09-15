# PowerPoint Local

`powerpoint-local` is a macOS stdio MCP server for turning local source documents into PowerPoint deck plans without making network calls. It uses the Python standard library, local files, AppleScript, screen capture, and an optional configurable bridge to Mermaid Studio on the same Mac.

## What it does

- Finds local `.pptx`, `.pptm`, `.potx`, and `.potm` files in Office template and cache locations.
- Reads local PowerPoint MRU/backstage cache metadata. SharePoint and Oracle hits are reported as references only; the server never opens a URL or downloads a template.
- Reads local `.md`, `.txt`, and `.docx` files and creates deterministic deck plans.
- Inspects template layouts and maps plan roles to locally available layouts, preferring Oracle-branded templates.
- Suggests Mermaid source for process, architecture, and timeline slides.
- Calls the local Mermaid Studio MCP over stdio to turn approved Mermaid source into a local PNG artifact for a planned slide.
- Creates a local editable template copy or blank `.pptx` plus a deck manifest for guided authoring, opens PowerPoint locally, exports slide PNGs, and captures a local screenshot for visual review.

## Local-only boundary

The server does not import networking libraries, invoke a shell, call HTTP endpoints, fetch packages, search online, or connect to cloud services. A SharePoint result is metadata only. Save an Oracle template locally before using it for strict local-only generation.

The Mermaid collaboration is also local: PowerPoint Local starts the configured Mermaid Studio MCP command over stdio. Mermaid Studio's artifact gateway must already be running with `npm run start` because its renderer owns the loopback connection. A sibling `mermaid-studio` checkout and the first `node` on `PATH` are detected automatically. Override them with `POWERPOINT_LOCAL_MERMAID_ENTRY` and `POWERPOINT_LOCAL_MERMAID_NODE`; `POWERPOINT_LOCAL_MERMAID_GATEWAY` accepts loopback URLs only.

## Register and check the server

For a standalone local install that does not download Python packages:

```sh
./scripts/setup.sh --check
./scripts/setup.sh --prefix "$HOME/.local"
```

Point the MCP client at the server script in your checkout:

```toml
[mcp_servers.powerpoint_local]
command = "/usr/bin/python3"
args = ["/path/to/powerpoint-local/src/powerpoint_local/server.py"]
```

After restarting the MCP client, call `powerpoint_local_status` first. It reports the local PowerPoint version, template and metadata roots, Mermaid bridge availability, and required macOS permissions.

A useful first workflow is:

1. Call `powerpoint_local_read_source` with a local Markdown, text, or DOCX file.
2. Call `powerpoint_local_find_templates` or provide a local template path.
3. Call `powerpoint_local_plan_deck` to produce a deterministic plan.
4. Review the plan before calling `powerpoint_local_prepare_deck` with an output path.

When a local template is supplied, the server creates an editable copy plus a manifest for guided authoring. Without a template, it asks PowerPoint to create and save a blank local `.pptx` at the requested path. It does not synthesize a finished presentation from arbitrary prose.

Run the local tests without installing anything:

```bash
/usr/bin/python3 -m unittest discover -s tests -v
```

Use `./scripts/release_check.sh` for the tests, offline-install smoke test, privacy/path guards, and MCP handshake in one command.

## Permission fixes

If a tool reports an automation or capture failure, enable these for the process that starts Codex (normally Codex or ChatGPT):

1. **System Settings → Privacy & Security → Automation**: allow it to control Microsoft PowerPoint.
2. **System Settings → Privacy & Security → Screen Recording**: allow it to capture a review screenshot.
3. Quit and reopen Codex after changing either permission.

`powerpoint_local_status` exposes the PowerPoint version and the exact local roots searched.

The tested floor for PowerPoint for Mac is **16.95**. This is the oldest version in the exercised 16.95.x line; older builds are reported as unsupported until tested. Passing on 16.95.1 does not prove that every later Microsoft 365 build behaves identically, so compatibility reports should always include the exact Office version.

On PowerPoint builds where the AppleScript PNG exporter returns without producing files, the MCP falls back locally: PowerPoint renders one-slide temporary copies as PDF and macOS `sips` converts each page to PNG. The original deck is never edited by the fallback.

## Upstream references

The `upstream/` directory contains pinned, unmodified clones used for design study:

- `jongalloway/pptx-tools` at `1ff4fd10e41a9bb33853eb7b8ca5a61675571673`
- `GongRzhe/Office-PowerPoint-MCP-Server` at `3631ba2ec0c24504476f78bf74d329c9be11caaa`

They are not runtime dependencies.

See [third-party notices](THIRD_PARTY_NOTICES.md) for exact revisions and licensing. The project is licensed under Apache-2.0; the [release-readiness record](docs/RELEASE_READINESS.md) tracks live Office and Mermaid verification.

## License

Copyright 2026 Gaurav Dama. Licensed under the [Apache License 2.0](LICENSE).
