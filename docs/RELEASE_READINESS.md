# Release readiness

Status: **source-ready**. The repository is licensed under Apache-2.0, local tests and offline installation pass, and the complete local PowerPoint/Mermaid workflow has live evidence.

## Product and provenance

- Intended user: a macOS PowerPoint user who wants deterministic local deck planning, local template inspection, and explicit review artifacts through an MCP client.
- Employer-facing story: a standard-library Python MCP that keeps source documents and Office automation on the Mac, models its privacy boundary in code, and delegates optional diagram rendering to a separately running local Mermaid Studio.
- Repository: `gauravsdama/powerpoint-local`, private, not a GitHub fork, with one author in the local commit history.
- Source: first-party Python. Two pinned MIT-licensed repositories were design references only; see [third-party notices](../THIRD_PARTY_NOTICES.md).
- Interface: stdio MCP only. Microsoft PowerPoint and Mermaid Studio own their respective interfaces, so this repository has no UI copy inventory or project screenshot requirement.

## Evidence recorded 2026-09-15

- `python3 -m unittest discover -s tests -v`: 8 tests passed, including mocked PDF-to-PNG fallback coverage, failure-path presentation cleanup, the full-package forbidden-network-import check, stdio handshake, deterministic planning, loopback-only Mermaid configuration, blank-deck creation boundary, and tested PowerPoint floor.
- `./scripts/setup.sh --check`: Python requirement passed; local PowerPoint and Mermaid availability were reported without modifying files.
- `./scripts/release_check.sh`: checks compilation, tests, offline temporary installation, MCP initialization, personal paths, and tracked private/generated artifacts.
- `powerpoint_local_status`: Microsoft PowerPoint 16.95.1 is installed and AppleScript automation responded; the configured Node executable and Mermaid MCP build were found.
- Mermaid Studio's dirty checkout was preserved, rebuilt, typechecked, and linted. PowerPoint Local then created and visually inspected a high-resolution Mermaid PNG through the stdio bridge and loopback renderer.
- A disposable three-slide `.pptx` exercised source reading, template discovery, planning, Mermaid suggestion/rendering, blank-deck creation, inspection, open, PNG export, and screen capture. The deck contained real text and a Mermaid image inserted by Microsoft PowerPoint.
- PowerPoint 16.95.1 returned success without files for direct AppleScript PNG export. The repaired fallback used PowerPoint-rendered one-slide PDFs plus macOS `sips`, produced three PNGs, and preserved the source deck. Independent and local visual review found no clipping or overlap after enlarging the diagram and aligning slide margins.
- Local upstream study checkouts match the two commits recorded in `THIRD_PARTY_NOTICES.md`; `upstream/` remains ignored.

## Local-only boundary

- PowerPoint Local imports no network client modules and does not open remote references. Cache hits can include sensitive metadata and are returned only when explicitly requested.
- The optional Mermaid child is started over stdio. Its renderer may use only a configured `localhost` or `127.0.0.1` gateway; non-loopback gateway values are rejected.
- Source paths, template paths, deck content, cache references, screenshots, and exported slides stay local but may still be sensitive. `exports/`, manifests, and upstream study clones are ignored.

## Release boundary and blockers

- PowerPoint for Mac 16.95 is the tested floor: older builds are reported unsupported until tested. This records the oldest version exercised successfully; it is not a guarantee that every later Microsoft 365 build behaves identically.
- Repeat the disposable workflow in a clean macOS user account or another compatible Mac before a tagged release.

## Release checklist

- Run `./scripts/release_check.sh` and `./scripts/setup.sh --check`.
- Configure Mermaid with `POWERPOINT_LOCAL_MERMAID_ENTRY` when it is not a sibling checkout; keep the gateway loopback-only.
- Exercise each mutating tool with disposable local files and inspect its artifacts.
- Repeat the install and MCP handshake in a clean macOS user account or compatible Mac.
- Run `./scripts/release_check.sh --publish` from a clean release commit.
