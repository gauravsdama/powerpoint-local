# Release readiness

Status: **source-ready**. The repository is licensed under Apache-2.0, local tests and offline installation pass, and the complete local PowerPoint/Mermaid workflow has live evidence.

## Product and provenance

- Intended user: a macOS PowerPoint user who wants deterministic local deck planning, local template inspection, and explicit review artifacts through an MCP client.
- Employer-facing story: a standard-library Python MCP that keeps source documents and Office automation on the Mac, models its privacy boundary in code, and delegates optional diagram rendering to a separately running local Mermaid Studio.
- Repository: `gauravsdama/powerpoint-local`, public, not a GitHub fork, with one author in the local commit history.
- Source: first-party Python. Two pinned MIT-licensed repositories were design references only; see [third-party notices](../THIRD_PARTY_NOTICES.md).
- Interface: stdio MCP only. Microsoft PowerPoint and Mermaid Studio own their respective interfaces, so this repository has no UI copy inventory or project screenshot requirement.

## Evidence recorded 2026-10-14

- `python3 -m unittest discover -s tests -v`: 18 tests passed, including transactional/stale-safe PNG and screenshot publication, failure-path presentation cleanup, no-overwrite artifact ownership, bounded source and ZIP inspection, deterministic JSON-RPC errors/batches/notifications, duplicate-title prevention, Mermaid deadline and child-reaping behavior, the full-package forbidden-network-import check, stdio handshake, deterministic planning, loopback-only Mermaid configuration, blank-deck creation, package-type rejection, timeout recovery guidance, and the tested PowerPoint floor.
- `./scripts/setup.sh --check`: Python requirement passed; local PowerPoint and Mermaid availability were reported without modifying files.
- `./scripts/setup_macos.sh --test --yes` and `./scripts/release_check.sh` passed. The release check compiled the package, installed it into a disposable directory without downloading runtime dependencies, completed an MCP handshake, and checked tracked paths and generated artifacts.
- Microsoft PowerPoint 16.95.1 from the Mac App Store was exercised on macOS 26.6.2 (Apple silicon). This remains the oldest tested version; an older licensed installation was not available.
- The local Mermaid bridge rendered a PNG through its stdio MCP and loopback renderer. A disposable three-slide presentation then exercised source reading, planning, Mermaid suggestion/rendering, blank-deck creation, editing in PowerPoint, embedded-image insertion, save/reopen, structural inspection, PNG export, and screen capture.
- PowerPoint 16.95.1 returned no files from direct AppleScript PNG export. The PDF-per-slide fallback produced three 960×540 PNGs through PowerPoint and macOS `sips`; visual review found the final title, bullets, and embedded Mermaid image unclipped and non-overlapping.
- The first folder-scoped PowerPoint access prompt was granted for the disposable destination. A second export completed without another staging-folder prompt after PowerPoint-facing transaction files were moved out of private per-operation subdirectories.
- A malformed presentation failed without publishing its output directory or leaving transaction files, and PowerPoint remained usable with the valid deck open. A denied system destination left no deck. A malformed Mermaid request returned a bounded parser error and left the Mermaid child-process set unchanged.
- No second non-system macOS account is present on this Mac. The disposable offline install proves that the package does not depend on files inside this checkout, but it does not substitute for the clean-account or second-Mac release check.
- Local upstream study checkouts match the two commits recorded in `THIRD_PARTY_NOTICES.md`; `upstream/` remains ignored.
- GitHub Actions run `38638110667` passed commit `bbc5929`; the hosted runner intentionally cannot claim live PowerPoint compatibility because PowerPoint is not installed there.

## Local-only boundary

- PowerPoint Local imports no network client modules and does not open remote references. Cache hits can include sensitive metadata and are returned only when explicitly requested.
- The optional Mermaid child is started over stdio. Its renderer may use only a configured `localhost` or `127.0.0.1` gateway; non-loopback gateway values are rejected.
- Source paths, template paths, deck content, cache references, screenshots, and exported slides stay local but may still be sensitive. `exports/`, manifests, and upstream study clones are ignored.
- Artifact-producing tools reject existing deck, manifest, export-directory, and screenshot paths. Failed PNG exports remain staged and do not publish a partial output directory.
- PowerPoint-facing transaction files stay directly in the user-selected destination directory. This respects PowerPoint's folder-scoped file-access grant and avoids a new permission prompt for every private staging subdirectory; transaction files are removed after success or failure.

## Release boundary and blockers

- PowerPoint for Mac 16.95 is the tested floor: older builds are reported unsupported until tested. This records the oldest version exercised successfully; it is not a guarantee that every later Microsoft 365 build behaves identically.
- Guided preparation accepts `.pptx` only. `.potx` requires explicit conversion in PowerPoint; `.pptm` and `.potm` remain rejected until the owner chooses whether macros should be preserved, stripped, or unsupported.
- Repeat the disposable workflow in a clean macOS user account or another compatible Mac before a tagged release.

## Release checklist

- Run `./scripts/release_check.sh` and `./scripts/setup.sh --check`.
- Configure Mermaid with `POWERPOINT_LOCAL_MERMAID_ENTRY` when it is not a sibling checkout; keep the gateway loopback-only.
- Exercise each mutating tool with disposable local files and inspect its artifacts.
- Repeat the install and MCP handshake in a clean macOS user account or compatible Mac.
- Run `./scripts/release_check.sh --publish` from a clean release commit.

### Clean-account check still required

Use an existing authorized test account or a second Mac; do not test against another person's Office profile or documents.

1. Install a licensed Microsoft PowerPoint and record the exact PowerPoint build, macOS version, and Mac architecture.
2. Copy or clone only the tracked repository, then run `./scripts/setup.sh --check`, `./scripts/release_check.sh`, and an install into a new local prefix.
3. Start the separately installed Mermaid Studio renderer and confirm `powerpoint_local_status` reports only the loopback gateway.
4. From a disposable local source file, run read, plan, Mermaid render, deck preparation, open, save/reopen, PNG export, and screen capture. Grant PowerPoint access only to the disposable destination folder.
5. Inspect every PNG, confirm the source deck still opens, confirm a repeated export uses a new output path without another staging-folder prompt, and verify no transaction files, personal paths, or artifacts appear in Git status.
