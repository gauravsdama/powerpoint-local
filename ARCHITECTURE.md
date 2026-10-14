# architect.md

## Objective

Ship a local-only Codex MCP that turns local documents into stylized PowerPoint deck plans and supports local template discovery, review exports, and guided deck creation.

## Selected Architect Modes

- Primary: Systems Architect
- Supporting: Assumption Auditor, Performance & UX Architect, Product Direction Architect
- Rationale: the strict offline boundary, macOS automation lifecycle, and template-selection flow are the high-risk product decisions.

## Assumption Ledger

| Item | Type | Current belief | Risk if wrong | Decision |
|---|---|---|---|---|
| Platform | known fact | macOS with Microsoft PowerPoint 16.95.1 | Medium | Use AppleScript only for PowerPoint actions. |
| Oracle template | unknown | A usable template might only be represented by SharePoint cache metadata | High | Never fetch; require the user to save it locally. |
| Dependencies | safe assumption | A standard-library server is sufficient for the required workflow | Low | Keep deployment dependency-free. |
| Local cache schema | risky assumption | Office cache storage changes across versions | Medium | Scan bounded files defensively and return evidence, not inferred template content. |
| Macro handling | unknown | The owner may need either macro preservation or macro removal | High | Do not guess: reject `.pptm` and `.potm` during `.pptx` preparation until policy is chosen. |

## Decisions Made

- Use stdio MCP and the Python standard library rather than a dependency-managed SDK.
- Own state in explicit local artifacts: deck plans and manifests; no hidden global presentation state.
- Treat a template as usable only after a local file exists. SharePoint hits are `reference_only`.
- Prefer Oracle candidate templates by filename/layout evidence; otherwise identify the best local candidate and disclose the fallback.
- Copy a supplied local `.pptx` for guided editing; require explicit PowerPoint conversion for `.potx`, and reject macro-enabled packages rather than byte-renaming them or guessing macro policy.
- Invoke Mermaid Studio only through a configurable local stdio child process. PowerPoint Local rejects non-loopback renderer gateway configuration; Mermaid Studio owns the loopback renderer connection.
- Reject existing artifact paths. Keep PowerPoint-facing transaction files in the selected destination directory, stage completed PNGs in a sibling directory, then publish only completed artifacts.
- Parse newline-framed JSON-RPC deterministically, including standard parse/invalid-request errors, notifications, and batches, with a 1 MiB request ceiling.

## Performance and UX Guardrails

- Bound plain-document reads to 250 KiB, DOCX packages to 64 MiB and their document XML to 2 MiB, inspected PowerPoint XML members to 4 MiB / 32 MiB total, MCP requests to 1 MiB, Mermaid responses to 2 MiB, and cache-metadata scans to 250 files / 8 MiB each.
- Paginate template search; do not recursively enumerate unrestricted home directories.
- Keep MCP responses structured and concise; write full plans to explicit local manifest files when requested.
- Return permission-specific recovery guidance for Automation and Screen Recording.
- Use monotonic Mermaid deadlines, drain both child pipes, and terminate, kill if necessary, wait for, and close every child resource.

## Non-goals

- Network retrieval, cloud API access, SharePoint authentication, online template discovery, package installation, or live collaboration.
- Claiming a remote template was downloaded or used.
- Choosing whether macros should be preserved, stripped, or rejected on the owner's behalf.

## Validation Checklist

- [x] Upstream references cloned locally and retained as non-runtime sources.
- [x] PowerPoint application presence and AppleScript responsiveness checked locally.
- [x] Automation, folder access, and Screen Recording were exercised in the current user's TCC context.
- [x] Unit and stdio smoke tests cover the deterministic core and local-only import policy.

## Owner Decision Required

Decide whether a future release should preserve macros in `.pptm`/`.potm`, strip them through an explicit PowerPoint conversion, or continue rejecting those packages. The current implementation safely rejects them and does not alter macro-bearing files.
