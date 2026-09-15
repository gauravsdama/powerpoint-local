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

## Decisions Made

- Use stdio MCP and the Python standard library rather than a dependency-managed SDK.
- Own state in explicit local artifacts: deck plans and manifests; no hidden global presentation state.
- Treat a template as usable only after a local file exists. SharePoint hits are `reference_only`.
- Prefer Oracle candidate templates by filename/layout evidence; otherwise identify the best local candidate and disclose the fallback.
- Copy a supplied local template for guided editing; without a template, write the manifest and activate PowerPoint rather than claiming hidden content mutation succeeded.
- Invoke Mermaid Studio only through a configurable local stdio child process. PowerPoint Local rejects non-loopback renderer gateway configuration; Mermaid Studio owns the loopback renderer connection.

## Performance and UX Guardrails

- Bound document reads to 250 KiB and cache-metadata scans to 250 files / 8 MiB each.
- Paginate template search; do not recursively enumerate unrestricted home directories.
- Keep MCP responses structured and concise; write full plans to explicit local manifest files when requested.
- Return permission-specific recovery guidance for Automation and Screen Recording.

## Non-goals

- Network retrieval, cloud API access, SharePoint authentication, online template discovery, package installation, or live collaboration.
- Claiming a remote template was downloaded or used.

## Validation Checklist

- [x] Upstream references cloned locally and retained as non-runtime sources.
- [x] PowerPoint application presence and AppleScript responsiveness checked locally.
- [ ] Validate Automation and Screen Recording in the user’s TCC context when first invoking export/capture.
- [x] Unit and stdio smoke tests cover the deterministic core and local-only import policy.
