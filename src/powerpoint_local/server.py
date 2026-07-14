from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Callable

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from powerpoint_local import VERSION
from powerpoint_local.documents import read_source
from powerpoint_local.local_files import cache_references, find_templates, metadata_roots, template_roots
from powerpoint_local.mermaid_bridge import bridge_status, create_diagram
from powerpoint_local.planning import create_plan, suggest_mermaid
from powerpoint_local.presentation import capture_screen, export_pngs, inspect_template, open_in_powerpoint, powerpoint_version, prepare_guided_deck

SERVER_NAME = "powerpoint-local"


def schema(properties: dict[str, Any] | None = None, required: list[str] | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {"type": "object", "properties": properties or {}, "additionalProperties": False}
    if required:
        value["required"] = required
    return value


TOOLS: list[dict[str, Any]] = [
    {"name": "powerpoint_local_status", "description": "Report local PowerPoint, local-only policy, and searched cache locations.", "inputSchema": schema()},
    {"name": "powerpoint_local_find_templates", "description": "Find local PowerPoint templates. Oracle candidates sort first; cache references are metadata only.", "inputSchema": schema({"query": {"type": "string"}, "limit": {"type": "integer", "minimum": 1, "maximum": 100}, "include_cache_references": {"type": "boolean"}})},
    {"name": "powerpoint_local_inspect_template", "description": "Inventory local PPTX/POTX layouts, placeholder roles, and slide text.", "inputSchema": schema({"file_path": {"type": "string", "description": "Absolute local .pptx/.potx path."}}, ["file_path"])},
    {"name": "powerpoint_local_read_source", "description": "Read a local Markdown, text, or DOCX source document into sections.", "inputSchema": schema({"file_path": {"type": "string"}}, ["file_path"])},
    {"name": "powerpoint_local_plan_deck", "description": "Create a deterministic local deck plan and map slide roles to a local template.", "inputSchema": schema({"source_path": {"type": "string"}, "template_path": {"type": "string"}, "max_slides": {"type": "integer", "minimum": 3, "maximum": 20}}, ["source_path"])},
    {"name": "powerpoint_local_suggest_mermaid", "description": "Suggest Mermaid source for eligible planned architecture, process, and timeline slides. It does not render online.", "inputSchema": schema({"source_path": {"type": "string"}, "plan": {"type": "object"}}, ["source_path"])},
    {"name": "powerpoint_local_create_mermaid_diagram", "description": "Call the registered local Mermaid Studio MCP over stdio to create a local PNG diagram artifact for a planned slide. No remote calls are made by this server.", "inputSchema": schema({"title": {"type": "string", "maxLength": 120}, "source": {"type": "string", "maxLength": 200000}, "theme": {"type": "string", "enum": ["default", "dark", "forest", "neutral", "base"]}, "scale": {"type": "integer", "minimum": 1, "maximum": 4}}, ["title", "source"])},
    {"name": "powerpoint_local_prepare_deck", "description": "Create an editable local template copy and deck manifest for guided authoring, then optionally open it in PowerPoint.", "inputSchema": schema({"plan": {"type": "object"}, "output_path": {"type": "string"}, "template_path": {"type": "string"}, "open_after": {"type": "boolean"}}, ["plan", "output_path"])},
    {"name": "powerpoint_local_open", "description": "Open a local PowerPoint file with macOS Automation.", "inputSchema": schema({"file_path": {"type": "string"}}, ["file_path"])},
    {"name": "powerpoint_local_export_pngs", "description": "Ask local Microsoft PowerPoint to export a local deck as PNG slides.", "inputSchema": schema({"file_path": {"type": "string"}, "output_dir": {"type": "string"}}, ["file_path", "output_dir"])},
    {"name": "powerpoint_local_capture_screen", "description": "Capture a local PowerPoint review screenshot after bringing PowerPoint forward.", "inputSchema": schema({"output_path": {"type": "string"}}, ["output_path"])},
]


def status() -> dict[str, Any]:
    return {"server": SERVER_NAME, "version": VERSION, "local_only": {"network_calls": False, "package_fetching": False, "cloud_api_access": False, "allowed_processes": ["/usr/bin/osascript", "/usr/sbin/screencapture", "/usr/bin/open", "/opt/homebrew/bin/node (fixed local Mermaid MCP bridge)"]}, "powerpoint": powerpoint_version(), "mermaid_bridge": bridge_status(), "template_roots": template_roots(), "metadata_roots": [str(path) for path in metadata_roots()], "permissions": {"automation": "Grant Codex Automation access to Microsoft PowerPoint if a tool fails.", "screen_recording": "Grant Codex Screen Recording access before screenshot capture."}}


HANDLERS: dict[str, Callable[..., dict[str, Any]]] = {
    "powerpoint_local_status": lambda **_: status(),
    "powerpoint_local_find_templates": lambda query="", limit=30, include_cache_references=False, **_: {**find_templates(query, limit), **({"cache_references": cache_references(query, limit)} if include_cache_references else {})},
    "powerpoint_local_inspect_template": lambda file_path, **_: inspect_template(file_path),
    "powerpoint_local_read_source": lambda file_path, **_: read_source(file_path),
    "powerpoint_local_plan_deck": lambda source_path, template_path=None, max_slides=10, **_: create_plan(source_path, template_path, max_slides),
    "powerpoint_local_suggest_mermaid": lambda source_path, plan=None, **_: suggest_mermaid(source_path, plan),
    "powerpoint_local_create_mermaid_diagram": lambda title, source, theme="default", scale=4, **_: create_diagram(title, source, theme, scale),
    "powerpoint_local_prepare_deck": lambda plan, output_path, template_path=None, open_after=True, **_: prepare_guided_deck(plan, output_path, template_path, open_after),
    "powerpoint_local_open": lambda file_path, **_: open_in_powerpoint(file_path),
    "powerpoint_local_export_pngs": lambda file_path, output_dir, **_: export_pngs(file_path, output_dir),
    "powerpoint_local_capture_screen": lambda output_path, **_: capture_screen(output_path),
}


def tool_result(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    try:
        output = HANDLERS[name](**arguments)
        return {"content": [{"type": "text", "text": json.dumps(output, indent=2)}], "structuredContent": output}
    except (KeyError, TypeError, ValueError, RuntimeError, OSError) as error:
        return {"isError": True, "content": [{"type": "text", "text": f"Error: {error}"}]}


def respond(request_id: Any, result: dict[str, Any] | None = None, error: dict[str, Any] | None = None) -> None:
    message: dict[str, Any] = {"jsonrpc": "2.0", "id": request_id}
    if error:
        message["error"] = error
    else:
        message["result"] = result
    sys.stdout.write(json.dumps(message, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def serve() -> None:
    for line in sys.stdin:
        try:
            request = json.loads(line)
            request_id = request.get("id")
            method = request.get("method")
            params = request.get("params", {})
            if method == "initialize":
                if request_id is not None:
                    respond(request_id, {"protocolVersion": params.get("protocolVersion", "2025-03-26"), "capabilities": {"tools": {"listChanged": False}}, "serverInfo": {"name": SERVER_NAME, "version": VERSION}})
            elif method == "tools/list":
                respond(request_id, {"tools": TOOLS})
            elif method == "tools/call":
                respond(request_id, tool_result(params.get("name", ""), params.get("arguments", {})))
            elif method == "ping":
                respond(request_id, {})
            elif method and method.startswith("notifications/"):
                continue
            else:
                respond(request_id, error={"code": -32601, "message": f"Method not found: {method}"})
        except json.JSONDecodeError:
            continue
        except Exception as error:  # Never leak a traceback onto stdio.
            if 'request_id' in locals() and request_id is not None:
                respond(request_id, error={"code": -32603, "message": f"Internal local server error: {error}"})


if __name__ == "__main__":
    serve()
