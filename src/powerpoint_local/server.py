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
MAX_REQUEST_BYTES = 1024 * 1024


def _reject_json_constant(value: str) -> None:
    raise json.JSONDecodeError(f"Non-standard JSON constant: {value}", value, 0)


def schema(properties: dict[str, Any] | None = None, required: list[str] | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {"type": "object", "properties": properties or {}, "additionalProperties": False}
    if required:
        value["required"] = required
    return value


TOOLS: list[dict[str, Any]] = [
    {"name": "powerpoint_local_status", "description": "Report local PowerPoint, local-only policy, and searched cache locations.", "inputSchema": schema()},
    {"name": "powerpoint_local_find_templates", "description": "Find local PowerPoint templates. Oracle candidates sort first; cache references are metadata only.", "inputSchema": schema({"query": {"type": "string"}, "limit": {"type": "integer", "minimum": 1, "maximum": 100}, "include_cache_references": {"type": "boolean"}})},
    {"name": "powerpoint_local_inspect_template", "description": "Inventory local PPTX/PPTM/POTX/POTM layouts, placeholder roles, and slide text.", "inputSchema": schema({"file_path": {"type": "string", "description": "Absolute local PowerPoint package path."}}, ["file_path"])},
    {"name": "powerpoint_local_read_source", "description": "Read a local Markdown, text, or DOCX source document into sections.", "inputSchema": schema({"file_path": {"type": "string"}}, ["file_path"])},
    {"name": "powerpoint_local_plan_deck", "description": "Create a deterministic local deck plan and map slide roles to a local template.", "inputSchema": schema({"source_path": {"type": "string"}, "template_path": {"type": "string"}, "max_slides": {"type": "integer", "minimum": 3, "maximum": 20}}, ["source_path"])},
    {"name": "powerpoint_local_suggest_mermaid", "description": "Suggest Mermaid source for eligible planned architecture, process, and timeline slides. It does not render online.", "inputSchema": schema({"source_path": {"type": "string"}, "plan": {"type": "object"}}, ["source_path"])},
    {"name": "powerpoint_local_create_mermaid_diagram", "description": "Call the registered local Mermaid Studio MCP over stdio to create a local PNG diagram artifact for a planned slide. No remote calls are made by this server.", "inputSchema": schema({"title": {"type": "string", "maxLength": 120}, "source": {"type": "string", "maxLength": 200000}, "theme": {"type": "string", "enum": ["default", "dark", "forest", "neutral", "base"]}, "scale": {"type": "integer", "minimum": 1, "maximum": 4}}, ["title", "source"])},
    {"name": "powerpoint_local_prepare_deck", "description": "Create a new .pptx and manifest without overwriting paths. Supplied templates must already be .pptx; template and macro packages are not byte-renamed.", "inputSchema": schema({"plan": {"type": "object"}, "output_path": {"type": "string"}, "template_path": {"type": "string"}, "open_after": {"type": "boolean"}}, ["plan", "output_path"])},
    {"name": "powerpoint_local_open", "description": "Open a local PowerPoint file with macOS Automation.", "inputSchema": schema({"file_path": {"type": "string"}}, ["file_path"])},
    {"name": "powerpoint_local_export_pngs", "description": "Create a new output directory containing only PNGs from this export. Existing output paths are rejected.", "inputSchema": schema({"file_path": {"type": "string"}, "output_dir": {"type": "string"}}, ["file_path", "output_dir"])},
    {"name": "powerpoint_local_capture_screen", "description": "Create a new local PowerPoint review screenshot after bringing PowerPoint forward. Existing output paths are rejected.", "inputSchema": schema({"output_path": {"type": "string"}}, ["output_path"])},
]

READ_ONLY_TOOLS = {
    "powerpoint_local_status",
    "powerpoint_local_find_templates",
    "powerpoint_local_inspect_template",
    "powerpoint_local_read_source",
    "powerpoint_local_plan_deck",
    "powerpoint_local_suggest_mermaid",
}
for tool in TOOLS:
    read_only = tool["name"] in READ_ONLY_TOOLS
    tool["annotations"] = {
        "readOnlyHint": read_only,
        "destructiveHint": False,
        "idempotentHint": read_only or tool["name"] == "powerpoint_local_open",
        "openWorldHint": False,
    }


def status() -> dict[str, Any]:
    return {"server": SERVER_NAME, "version": VERSION, "local_only": {"network_calls": False, "package_fetching": False, "cloud_api_access": False, "allowed_processes": ["/usr/bin/osascript", "/usr/sbin/screencapture", "/usr/bin/open", "/usr/bin/sips", "configured local Node executable (Mermaid MCP bridge)"]}, "powerpoint": powerpoint_version(), "mermaid_bridge": bridge_status(), "template_roots": template_roots(), "metadata_roots": [str(path) for path in metadata_roots()], "permissions": {"automation": "Grant Codex Automation access to Microsoft PowerPoint if a tool fails.", "screen_recording": "Grant Codex Screen Recording access before screenshot capture."}}


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


def response_message(request_id: Any, result: dict[str, Any] | None = None, error: dict[str, Any] | None = None) -> dict[str, Any]:
    message: dict[str, Any] = {"jsonrpc": "2.0", "id": request_id}
    if error is not None:
        message["error"] = error
    else:
        message["result"] = result
    return message


def respond(message: dict[str, Any] | list[dict[str, Any]]) -> None:
    sys.stdout.write(json.dumps(message, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def _valid_request_id(value: Any) -> bool:
    return value is None or (isinstance(value, (str, int, float)) and not isinstance(value, bool))


def _process_request(request: Any) -> dict[str, Any] | None:
    if not isinstance(request, dict):
        return response_message(None, error={"code": -32600, "message": "Invalid Request"})
    request_id = request.get("id") if "id" in request else None
    if (
        request.get("jsonrpc") != "2.0"
        or not isinstance(request.get("method"), str)
        or ("id" in request and not _valid_request_id(request_id))
    ):
        return response_message(None, error={"code": -32600, "message": "Invalid Request"})
    method = request["method"]
    params = request.get("params", {})
    is_notification = "id" not in request
    if not isinstance(params, dict):
        response = response_message(request_id, error={"code": -32602, "message": "Invalid params"})
        return None if is_notification else response
    response: dict[str, Any]
    if method == "initialize":
        response = response_message(request_id, {"protocolVersion": params.get("protocolVersion", "2025-03-26"), "capabilities": {"tools": {"listChanged": False}}, "serverInfo": {"name": SERVER_NAME, "version": VERSION}})
    elif method == "tools/list":
        response = response_message(request_id, {"tools": TOOLS})
    elif method == "tools/call":
        name = params.get("name")
        arguments = params.get("arguments", {})
        if not isinstance(name, str) or not isinstance(arguments, dict):
            response = response_message(request_id, error={"code": -32602, "message": "Invalid params"})
        else:
            response = response_message(request_id, tool_result(name, arguments))
    elif method == "ping":
        response = response_message(request_id, {})
    elif is_notification and method.startswith("notifications/"):
        return None
    else:
        response = response_message(request_id, error={"code": -32601, "message": f"Method not found: {method}"})
    return None if is_notification else response


def _process_message(message: Any) -> dict[str, Any] | list[dict[str, Any]] | None:
    if isinstance(message, list):
        if not message:
            return response_message(None, error={"code": -32600, "message": "Invalid Request"})
        responses = [response for item in message if (response := _process_request(item)) is not None]
        return responses or None
    return _process_request(message)


def serve() -> None:
    stream = sys.stdin.buffer
    while line := stream.readline(MAX_REQUEST_BYTES + 1):
        oversized = len(line) > MAX_REQUEST_BYTES
        if oversized and not line.endswith(b"\n"):
            while continuation := stream.readline(MAX_REQUEST_BYTES + 1):
                if continuation.endswith(b"\n"):
                    break
        if oversized:
            respond(response_message(None, error={"code": -32600, "message": "Request exceeds the 1048576-byte limit"}))
            continue
        try:
            message = json.loads(line.decode("utf-8"), parse_constant=_reject_json_constant)
            response = _process_message(message)
            if response is not None:
                respond(response)
        except (UnicodeDecodeError, json.JSONDecodeError):
            respond(response_message(None, error={"code": -32700, "message": "Parse error"}))
        except Exception as error:  # Never leak a traceback onto stdio.
            respond(response_message(None, error={"code": -32603, "message": f"Internal local server error: {error}"}))


if __name__ == "__main__":
    serve()
