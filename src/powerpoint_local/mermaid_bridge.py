from __future__ import annotations

import json
import os
import select
import subprocess
from pathlib import Path
from typing import Any

MERMAID_NODE = Path("/opt/homebrew/bin/node")
MERMAID_ENTRY = Path("/Users/gauravsdama/git/mermaid-studio/dist/mcp/index.js")
MERMAID_GATEWAY = "http://127.0.0.1:8787"


def bridge_status() -> dict[str, Any]:
    return {
        "transport": "local_stdio_child_process",
        "node": str(MERMAID_NODE),
        "entry": str(MERMAID_ENTRY),
        "available": MERMAID_NODE.is_file() and MERMAID_ENTRY.is_file(),
        "gateway": MERMAID_GATEWAY,
        "boundary": "PowerPoint Local uses only stdio. Mermaid Studio owns its loopback renderer connection.",
    }


def _read_response(process: subprocess.Popen[str], request_id: int, timeout_seconds: int = 90) -> dict[str, Any]:
    if process.stdout is None:
        raise RuntimeError("Mermaid MCP did not expose stdout.")
    remaining = float(timeout_seconds)
    while remaining > 0:
        ready, _, _ = select.select([process.stdout], [], [], min(1.0, remaining))
        remaining -= 1.0
        if not ready:
            continue
        line = process.stdout.readline()
        if not line:
            break
        message = json.loads(line)
        if message.get("id") == request_id:
            if "error" in message:
                raise RuntimeError(f"Mermaid MCP error: {message['error'].get('message', 'unknown error')}")
            return message["result"]
    raise RuntimeError("Mermaid MCP did not respond. Start Mermaid Studio locally with `npm run start`, then retry.")


def create_diagram(title: str, source: str, theme: str = "default", scale: int = 4) -> dict[str, Any]:
    if not MERMAID_NODE.is_file() or not MERMAID_ENTRY.is_file():
        raise RuntimeError("Local Mermaid Studio MCP is unavailable at its registered local path.")
    if not title or len(title) > 120:
        raise ValueError("title must contain 1–120 characters.")
    if not source or len(source) > 200_000:
        raise ValueError("source must contain 1–200,000 characters.")
    if theme not in {"default", "dark", "forest", "neutral", "base"}:
        raise ValueError("theme must be one of: default, dark, forest, neutral, base.")
    if not isinstance(scale, int) or not 1 <= scale <= 4:
        raise ValueError("scale must be an integer from 1 to 4.")
    environment = {"PATH": os.environ.get("PATH", "/usr/bin:/bin"), "MCP_API_BASE_URL": MERMAID_GATEWAY}
    process = subprocess.Popen([str(MERMAID_NODE), str(MERMAID_ENTRY)], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=environment)
    try:
        if process.stdin is None:
            raise RuntimeError("Mermaid MCP did not expose stdin.")
        process.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-03-26", "capabilities": {}, "clientInfo": {"name": "powerpoint-local", "version": "0.1.0"}}}) + "\n")
        process.stdin.flush()
        _read_response(process, 1)
        process.stdin.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}) + "\n")
        process.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "mermaid_studio_create_diagram", "arguments": {"title": title, "source": source, "theme": theme, "scale": scale, "response_format": "json"}}}) + "\n")
        process.stdin.flush()
        result = _read_response(process, 2)
        if result.get("isError"):
            message = result.get("content", [{}])[0].get("text", "Mermaid Studio could not create the diagram.")
            raise RuntimeError(message)
        artifact = result.get("structuredContent", {}).get("artifact")
        if not isinstance(artifact, dict) or not artifact.get("pngPath"):
            raise RuntimeError("Mermaid Studio returned no local PNG artifact path.")
        return {"artifact": artifact, "bridge": bridge_status()}
    finally:
        process.terminate()
        try:
            process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            process.kill()
