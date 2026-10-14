from __future__ import annotations

import json
import os
import re
import selectors
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any

DEFAULT_MERMAID_GATEWAY = "http://127.0.0.1:8787"
MERMAID_RESPONSE_TIMEOUT_SECONDS = 90.0
MERMAID_SHUTDOWN_TIMEOUT_SECONDS = 3.0
MAX_MERMAID_STDERR_BYTES = 8192
MAX_MERMAID_RESPONSE_BYTES = 2 * 1024 * 1024


def bridge_config() -> tuple[Path, Path, str]:
    node_value = os.environ.get("POWERPOINT_LOCAL_MERMAID_NODE") or shutil.which("node") or "/usr/bin/node"
    default_entry = Path(__file__).resolve().parents[2].parent / "mermaid-studio" / "dist" / "mcp" / "index.js"
    entry_value = os.environ.get("POWERPOINT_LOCAL_MERMAID_ENTRY") or str(default_entry)
    gateway = os.environ.get("POWERPOINT_LOCAL_MERMAID_GATEWAY", DEFAULT_MERMAID_GATEWAY)
    match = re.fullmatch(r"http://(?:127\.0\.0\.1|localhost):([0-9]{1,5})/?", gateway)
    if match is None or not 1 <= int(match.group(1)) <= 65535:
        raise ValueError("POWERPOINT_LOCAL_MERMAID_GATEWAY must use localhost or 127.0.0.1.")
    return Path(node_value).expanduser(), Path(entry_value).expanduser(), gateway


def bridge_status() -> dict[str, Any]:
    node, entry, gateway = bridge_config()
    return {
        "transport": "local_stdio_child_process",
        "node": str(node),
        "entry": str(entry),
        "available": node.is_file() and entry.is_file(),
        "gateway": gateway,
        "boundary": "PowerPoint Local uses only stdio. Mermaid Studio owns its loopback renderer connection.",
    }


def _read_response(process: subprocess.Popen[str], request_id: int, timeout_seconds: float) -> dict[str, Any]:
    if process.stdout is None or process.stderr is None:
        raise RuntimeError("Mermaid MCP did not expose stdout and stderr.")
    deadline = time.monotonic() + timeout_seconds
    stdout_buffer = bytearray()
    stderr_tail = bytearray()
    with selectors.DefaultSelector() as selector:
        selector.register(process.stdout, selectors.EVENT_READ, "stdout")
        selector.register(process.stderr, selectors.EVENT_READ, "stderr")
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            events = selector.select(remaining)
            if not events:
                break
            for key, _ in events:
                chunk = os.read(key.fileobj.fileno(), 65_536)
                if not chunk:
                    selector.unregister(key.fileobj)
                    continue
                if key.data == "stderr":
                    stderr_tail.extend(chunk)
                    if len(stderr_tail) > MAX_MERMAID_STDERR_BYTES:
                        del stderr_tail[:-MAX_MERMAID_STDERR_BYTES]
                    continue
                stdout_buffer.extend(chunk)
                if len(stdout_buffer) > MAX_MERMAID_RESPONSE_BYTES:
                    raise RuntimeError("Mermaid MCP response exceeded the 2097152-byte limit.")
                while b"\n" in stdout_buffer:
                    raw_line, _, remainder = stdout_buffer.partition(b"\n")
                    stdout_buffer = bytearray(remainder)
                    try:
                        message = json.loads(raw_line.decode("utf-8"))
                    except (UnicodeDecodeError, json.JSONDecodeError) as error:
                        raise RuntimeError("Mermaid MCP emitted an invalid JSON-RPC message.") from error
                    if message.get("id") == request_id:
                        if "error" in message:
                            detail = message["error"].get("message", "unknown error")
                            raise RuntimeError(f"Mermaid MCP error: {detail}")
                        if "result" not in message:
                            raise RuntimeError("Mermaid MCP response did not contain a result.")
                        return message["result"]
            if process.poll() is not None and not selector.get_map():
                break
    detail = stderr_tail.decode("utf-8", errors="replace").strip()
    suffix = f" Child stderr: {detail}" if detail else ""
    raise RuntimeError(
        "Mermaid MCP did not respond before the deadline. "
        "Start Mermaid Studio locally with `npm run start`, then retry." + suffix
    )


def create_diagram(title: str, source: str, theme: str = "default", scale: int = 4) -> dict[str, Any]:
    node, entry, gateway = bridge_config()
    if not node.is_file() or not entry.is_file():
        raise RuntimeError(
            "Local Mermaid Studio MCP is unavailable. Set POWERPOINT_LOCAL_MERMAID_NODE "
            "and POWERPOINT_LOCAL_MERMAID_ENTRY, or place mermaid-studio beside this checkout."
        )
    if not title or len(title) > 120:
        raise ValueError("title must contain 1–120 characters.")
    if not source or len(source) > 200_000:
        raise ValueError("source must contain 1–200,000 characters.")
    if theme not in {"default", "dark", "forest", "neutral", "base"}:
        raise ValueError("theme must be one of: default, dark, forest, neutral, base.")
    if not isinstance(scale, int) or not 1 <= scale <= 4:
        raise ValueError("scale must be an integer from 1 to 4.")
    environment = {"PATH": os.environ.get("PATH", "/usr/bin:/bin"), "MCP_API_BASE_URL": gateway}
    process = subprocess.Popen([str(node), str(entry)], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=environment)
    try:
        if process.stdin is None:
            raise RuntimeError("Mermaid MCP did not expose stdin.")
        process.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-03-26", "capabilities": {}, "clientInfo": {"name": "powerpoint-local", "version": "0.1.0"}}}) + "\n")
        process.stdin.flush()
        _read_response(process, 1, MERMAID_RESPONSE_TIMEOUT_SECONDS)
        process.stdin.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}) + "\n")
        process.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "mermaid_studio_create_diagram", "arguments": {"title": title, "source": source, "theme": theme, "scale": scale, "response_format": "json"}}}) + "\n")
        process.stdin.flush()
        result = _read_response(process, 2, MERMAID_RESPONSE_TIMEOUT_SECONDS)
        if result.get("isError"):
            message = result.get("content", [{}])[0].get("text", "Mermaid Studio could not create the diagram.")
            raise RuntimeError(message)
        artifact = result.get("structuredContent", {}).get("artifact")
        if not isinstance(artifact, dict) or not artifact.get("pngPath"):
            raise RuntimeError("Mermaid Studio returned no local PNG artifact path.")
        return {"artifact": artifact, "bridge": bridge_status()}
    finally:
        if process.stdin is not None:
            try:
                process.stdin.close()
            except OSError:
                pass
        if process.poll() is None:
            process.terminate()
        try:
            process.wait(timeout=MERMAID_SHUTDOWN_TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=MERMAID_SHUTDOWN_TIMEOUT_SECONDS)
        for pipe in (process.stdout, process.stderr):
            if pipe is not None:
                pipe.close()
