from __future__ import annotations

import json
import re
import shutil
import subprocess
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .local_files import POWERPOINT_EXTENSIONS, ensure_local_file

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}


def _natural(value: str) -> list[object]:
    return [int(piece) if piece.isdigit() else piece.lower() for piece in re.split(r"(\d+)", value)]


def _layout_role(name: str, placeholders: list[str]) -> str:
    text = f"{name} {' '.join(placeholders)}".lower()
    if "title slide" in text or "center title" in text:
        return "title"
    if "section" in text:
        return "section"
    if "comparison" in text:
        return "comparison"
    if "two" in text and ("content" in text or "object" in text):
        return "two_column"
    if "picture" in text or "image" in text:
        return "image"
    if "blank" in text:
        return "diagram"
    if "title" in text and "only" in text:
        return "section"
    return "content"


def inspect_template(path_value: str) -> dict:
    path = ensure_local_file(path_value, POWERPOINT_EXTENSIONS)
    layouts: list[dict] = []
    slides: list[dict] = []
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        for name in sorted((item for item in names if item.startswith("ppt/slideLayouts/slideLayout") and item.endswith(".xml")), key=_natural):
            root = ET.fromstring(archive.read(name))
            common = root.find("p:cSld", NS)
            layout_name = common.get("name", Path(name).stem) if common is not None else Path(name).stem
            placeholders = [element.get("type", "obj") for element in root.findall(".//p:ph", NS)]
            layouts.append({"name": layout_name, "file": name, "placeholders": placeholders, "role": _layout_role(layout_name, placeholders)})
        for name in sorted((item for item in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", item)), key=_natural):
            root = ET.fromstring(archive.read(name))
            text = [node.text or "" for node in root.findall(".//a:t", NS) if (node.text or "").strip()]
            slides.append({"file": name, "text": text[:30]})
    role_map: dict[str, list[str]] = {}
    for layout in layouts:
        role_map.setdefault(layout["role"], []).append(layout["name"])
    oracle_evidence = "oracle" in path.name.lower() or any("oracle" in layout["name"].lower() for layout in layouts)
    return {"path": str(path), "layout_count": len(layouts), "layouts": layouts, "slides": slides, "role_map": role_map, "oracle_candidate": oracle_evidence}


def _run_osascript(lines: list[str], arguments: list[str] | None = None) -> str:
    command = ["/usr/bin/osascript"]
    for line in lines:
        command.extend(["-e", line])
    command.extend(arguments or [])
    completed = subprocess.run(command, capture_output=True, text=True, timeout=60, check=False)
    if completed.returncode:
        detail = completed.stderr.strip() or completed.stdout.strip() or "Unknown AppleScript error"
        raise RuntimeError(f"PowerPoint automation failed: {detail}. Enable Automation permission for Codex, then retry.")
    return completed.stdout.strip()


def powerpoint_version() -> dict:
    try:
        version = _run_osascript(['tell application "Microsoft PowerPoint" to return version'])
        return {"installed": True, "version": version, "automation": "available"}
    except RuntimeError as error:
        return {"installed": Path("/Applications/Microsoft PowerPoint.app").exists(), "automation": "unavailable", "detail": str(error)}


def open_in_powerpoint(path_value: str) -> dict:
    path = ensure_local_file(path_value, POWERPOINT_EXTENSIONS)
    completed = subprocess.run(["/usr/bin/open", "-a", "Microsoft PowerPoint", str(path)], capture_output=True, text=True, timeout=30, check=False)
    if completed.returncode:
        detail = completed.stderr.strip() or "macOS could not launch Microsoft PowerPoint."
        raise RuntimeError(f"PowerPoint launch failed: {detail}")
    time.sleep(2)
    _run_osascript(['tell application "Microsoft PowerPoint" to activate'])
    return {"opened": str(path), "note": "Opened locally in Microsoft PowerPoint."}


def prepare_guided_deck(plan: dict, output_path_value: str, template_path_value: str | None = None, open_after: bool = True) -> dict:
    output = Path(output_path_value).expanduser().resolve()
    if output.suffix.lower() != ".pptx":
        raise ValueError("output_path must end in .pptx.")
    output.parent.mkdir(parents=True, exist_ok=True)
    manifest = output.with_suffix(".deck-manifest.json")
    template = ensure_local_file(template_path_value, POWERPOINT_EXTENSIONS) if template_path_value else None
    artifact: dict[str, str]
    if template:
        shutil.copy2(template, output)
        artifact = {"deck_path": str(output), "creation_mode": "local_template_copy", "template": str(template)}
        if open_after:
            open_in_powerpoint(str(output))
    else:
        _run_osascript(['tell application "Microsoft PowerPoint" to activate'])
        artifact = {"creation_mode": "guided_empty_presentation", "template": "none"}
    manifest.write_text(json.dumps({"plan": plan, "artifact": artifact, "instructions": [
        "Use the mapped template role for each planned slide.",
        "Replace template placeholders; do not paste long source paragraphs unchanged.",
        "Export PNGs and run a visual review before sharing the deck.",
    ]}, indent=2), encoding="utf-8")
    return {**artifact, "manifest_path": str(manifest), "note": "The plan and local copy are ready for guided authoring. No network template retrieval was attempted."}


def export_pngs(path_value: str, output_dir_value: str) -> dict:
    path = ensure_local_file(path_value, POWERPOINT_EXTENSIONS)
    output_dir = Path(output_dir_value).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = output_dir / "slide.png"
    open_in_powerpoint(str(path))
    _run_osascript([
        'on run argv',
        'tell application "Microsoft PowerPoint"',
        'set deck to presentation (item 3 of argv)',
        'save deck in (POSIX file (item 1 of argv)) as save as PNG',
        'end tell',
        'end run',
    ], [str(stem), str(path), path.name])
    images = sorted(str(item) for item in output_dir.glob("*.png"))
    if not images:
        raise RuntimeError("PowerPoint completed the PNG export command but produced no PNG files. Keep the deck open in PowerPoint, confirm it is writable, and retry after granting Automation permission.")
    return {"source": str(path), "output_dir": str(output_dir), "images": images, "count": len(images), "note": "Inspect these local PNGs in Codex for visual review."}


def capture_screen(output_path_value: str) -> dict:
    output = Path(output_path_value).expanduser().resolve()
    if output.suffix.lower() != ".png":
        raise ValueError("output_path must end in .png.")
    output.parent.mkdir(parents=True, exist_ok=True)
    _run_osascript(['tell application "Microsoft PowerPoint" to activate'])
    completed = subprocess.run(["/usr/sbin/screencapture", "-x", str(output)], capture_output=True, text=True, timeout=30, check=False)
    if completed.returncode:
        detail = completed.stderr.strip() or "Screen Recording permission may be disabled."
        raise RuntimeError(f"Local screenshot capture failed: {detail}")
    return {"screenshot": str(output), "note": "Local full-screen capture. Inspect it in Codex for a visual review."}
