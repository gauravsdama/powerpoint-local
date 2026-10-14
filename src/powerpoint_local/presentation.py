from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .local_files import POWERPOINT_EXTENSIONS, ensure_local_file

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}
MINIMUM_POWERPOINT_VERSION = "16.95"
MAX_POWERPOINT_PACKAGE_BYTES = 512 * 1024 * 1024
MAX_POWERPOINT_PACKAGE_MEMBERS = 10_000
MAX_INSPECTION_XML_BYTES = 4 * 1024 * 1024
MAX_TOTAL_INSPECTION_XML_BYTES = 32 * 1024 * 1024
MAX_INSPECTED_LAYOUTS = 500
MAX_INSPECTED_SLIDES = 1000


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
    if path.stat().st_size > MAX_POWERPOINT_PACKAGE_BYTES:
        raise ValueError(f"PowerPoint package exceeds the {MAX_POWERPOINT_PACKAGE_BYTES}-byte inspection limit.")
    layouts: list[dict] = []
    slides: list[dict] = []
    try:
        archive_context = zipfile.ZipFile(path)
    except zipfile.BadZipFile as error:
        raise ValueError(f"Invalid PowerPoint ZIP package: {path}") from error
    with archive_context as archive:
        members = archive.infolist()
        if len(members) > MAX_POWERPOINT_PACKAGE_MEMBERS:
            raise ValueError("PowerPoint package contains too many ZIP members.")
        selected = {
            member.filename: member
            for member in members
            if (member.filename.startswith("ppt/slideLayouts/slideLayout") and member.filename.endswith(".xml"))
            or re.fullmatch(r"ppt/slides/slide\d+\.xml", member.filename)
        }
        layout_names = [name for name in selected if name.startswith("ppt/slideLayouts/slideLayout")]
        slide_names = [name for name in selected if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)]
        if len(layout_names) > MAX_INSPECTED_LAYOUTS or len(slide_names) > MAX_INSPECTED_SLIDES:
            raise ValueError("PowerPoint package exceeds the layout or slide inspection count limit.")
        total_xml_bytes = sum(member.file_size for member in selected.values())
        if total_xml_bytes > MAX_TOTAL_INSPECTION_XML_BYTES:
            raise ValueError("PowerPoint inspection XML exceeds the total uncompressed size limit.")

        def read_xml(name: str) -> bytes:
            member = selected[name]
            if member.file_size > MAX_INSPECTION_XML_BYTES:
                raise ValueError(f"PowerPoint XML member exceeds the inspection limit: {name}")
            with archive.open(member) as source:
                content = source.read(MAX_INSPECTION_XML_BYTES + 1)
            if len(content) > MAX_INSPECTION_XML_BYTES:
                raise ValueError(f"PowerPoint XML member exceeds the inspection limit: {name}")
            return content

        for name in sorted(layout_names, key=_natural):
            try:
                root = ET.fromstring(read_xml(name))
            except ET.ParseError as error:
                raise ValueError(f"PowerPoint package contains invalid layout XML: {name}") from error
            common = root.find("p:cSld", NS)
            layout_name = common.get("name", Path(name).stem) if common is not None else Path(name).stem
            placeholders = [element.get("type", "obj") for element in root.findall(".//p:ph", NS)]
            layouts.append({"name": layout_name, "file": name, "placeholders": placeholders, "role": _layout_role(layout_name, placeholders)})
        for name in sorted(slide_names, key=_natural):
            try:
                root = ET.fromstring(read_xml(name))
            except ET.ParseError as error:
                raise ValueError(f"PowerPoint package contains invalid slide XML: {name}") from error
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
    completed = _run_local_command(
        command,
        timeout=60,
        timeout_message="PowerPoint automation timed out. Resolve any open permission or Office dialog, then retry.",
    )
    if completed.returncode:
        detail = completed.stderr.strip() or completed.stdout.strip() or "Unknown AppleScript error"
        raise RuntimeError(f"PowerPoint automation failed: {detail}. Enable Automation permission for Codex, then retry.")
    return completed.stdout.strip()


def _run_local_command(command: list[str], timeout: int, timeout_message: str) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(timeout_message) from error


def powerpoint_version() -> dict:
    try:
        version = _run_osascript(['tell application "Microsoft PowerPoint" to return version'])
        installed_parts = tuple(int(piece) for piece in re.findall(r"\d+", version)[:3])
        minimum_parts = tuple(int(piece) for piece in MINIMUM_POWERPOINT_VERSION.split("."))
        installed_parts += (0,) * (len(minimum_parts) - len(installed_parts))
        return {
            "installed": True,
            "version": version,
            "automation": "available",
            "minimum_supported_version": MINIMUM_POWERPOINT_VERSION,
            "supported": installed_parts >= minimum_parts,
        }
    except RuntimeError as error:
        return {
            "installed": Path("/Applications/Microsoft PowerPoint.app").exists(),
            "automation": "unavailable",
            "minimum_supported_version": MINIMUM_POWERPOINT_VERSION,
            "supported": False,
            "detail": str(error),
        }


def open_in_powerpoint(path_value: str) -> dict:
    path = ensure_local_file(path_value, POWERPOINT_EXTENSIONS)
    completed = _run_local_command(
        ["/usr/bin/open", "-a", "Microsoft PowerPoint", str(path)],
        timeout=30,
        timeout_message="PowerPoint launch timed out. Resolve any open Office dialog, then retry.",
    )
    if completed.returncode:
        detail = completed.stderr.strip() or "macOS could not launch Microsoft PowerPoint."
        raise RuntimeError(f"PowerPoint launch failed: {detail}")
    time.sleep(2)
    _run_osascript(['tell application "Microsoft PowerPoint" to activate'])
    return {"opened": str(path), "note": "Opened locally in Microsoft PowerPoint."}


def _close_presentation(name: str) -> None:
    _run_osascript([
        'on run argv',
        'tell application "Microsoft PowerPoint"',
        'if exists presentation (item 1 of argv) then close presentation (item 1 of argv) saving no',
        'end tell',
        'end run',
    ], [name])


def _create_empty_presentation(output: Path, open_after: bool) -> None:
    _run_osascript([
        "on run argv",
        'tell application "Microsoft PowerPoint"',
        "set deck to make new presentation",
        "save deck in (POSIX file (item 1 of argv)) as save as Open XML presentation",
        'if (item 2 of argv) is "false" then close active presentation saving no',
        "end tell",
        "end run",
    ], [str(output), "true" if open_after else "false"])


def prepare_guided_deck(plan: dict, output_path_value: str, template_path_value: str | None = None, open_after: bool = True) -> dict:
    output = Path(output_path_value).expanduser().resolve()
    if output.suffix.lower() != ".pptx":
        raise ValueError("output_path must end in .pptx.")
    manifest = output.with_suffix(".deck-manifest.json")
    if output.exists() or manifest.exists():
        raise ValueError("Refusing to overwrite an existing deck or manifest; choose a new output_path.")
    template = ensure_local_file(template_path_value, POWERPOINT_EXTENSIONS) if template_path_value else None
    if template and template.suffix.lower() != ".pptx":
        raise ValueError(
            "Guided deck preparation accepts only .pptx templates. "
            "Convert .potx explicitly in PowerPoint; macro-enabled .pptm/.potm policy must be chosen by the owner."
        )
    artifact: dict[str, str] = {
        "deck_path": str(output),
        "creation_mode": "local_template_copy" if template else "local_blank_powerpoint",
        "template": str(template) if template else "none",
    }
    manifest_text = json.dumps({"plan": plan, "artifact": artifact, "instructions": [
        "Use the mapped template role for each planned slide.",
        "Replace template placeholders; do not paste long source paragraphs unchanged.",
        "Export PNGs and run a visual review before sharing the deck.",
    ]}, indent=2)
    output.parent.mkdir(parents=True, exist_ok=True)
    transaction_id = uuid.uuid4().hex
    staged_deck = output.with_name(f"{output.stem}.powerpoint-local-staging-{transaction_id}.pptx")
    staged_manifest = manifest.with_name(f"{manifest.stem}.powerpoint-local-staging-{transaction_id}.json")
    try:
        if template:
            shutil.copy2(template, staged_deck)
        else:
            _create_empty_presentation(staged_deck, False)
        if not staged_deck.is_file():
            raise RuntimeError("PowerPoint did not create the requested local deck.")
        staged_manifest.write_text(manifest_text, encoding="utf-8")
        staged_deck.rename(output)
        try:
            staged_manifest.rename(manifest)
        except OSError:
            output.unlink(missing_ok=True)
            raise
    finally:
        staged_deck.unlink(missing_ok=True)
        staged_manifest.unlink(missing_ok=True)
    open_status = "not_requested"
    open_error = None
    if open_after:
        try:
            open_in_powerpoint(str(output))
            open_status = "opened"
        except RuntimeError as error:
            open_status = "failed"
            open_error = str(error)
    if template:
        note = "The plan and local template copy are ready for guided authoring."
    else:
        note = "The plan and a local blank PowerPoint deck are ready for guided authoring."
    return {
        **artifact,
        "manifest_path": str(manifest),
        "owned_artifacts": [str(output), str(manifest)],
        "overwrite_policy": "existing deck or manifest paths are rejected",
        "open_status": open_status,
        **({"open_error": open_error} if open_error else {}),
        "note": f"{note} No network template retrieval was attempted.",
    }


def export_pngs(path_value: str, output_dir_value: str) -> dict:
    path = ensure_local_file(path_value, POWERPOINT_EXTENSIONS)
    output_dir = Path(output_dir_value).expanduser().resolve()
    if output_dir.exists():
        raise ValueError("Refusing to merge with or overwrite an existing output_dir; choose a new directory.")
    output_dir.parent.mkdir(parents=True, exist_ok=True)
    open_in_powerpoint(str(path))
    transaction_name = f"{output_dir.name}.powerpoint-local-staging-{uuid.uuid4().hex}"
    staged_output = output_dir.with_name(transaction_name)
    staged_output.mkdir()
    temporary_paths: list[Path] = []
    try:
        image_names: list[str] = []
        direct_stem = output_dir.parent / f"{transaction_name}-direct.png"
        _run_osascript([
            'on run argv',
            'tell application "Microsoft PowerPoint"',
            'set deck to presentation (item 3 of argv)',
            'save deck in (POSIX file (item 1 of argv)) as save as PNG',
            'end tell',
            'end run',
        ], [str(direct_stem), str(path), path.name])
        direct_roots = list(output_dir.parent.glob(f"{transaction_name}-direct*"))
        temporary_paths.extend(direct_roots)
        direct_images: list[Path] = []
        for direct_root in direct_roots:
            if direct_root.is_file() and direct_root.suffix.lower() == ".png":
                direct_images.append(direct_root)
            elif direct_root.is_dir():
                direct_images.extend(direct_root.rglob("*.png"))
        for slide_number, direct_image in enumerate(sorted(direct_images, key=lambda item: _natural(str(item))), start=1):
            image_name = f"Slide-{slide_number}.png"
            shutil.copy2(direct_image, staged_output / image_name)
            image_names.append(image_name)
        export_mode = "powerpoint_png"
        if not image_names:
            export_mode = "powerpoint_pdf_single_slide_fallback"
            slide_count = len(inspect_template(str(path))["slides"])
            if slide_count == 0:
                raise RuntimeError("The presentation contains no slides to export.")
            for slide_number in range(1, slide_count + 1):
                working_copy = output_dir.parent / f"{transaction_name}-slide-{slide_number}.pptx"
                pdf_path = output_dir.parent / f"{transaction_name}-slide-{slide_number}.pdf"
                temporary_paths.extend((working_copy, pdf_path))
                image_name = f"Slide-{slide_number}.png"
                png_path = staged_output / image_name
                shutil.copy2(path, working_copy)
                completed = _run_local_command(
                    ["/usr/bin/open", "-a", "Microsoft PowerPoint", str(working_copy)],
                    timeout=30,
                    timeout_message="PowerPoint launch timed out during PNG fallback.",
                )
                if completed.returncode:
                    raise RuntimeError(f"PowerPoint launch failed during PNG fallback: {completed.stderr.strip()}")
                time.sleep(1)
                try:
                    _run_osascript([
                        "on run argv",
                        'tell application "Microsoft PowerPoint"',
                        "set deck to presentation (item 4 of argv)",
                        "set keepSlide to (item 2 of argv) as integer",
                        "set slideTotal to (item 3 of argv) as integer",
                        "repeat with slideNumber from slideTotal to 1 by -1",
                        "if slideNumber is not keepSlide then delete slide slideNumber of deck",
                        "end repeat",
                        "save deck in (POSIX file (item 1 of argv)) as save as PDF",
                        "end tell",
                        "end run",
                    ], [str(pdf_path), str(slide_number), str(slide_count), working_copy.name])
                finally:
                    _close_presentation(working_copy.name)
                converted = _run_local_command(
                    ["/usr/bin/sips", "-s", "format", "png", str(pdf_path), "--out", str(png_path)],
                    timeout=60,
                    timeout_message=f"PNG conversion timed out for slide {slide_number}.",
                )
                if converted.returncode or not png_path.is_file():
                    detail = converted.stderr.strip() or converted.stdout.strip() or "sips produced no image"
                    raise RuntimeError(f"PowerPoint PDF fallback could not render slide {slide_number}: {detail}")
                image_names.append(image_name)
        staged_output.rename(output_dir)
    finally:
        if staged_output.exists():
            shutil.rmtree(staged_output)
        for temporary_path in temporary_paths:
            if temporary_path.is_symlink():
                temporary_path.unlink(missing_ok=True)
            elif temporary_path.is_dir():
                shutil.rmtree(temporary_path)
            else:
                temporary_path.unlink(missing_ok=True)
    images = [str(output_dir / name) for name in image_names]
    return {
        "source": str(path),
        "output_dir": str(output_dir),
        "images": images,
        "count": len(images),
        "export_mode": export_mode,
        "overwrite_policy": "output_dir must not already exist",
        "note": "This tool owns the newly created output directory. Inspect these local PNGs in Codex for visual review.",
    }


def capture_screen(output_path_value: str) -> dict:
    output = Path(output_path_value).expanduser().resolve()
    if output.suffix.lower() != ".png":
        raise ValueError("output_path must end in .png.")
    if output.exists():
        raise ValueError("Refusing to overwrite an existing screenshot; choose a new output_path.")
    output.parent.mkdir(parents=True, exist_ok=True)
    _run_osascript(['tell application "Microsoft PowerPoint" to activate'])
    with tempfile.TemporaryDirectory(prefix="powerpoint-local-capture-", dir=output.parent) as directory:
        staged_output = Path(directory) / output.name
        completed = _run_local_command(
            ["/usr/sbin/screencapture", "-x", str(staged_output)],
            timeout=30,
            timeout_message="Local screenshot capture timed out.",
        )
        if completed.returncode:
            detail = completed.stderr.strip() or "Screen Recording permission may be disabled."
            raise RuntimeError(f"Local screenshot capture failed: {detail}")
        if not staged_output.is_file():
            raise RuntimeError("Screen capture reported success but produced no PNG artifact.")
        staged_output.rename(output)
    return {
        "screenshot": str(output),
        "owned_artifacts": [str(output)],
        "overwrite_policy": "existing output paths are rejected",
        "note": "Local full-screen capture. Inspect it in Codex for a visual review.",
    }
