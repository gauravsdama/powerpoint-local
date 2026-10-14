from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .local_files import SOURCE_EXTENSIONS, ensure_local_file

MAX_SOURCE_BYTES = 250 * 1024
MAX_DOCX_PACKAGE_BYTES = 64 * 1024 * 1024
MAX_DOCX_XML_BYTES = 2 * 1024 * 1024
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _clean(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _read_bounded_file(path: Path) -> str:
    with path.open("rb") as source:
        content = source.read(MAX_SOURCE_BYTES + 1)
    if len(content) > MAX_SOURCE_BYTES:
        raise ValueError(f"Source file exceeds the {MAX_SOURCE_BYTES}-byte limit: {path}")
    return content.decode("utf-8", errors="ignore")


def _read_bounded_zip_member(archive: zipfile.ZipFile, name: str, limit: int) -> bytes:
    try:
        member = archive.getinfo(name)
    except KeyError as error:
        raise ValueError(f"DOCX is missing required member: {name}") from error
    if member.file_size > limit:
        raise ValueError(f"DOCX member {name} exceeds the {limit}-byte uncompressed limit.")
    with archive.open(member) as source:
        content = source.read(limit + 1)
    if len(content) > limit:
        raise ValueError(f"DOCX member {name} exceeds the {limit}-byte uncompressed limit.")
    return content


def _read_docx(path: Path) -> list[tuple[str, int | None]]:
    if path.stat().st_size > MAX_DOCX_PACKAGE_BYTES:
        raise ValueError(f"DOCX package exceeds the {MAX_DOCX_PACKAGE_BYTES}-byte limit.")
    try:
        with zipfile.ZipFile(path) as archive:
            if len(archive.infolist()) > 10_000:
                raise ValueError("DOCX contains too many ZIP members.")
            xml = _read_bounded_zip_member(archive, "word/document.xml", MAX_DOCX_XML_BYTES)
    except zipfile.BadZipFile as error:
        raise ValueError(f"Invalid DOCX ZIP package: {path}") from error
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as error:
        raise ValueError("DOCX document.xml contains invalid XML.") from error
    paragraphs: list[tuple[str, int | None]] = []
    for paragraph in root.findall(".//w:p", WORD_NS):
        text = _clean("".join(node.text or "" for node in paragraph.findall(".//w:t", WORD_NS)))
        if not text:
            continue
        style = paragraph.find("./w:pPr/w:pStyle", WORD_NS)
        style_name = style.get("{" + WORD_NS["w"] + "}val", "") if style is not None else ""
        match = re.search(r"heading\s*([1-9])", style_name, re.IGNORECASE)
        paragraphs.append((text, int(match.group(1)) if match else None))
    return paragraphs


def _read_markdown(path: Path) -> list[tuple[str, int | None]]:
    result: list[tuple[str, int | None]] = []
    for raw in _read_bounded_file(path).splitlines():
        match = re.match(r"^(#{1,6})\s+(.+)$", raw.strip())
        if match:
            result.append((_clean(match.group(2)), len(match.group(1))))
        elif raw.strip():
            result.append((_clean(re.sub(r"^[-*+]\s+", "", raw)), None))
    return result


def _read_text(path: Path) -> list[tuple[str, int | None]]:
    return [(_clean(value), None) for value in _read_bounded_file(path).splitlines() if _clean(value)]


def read_source(path_value: str) -> dict:
    path = ensure_local_file(path_value, SOURCE_EXTENSIONS)
    if path.suffix.lower() == ".docx":
        paragraphs = _read_docx(path)
    elif path.suffix.lower() == ".md":
        paragraphs = _read_markdown(path)
    else:
        paragraphs = _read_text(path)
    title = next((text for text, level in paragraphs if level == 1), path.stem.replace("-", " ").replace("_", " ").title())
    sections: list[dict] = []
    current: dict | None = None
    for text, level in paragraphs:
        if level and (level <= 2 or current is None):
            current = {"heading": text, "items": []}
            sections.append(current)
        elif current is not None and text != title:
            current["items"].append(text)
    if not sections:
        sections = [{"heading": title, "items": [text for text, _ in paragraphs]}]
    return {"path": str(path), "title": title, "paragraph_count": len(paragraphs), "sections": sections, "excerpt": "\n".join(text for text, _ in paragraphs)[:4000]}
