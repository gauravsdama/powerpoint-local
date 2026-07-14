from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .local_files import SOURCE_EXTENSIONS, ensure_local_file

MAX_SOURCE_CHARS = 250_000
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _clean(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _read_docx(path: Path) -> list[tuple[str, int | None]]:
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    root = ET.fromstring(xml)
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
    for raw in path.read_text(encoding="utf-8", errors="ignore")[:MAX_SOURCE_CHARS].splitlines():
        match = re.match(r"^(#{1,6})\s+(.+)$", raw.strip())
        if match:
            result.append((_clean(match.group(2)), len(match.group(1))))
        elif raw.strip():
            result.append((_clean(re.sub(r"^[-*+]\s+", "", raw)), None))
    return result


def _read_text(path: Path) -> list[tuple[str, int | None]]:
    return [(_clean(value), None) for value in path.read_text(encoding="utf-8", errors="ignore")[:MAX_SOURCE_CHARS].splitlines() if _clean(value)]


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
