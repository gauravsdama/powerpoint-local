from __future__ import annotations

import json
import os
import plistlib
import re
import sqlite3
from pathlib import Path
from typing import Any, Iterable

POWERPOINT_EXTENSIONS = {".pptx", ".pptm", ".potx", ".potm"}
SOURCE_EXTENSIONS = {".md", ".txt", ".docx"}
MAX_METADATA_BYTES = 8 * 1024 * 1024
MAX_METADATA_FILES = 250


def home_path(*parts: str) -> Path:
    return Path.home().joinpath(*parts)


def template_roots() -> list[dict[str, str]]:
    return [
        {"kind": "Office user templates", "path": str(home_path("Library", "Group Containers", "UBF8T346G9.Office", "User Content.localized", "Templates.localized"))},
        {"kind": "PowerPoint downloaded-template cache", "path": str(home_path("Library", "Containers", "com.microsoft.Powerpoint", "Data", "Library", "Caches"))},
        {"kind": "Office downloaded-template cache", "path": str(home_path("Library", "Group Containers", "UBF8T346G9.Office", "Library", "Caches"))},
    ]


def metadata_roots() -> list[Path]:
    return [
        home_path("Library", "Containers", "com.microsoft.Powerpoint", "Data", "Library", "Preferences"),
        home_path("Library", "Containers", "com.microsoft.Powerpoint", "Data", "Library", "HTTPStorages"),
        home_path("Library", "Group Containers", "UBF8T346G9.Office"),
    ]


def ensure_local_file(value: str, extensions: set[str] | None = None) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ValueError("A non-empty local file path is required.")
    if "://" in value or value.startswith("file://"):
        raise ValueError("URLs are not supported. Save the file locally and pass its filesystem path.")
    path = Path(value).expanduser().resolve()
    if not path.is_file():
        raise ValueError(f"Local file not found: {path}")
    if extensions and path.suffix.lower() not in extensions:
        choices = ", ".join(sorted(extensions))
        raise ValueError(f"Unsupported file type {path.suffix!r}; expected one of: {choices}.")
    return path


def iter_files(root: Path, suffixes: set[str], max_count: int) -> Iterable[Path]:
    if not root.is_dir():
        return
    count = 0
    try:
        for path in root.rglob("*"):
            if count >= max_count:
                return
            if path.is_file() and path.suffix.lower() in suffixes:
                count += 1
                yield path
    except OSError:
        return


def find_templates(query: str = "", limit: int = 30) -> dict[str, Any]:
    limit = max(1, min(int(limit), 100))
    query = query.lower().strip()
    candidates: list[dict[str, Any]] = []
    for root in template_roots():
        path = Path(root["path"])
        for file_path in iter_files(path, POWERPOINT_EXTENSIONS, 300):
            display = str(file_path)
            if query and query not in display.lower():
                continue
            stat = file_path.stat()
            candidates.append({
                "path": display,
                "kind": root["kind"],
                "extension": file_path.suffix.lower(),
                "bytes": stat.st_size,
                "modified_epoch": int(stat.st_mtime),
                "oracle_candidate": "oracle" in display.lower(),
                "availability": "local_file",
            })
    candidates.sort(key=lambda item: (not item["oracle_candidate"], -item["modified_epoch"], item["path"].lower()))
    return {"total": len(candidates), "items": candidates[:limit], "has_more": len(candidates) > limit, "roots": template_roots()}


def _flatten_strings(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, bytes):
        yield value.decode("utf-8", errors="ignore")
    elif isinstance(value, dict):
        for key, nested in value.items():
            yield str(key)
            yield from _flatten_strings(nested)
    elif isinstance(value, (list, tuple)):
        for nested in value:
            yield from _flatten_strings(nested)


def _strings_from_sqlite(path: Path) -> Iterable[str]:
    uri = path.as_uri() + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True)
    try:
        tables = connection.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        for (table,) in tables[:20]:
            quoted = '"' + table.replace('"', '""') + '"'
            columns = connection.execute(f"PRAGMA table_info({quoted})").fetchall()
            text_columns = [row[1] for row in columns if str(row[2]).upper() in {"TEXT", "VARCHAR", "CLOB"}]
            if not text_columns:
                continue
            selected = ", ".join('"' + column.replace('"', '""') + '"' for column in text_columns[:8])
            for row in connection.execute(f"SELECT {selected} FROM {quoted} LIMIT 100"):
                yield from _flatten_strings(row)
    finally:
        connection.close()


REFERENCE_PATTERN = re.compile(r"(?:https?://[^\s\"'<>()]+|[^\s\"'<>()]*(?:sharepoint|onedrive|oracle)[^\s\"'<>()]*)", re.IGNORECASE)


def cache_references(query: str = "", limit: int = 30) -> dict[str, Any]:
    query_lower = query.lower().strip()
    references: list[dict[str, str]] = []
    examined = 0
    seen: set[tuple[str, str]] = set()
    supported = {".plist", ".sqlite", ".db", ".json", ".xml", ".txt"}
    for root in metadata_roots():
        for path in iter_files(root, supported, MAX_METADATA_FILES - examined):
            examined += 1
            try:
                if path.stat().st_size > MAX_METADATA_BYTES:
                    continue
                if path.suffix.lower() == ".plist":
                    strings = _flatten_strings(plistlib.loads(path.read_bytes()))
                elif path.suffix.lower() in {".sqlite", ".db"}:
                    strings = _strings_from_sqlite(path)
                elif path.suffix.lower() == ".json":
                    strings = _flatten_strings(json.loads(path.read_text(encoding="utf-8", errors="ignore")))
                else:
                    strings = [path.read_text(encoding="utf-8", errors="ignore")]
                for text in strings:
                    for match in REFERENCE_PATTERN.findall(text):
                        normalized = match.strip(" .,;:[]")
                        if query_lower and query_lower not in normalized.lower():
                            continue
                        key = (str(path), normalized)
                        if key in seen:
                            continue
                        seen.add(key)
                        lowered = normalized.lower()
                        category = "sharepoint_reference" if "sharepoint" in lowered or "onedrive" in lowered else "oracle_reference" if "oracle" in lowered else "cache_reference"
                        references.append({"source_cache": str(path), "reference": normalized, "category": category, "availability": "reference_only"})
                        if len(references) >= limit:
                            return {"items": references, "examined_files": examined, "note": "References are metadata only. This server never fetches them."}
            except (OSError, ValueError, plistlib.InvalidFileException, sqlite3.Error, json.JSONDecodeError):
                continue
    return {"items": references, "examined_files": examined, "note": "References are metadata only. This server never fetches them."}
