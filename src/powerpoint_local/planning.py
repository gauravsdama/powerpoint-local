from __future__ import annotations

import re
from typing import Any

from .documents import read_source
from .presentation import inspect_template


def _bullets(items: list[str]) -> list[str]:
    result: list[str] = []
    for item in items:
        for part in re.split(r"(?<=[.!?])\s+", item):
            cleaned = re.sub(r"\s+", " ", part).strip(" -•\t")
            if cleaned and cleaned not in result:
                result.append(cleaned[:180])
            if len(result) == 4:
                return result
    return result


def _template_role(role: str, template: dict | None) -> tuple[str, str | None]:
    if not template:
        return role, None
    candidates = template.get("role_map", {}).get(role) or template.get("role_map", {}).get("content") or []
    return role, candidates[0] if candidates else None


def create_plan(source_path: str, template_path: str | None = None, max_slides: int = 10) -> dict[str, Any]:
    max_slides = max(3, min(int(max_slides), 20))
    source = read_source(source_path)
    template = inspect_template(template_path) if template_path else None
    slides: list[dict[str, Any]] = []
    role, layout = _template_role("title", template)
    slides.append({"number": 1, "role": role, "template_layout": layout, "title": source["title"], "bullets": _bullets(source["sections"][0]["items"])[:2], "visual": "Use an Oracle title treatment if a local Oracle template is selected."})
    for section in source["sections"]:
        if len(slides) >= max_slides - 1:
            break
        heading = section["heading"]
        body = _bullets(section["items"])
        lowered = f"{heading} {' '.join(body)}".lower()
        role_name = "comparison" if any(token in lowered for token in ("versus", " vs ", "compare", "before", "after")) else "diagram" if any(token in lowered for token in ("architecture", "workflow", "process", "timeline", "sequence")) else "two_column" if len(body) >= 4 else "content"
        role, layout = _template_role(role_name, template)
        slides.append({"number": len(slides) + 1, "role": role, "template_layout": layout, "title": heading, "bullets": body or ["Add concise supporting evidence from the source document."], "visual": "Mermaid candidate" if role_name == "diagram" else "Use one supporting visual, icon group, or short evidence block."})
    role, layout = _template_role("content", template)
    slides.append({"number": len(slides) + 1, "role": role, "template_layout": layout, "title": "Recommendations and next steps", "bullets": ["Summarize the decision or action requested.", "Name the immediate owner and the next checkpoint."], "visual": "Use a decisive close with two action callouts."})
    return {"source": source["path"], "title": source["title"], "template": template["path"] if template else None, "template_preference": "Oracle" if template and template.get("oracle_candidate") else "Oracle when a local Oracle template is available", "slides": slides, "generation": "deterministic_local_plan"}


def suggest_mermaid(source_path: str, plan: dict[str, Any] | None = None) -> dict[str, Any]:
    plan = plan or create_plan(source_path)
    suggestions: list[dict[str, Any]] = []
    for slide in plan["slides"]:
        if slide["role"] not in {"diagram", "two_column"} and not any(word in slide["title"].lower() for word in ("process", "architecture", "timeline", "workflow")):
            continue
        labels = [re.sub(r"[\[\]{}<>|]", "", item)[:48] for item in slide["bullets"][:4]] or [slide["title"]]
        nodes = [f"N{index}[{label}]" for index, label in enumerate(labels, 1)]
        links = " --> ".join(f"N{index}" for index in range(1, len(nodes) + 1))
        suggestions.append({"slide_number": slide["number"], "title": slide["title"], "mermaid": "flowchart LR\n    " + "\n    ".join(nodes) + "\n    " + links, "note": "Mermaid source only; render it locally if an image is needed."})
    return {"suggestions": suggestions, "count": len(suggestions)}
