from __future__ import annotations

import ast
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from powerpoint_local.documents import read_source
from powerpoint_local.planning import create_plan, suggest_mermaid
from powerpoint_local.presentation import inspect_template


class LocalServerTests(unittest.TestCase):
    def test_markdown_to_deterministic_plan(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.md"
            source.write_text("# Local Deck\n\n## Workflow\n\nCollect input. Review it. Share results.\n", encoding="utf-8")
            parsed = read_source(str(source))
            plan = create_plan(str(source), max_slides=6)
            self.assertEqual(parsed["title"], "Local Deck")
            self.assertEqual(plan["slides"][0]["role"], "title")
            self.assertGreaterEqual(len(plan["slides"]), 3)
            self.assertEqual(suggest_mermaid(str(source), plan)["count"], 1)

    def test_template_inventory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            presentation = Path(directory) / "Oracle-template.potx"
            layout = '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Title Slide"><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sldLayout>'
            with zipfile.ZipFile(presentation, "w") as archive:
                archive.writestr("ppt/slideLayouts/slideLayout1.xml", layout)
                archive.writestr("ppt/slides/slide1.xml", '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>Welcome</a:t></p:sld>')
            inventory = inspect_template(str(presentation))
            self.assertTrue(inventory["oracle_candidate"])
            self.assertEqual(inventory["layouts"][0]["role"], "title")

    def test_server_has_no_network_imports_and_stdio_smoke(self) -> None:
        forbidden = {"urllib", "http", "requests", "socket", "webbrowser", "ftplib"}
        for source in (ROOT / "src" / "powerpoint_local").glob("*.py"):
            tree = ast.parse(source.read_text(encoding="utf-8"))
            imports = {
                alias.name.split(".")[0] if isinstance(node, ast.Import) else (node.module or "").split(".")[0]
                for node in ast.walk(tree)
                if isinstance(node, (ast.Import, ast.ImportFrom))
                for alias in (node.names if isinstance(node, ast.Import) else [node])
            }
            self.assertFalse(imports & forbidden, f"network import in {source}")
        process = subprocess.run([sys.executable, str(ROOT / "src" / "powerpoint_local" / "server.py")], input=json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-03-26"}}) + "\n" + json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}}) + "\n", capture_output=True, text=True, timeout=10, check=True)
        replies = [json.loads(line) for line in process.stdout.splitlines()]
        self.assertEqual(replies[0]["result"]["serverInfo"]["name"], "powerpoint-local")
        self.assertGreaterEqual(len(replies[1]["result"]["tools"]), 10)


if __name__ == "__main__":
    unittest.main()
