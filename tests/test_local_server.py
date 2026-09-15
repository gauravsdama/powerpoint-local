from __future__ import annotations

import ast
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from powerpoint_local.documents import read_source
from powerpoint_local.mermaid_bridge import bridge_config
from powerpoint_local.planning import create_plan, suggest_mermaid
from powerpoint_local.presentation import export_pngs, inspect_template, powerpoint_version, prepare_guided_deck


class LocalServerTests(unittest.TestCase):
    @staticmethod
    def _presentation_fixture(path: Path, slide_count: int = 1) -> None:
        with zipfile.ZipFile(path, "w") as archive:
            for number in range(1, slide_count + 1):
                archive.writestr(
                    f"ppt/slides/slide{number}.xml",
                    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>Slide</a:t></p:sld>',
                )

    def test_prepare_without_template_creates_blank_powerpoint(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "deck.pptx"
            def create_fixture(_lines: list[str], arguments: list[str] | None = None) -> str:
                self.assertIsNotNone(arguments)
                Path(arguments[0]).write_bytes(b"pptx")
                return ""

            with patch("powerpoint_local.presentation._run_osascript", side_effect=create_fixture):
                result = prepare_guided_deck({"slides": []}, str(output), open_after=False)
            self.assertEqual(result["creation_mode"], "local_blank_powerpoint")
            self.assertEqual(result["deck_path"], str(output.resolve()))
            self.assertTrue(output.exists())
            self.assertTrue(Path(result["manifest_path"]).is_file())

    def test_powerpoint_version_enforces_tested_floor(self) -> None:
        with patch("powerpoint_local.presentation._run_osascript", return_value="16.95.1"):
            self.assertTrue(powerpoint_version()["supported"])
        with patch("powerpoint_local.presentation._run_osascript", return_value="16.94"):
            self.assertFalse(powerpoint_version()["supported"])

    def test_pdf_to_png_fallback_closes_temporary_presentation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            presentation = Path(directory) / "fixture.pptx"
            output = Path(directory) / "pngs"
            self._presentation_fixture(presentation)
            scripts: list[tuple[list[str], list[str] | None]] = []

            def applescript(lines: list[str], arguments: list[str] | None = None) -> str:
                scripts.append((lines, arguments))
                if any("save as PDF" in line for line in lines):
                    Path(arguments[0]).write_bytes(b"pdf")
                return ""

            def command(arguments: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
                if arguments[0] == "/usr/bin/sips":
                    Path(arguments[arguments.index("--out") + 1]).write_bytes(b"png")
                return subprocess.CompletedProcess(arguments, 0, "", "")

            with patch("powerpoint_local.presentation.open_in_powerpoint"), patch(
                "powerpoint_local.presentation._run_osascript", side_effect=applescript
            ), patch("powerpoint_local.presentation.subprocess.run", side_effect=command), patch(
                "powerpoint_local.presentation.time.sleep"
            ):
                result = export_pngs(str(presentation), str(output))

            self.assertEqual(result["export_mode"], "powerpoint_pdf_single_slide_fallback")
            self.assertEqual(result["count"], 1)
            self.assertTrue(any(any("close presentation" in line for line in lines) for lines, _ in scripts))

    def test_pdf_fallback_failure_still_closes_temporary_presentation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            presentation = Path(directory) / "fixture.pptx"
            self._presentation_fixture(presentation)
            closed: list[str] = []

            def applescript(lines: list[str], arguments: list[str] | None = None) -> str:
                if any("save as PDF" in line for line in lines):
                    raise RuntimeError("simulated PDF export failure")
                if any("close presentation" in line for line in lines):
                    closed.append(arguments[0])
                return ""

            with patch("powerpoint_local.presentation.open_in_powerpoint"), patch(
                "powerpoint_local.presentation._run_osascript", side_effect=applescript
            ), patch("powerpoint_local.presentation.subprocess.run", return_value=subprocess.CompletedProcess([], 0, "", "")), patch(
                "powerpoint_local.presentation.time.sleep"
            ):
                with self.assertRaisesRegex(RuntimeError, "simulated PDF export failure"):
                    export_pngs(str(presentation), str(Path(directory) / "pngs"))

            self.assertEqual(closed, ["slide-1.pptx"])

    def test_mermaid_bridge_configuration_is_portable_and_loopback_only(self) -> None:
        with patch.dict(
            "os.environ",
            {
                "POWERPOINT_LOCAL_MERMAID_NODE": "/tmp/node",
                "POWERPOINT_LOCAL_MERMAID_ENTRY": "/tmp/mermaid/index.js",
                "POWERPOINT_LOCAL_MERMAID_GATEWAY": "http://localhost:8787",
            },
        ):
            node, entry, gateway = bridge_config()
        self.assertEqual(node, Path("/tmp/node"))
        self.assertEqual(entry, Path("/tmp/mermaid/index.js"))
        self.assertEqual(gateway, "http://localhost:8787")

        with patch.dict("os.environ", {"POWERPOINT_LOCAL_MERMAID_GATEWAY": "https://example.com"}):
            with self.assertRaisesRegex(ValueError, "localhost"):
                bridge_config()
        with patch.dict("os.environ", {"POWERPOINT_LOCAL_MERMAID_GATEWAY": "http://localhost:8787@example.com"}):
            with self.assertRaisesRegex(ValueError, "localhost"):
                bridge_config()

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
