from __future__ import annotations

import ast
import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from powerpoint_local.documents import MAX_DOCX_XML_BYTES, MAX_SOURCE_BYTES, read_source
from powerpoint_local.mermaid_bridge import _read_response, bridge_config, create_diagram
from powerpoint_local.planning import create_plan, suggest_mermaid
from powerpoint_local.presentation import MAX_INSPECTION_XML_BYTES, capture_screen, export_pngs, inspect_template, open_in_powerpoint, powerpoint_version, prepare_guided_deck


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
            scripts: list[list[str]] = []

            def create_fixture(_lines: list[str], arguments: list[str] | None = None) -> str:
                scripts.append(_lines)
                self.assertIsNotNone(arguments)
                self.assertEqual(Path(arguments[0]).parent, output.parent.resolve())
                self.assertIn(".powerpoint-local-staging-", Path(arguments[0]).name)
                Path(arguments[0]).write_bytes(b"pptx")
                return ""

            with patch("powerpoint_local.presentation._run_osascript", side_effect=create_fixture):
                result = prepare_guided_deck({"slides": []}, str(output), open_after=False)
            self.assertEqual(result["creation_mode"], "local_blank_powerpoint")
            self.assertEqual(result["deck_path"], str(output.resolve()))
            self.assertTrue(output.exists())
            self.assertTrue(Path(result["manifest_path"]).is_file())
            self.assertEqual(result["owned_artifacts"], [str(output.resolve()), result["manifest_path"]])
            self.assertTrue(any("close active presentation saving no" in line for line in scripts[0]))
            self.assertEqual(list(output.parent.glob("*.powerpoint-local-staging-*")), [])

    def test_prepare_rejects_overwrite_and_non_pptx_template_packages(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "deck.pptx"
            output.write_bytes(b"existing")
            with self.assertRaisesRegex(ValueError, "overwrite"):
                prepare_guided_deck({"slides": []}, str(output), open_after=False)
            output.unlink()
            for extension in (".potx", ".pptm", ".potm"):
                template = root / f"template{extension}"
                template.write_bytes(b"package")
                with self.assertRaisesRegex(ValueError, "only \\.pptx"):
                    prepare_guided_deck({"slides": []}, str(output), str(template), open_after=False)

    def test_powerpoint_version_enforces_tested_floor(self) -> None:
        with patch("powerpoint_local.presentation._run_osascript", return_value="16.95.1"):
            self.assertTrue(powerpoint_version()["supported"])
        with patch("powerpoint_local.presentation._run_osascript", return_value="16.94"):
            self.assertFalse(powerpoint_version()["supported"])

    def test_powerpoint_launch_timeout_has_recovery_guidance(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            presentation = Path(directory) / "fixture.pptx"
            self._presentation_fixture(presentation)
            with patch(
                "powerpoint_local.presentation.subprocess.run",
                side_effect=subprocess.TimeoutExpired(["/usr/bin/open"], 30),
            ):
                with self.assertRaisesRegex(RuntimeError, "Resolve any open Office dialog"):
                    open_in_powerpoint(str(presentation))

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
            self.assertEqual(result["images"], [str(output.resolve() / "Slide-1.png")])
            self.assertTrue(any(any("close presentation" in line for line in lines) for lines, _ in scripts))
            path_arguments = [
                Path(argument)
                for _lines, arguments in scripts
                for argument in (arguments or [])
                if Path(argument).is_absolute() and argument.endswith((".png", ".pptx", ".pdf"))
            ]
            self.assertTrue(path_arguments)
            self.assertTrue(all(argument.parent == output.parent.resolve() for argument in path_arguments))

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

            self.assertEqual(len(closed), 1)
            self.assertTrue(closed[0].endswith("-slide-1.pptx"))
            self.assertFalse((Path(directory) / "pngs").exists())

    def test_export_rejects_existing_directory_before_opening_powerpoint(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            presentation = root / "fixture.pptx"
            output = root / "pngs"
            self._presentation_fixture(presentation)
            output.mkdir()
            (output / "Slide-99.png").write_bytes(b"stale")
            with patch("powerpoint_local.presentation.open_in_powerpoint") as open_mock:
                with self.assertRaisesRegex(ValueError, "existing output_dir"):
                    export_pngs(str(presentation), str(output))
            open_mock.assert_not_called()
            self.assertTrue((output / "Slide-99.png").is_file())

    def test_capture_rejects_existing_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "capture.png"
            output.write_bytes(b"existing")
            with patch("powerpoint_local.presentation._run_osascript") as automation_mock:
                with self.assertRaisesRegex(ValueError, "overwrite"):
                    capture_screen(str(output))
            automation_mock.assert_not_called()

    def test_capture_publishes_only_completed_staged_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "capture.png"

            def capture(arguments: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
                staged_output = Path(arguments[-1])
                self.assertNotEqual(staged_output, output)
                staged_output.write_bytes(b"png")
                return subprocess.CompletedProcess(arguments, 0, "", "")

            with patch("powerpoint_local.presentation._run_osascript"), patch(
                "powerpoint_local.presentation.subprocess.run", side_effect=capture
            ):
                result = capture_screen(str(output))
            self.assertEqual(output.read_bytes(), b"png")
            self.assertEqual(result["owned_artifacts"], [str(output.resolve())])

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
            self.assertEqual(sum(slide["title"] == "Local Deck" for slide in plan["slides"]), 1)
            self.assertGreaterEqual(len(plan["slides"]), 3)
            self.assertEqual(suggest_mermaid(str(source), plan)["count"], 1)

    def test_source_and_docx_reads_enforce_uncompressed_limits(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            oversized_text = root / "large.txt"
            oversized_text.write_bytes(b"x" * (MAX_SOURCE_BYTES + 1))
            with self.assertRaisesRegex(ValueError, "byte limit"):
                read_source(str(oversized_text))

            oversized_docx = root / "large.docx"
            with zipfile.ZipFile(oversized_docx, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("word/document.xml", b"x" * (MAX_DOCX_XML_BYTES + 1))
            with self.assertRaisesRegex(ValueError, "uncompressed limit"):
                read_source(str(oversized_docx))

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

    def test_template_inventory_rejects_oversized_uncompressed_xml(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            presentation = Path(directory) / "template.potx"
            with zipfile.ZipFile(presentation, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("ppt/slideLayouts/slideLayout1.xml", b"x" * (MAX_INSPECTION_XML_BYTES + 1))
            with self.assertRaisesRegex(ValueError, "inspection limit"):
                inspect_template(str(presentation))

    def test_mermaid_read_deadline_is_not_blocked_by_partial_line(self) -> None:
        process = subprocess.Popen(
            [sys.executable, "-c", "import sys,time; sys.stdout.write('{'); sys.stdout.flush(); time.sleep(5)"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        started = time.monotonic()
        try:
            with self.assertRaisesRegex(RuntimeError, "deadline"):
                _read_response(process, 1, 0.1)
            self.assertLess(time.monotonic() - started, 1.0)
        finally:
            process.kill()
            process.wait(timeout=2)
            if process.stdout:
                process.stdout.close()
            if process.stderr:
                process.stderr.close()

    def test_mermaid_timeout_reaps_child_and_closes_pipes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            script = Path(directory) / "slow_bridge.py"
            pid_path = script.with_suffix(".pid")
            script.write_text(
                "import os,sys,time\n"
                "from pathlib import Path\n"
                "Path(__file__).with_suffix('.pid').write_text(str(os.getpid()))\n"
                "sys.stdout.write('{')\n"
                "sys.stdout.flush()\n"
                "time.sleep(5)\n",
                encoding="utf-8",
            )
            with patch(
                "powerpoint_local.mermaid_bridge.bridge_config",
                return_value=(Path(sys.executable), script, "http://127.0.0.1:8787"),
            ), patch("powerpoint_local.mermaid_bridge.MERMAID_RESPONSE_TIMEOUT_SECONDS", 0.1), patch(
                "powerpoint_local.mermaid_bridge.MERMAID_SHUTDOWN_TIMEOUT_SECONDS", 0.5
            ):
                with self.assertRaisesRegex(RuntimeError, "deadline"):
                    create_diagram("Title", "flowchart LR\nA --> B")
            child_pid = int(pid_path.read_text(encoding="utf-8"))
            with self.assertRaises(ProcessLookupError):
                os.kill(child_pid, 0)

    def test_json_rpc_errors_batches_notifications_and_request_limit(self) -> None:
        requests = [
            "{not-json",
            '{"jsonrpc":"2.0","id":NaN,"method":"ping"}',
            "[]",
            json.dumps({"jsonrpc": "2.0", "id": True, "method": "ping"}),
            json.dumps({"jsonrpc": "2.0", "id": 6, "method": "ping", "params": []}),
            json.dumps([
                {"jsonrpc": "2.0", "id": 7, "method": "ping"},
                {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}},
            ]),
            json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}),
            "x" * (1024 * 1024 + 1),
            json.dumps({"jsonrpc": "2.0", "id": 8, "method": "ping"}),
        ]
        process = subprocess.run(
            [sys.executable, str(ROOT / "src" / "powerpoint_local" / "server.py")],
            input="\n".join(requests) + "\n",
            capture_output=True,
            text=True,
            timeout=10,
            check=True,
        )
        replies = [json.loads(line) for line in process.stdout.splitlines()]
        self.assertEqual([reply["error"]["code"] for reply in replies[:5]], [-32700, -32700, -32600, -32600, -32602])
        self.assertEqual(replies[5], [{"jsonrpc": "2.0", "id": 7, "result": {}}])
        self.assertEqual(replies[6]["error"]["code"], -32600)
        self.assertEqual(replies[7], {"jsonrpc": "2.0", "id": 8, "result": {}})

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
