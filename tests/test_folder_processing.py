import base64
import contextlib
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import httpx
from PIL import Image

import office2llm


def write_pdf(path, colors=("white", "black")):
    pages = [Image.new("RGB", (72, 72), color) for color in colors]
    pages[0].save(path, save_all=True, append_images=pages[1:], resolution=72)


def ocr_reply(client, request, **kwargs):
    body = json.loads(request.content)
    data = body["contents"][0]["parts"][0]["inlineData"]["data"]
    with Image.open(io.BytesIO(base64.urlsafe_b64decode(data))) as image:
        text = f"Page {image.getpixel((0, 0))[0]}"
    return httpx.Response(200, request=request, json={"candidates": [{
        "content": {"role": "model", "parts": [{"text": text}]},
        "finishReason": "STOP",
    }]})


@contextlib.contextmanager
def cli_environment(*, api_key="test-key", response=ocr_reply):
    output = io.StringIO()
    with (
        patch.dict(os.environ, {
            "GEMINI_API_KEY": api_key, "PATH": os.environ.get("PATH", os.defpath),
        }, clear=True),
        patch.object(httpx.Client, "send", autospec=True, side_effect=response) as send,
        patch("builtins.input", return_value="yes") as prompt,
        contextlib.redirect_stdout(output),
    ):
        yield output, send, prompt


class FolderProcessingTests(unittest.TestCase):
    def test_discovery_is_shallow_by_default_and_recursive_on_request(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "top.pdf").touch()
            (root / "nested").mkdir()
            (root / "nested" / "child.pdf").touch()
            (root / "nested" / "loop").symlink_to(root, target_is_directory=True)
            for flags, expected in (([], 1), (["--recursive"], 2)):
                with self.subTest(flags=flags), cli_environment(api_key="") as (out, send, prompt):
                    self.assertEqual(office2llm.main([
                        "--input", str(root), "--dry-run", *flags,
                    ]), 0)
                    self.assertRegex(out.getvalue(), rf"\| Pending\s+\| {expected}\s+\|")
                    self.assertEqual("child.pdf" in out.getvalue(), bool(flags))
                    send.assert_not_called()
                    prompt.assert_not_called()
            self.assertEqual(len(list(root.rglob("*.txt"))), 0)

    def test_dry_run_excludes_new_and_legacy_artifact_directories_even_when_forced(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "report.pdf").touch()
            for name in ("report", "report.pdf__pages__", "orphan.png__pages__"):
                artifact_dir = root / name
                artifact_dir.mkdir()
                (artifact_dir / "page_0001.png").touch()
            with cli_environment(api_key="") as (out, send, prompt):
                self.assertEqual(office2llm.main([
                    "--input", str(root), "--recursive", "--dry-run", "--force-overwrite",
                ]), 0)
                self.assertRegex(out.getvalue(), r"\| Pending\s+\| 1\s+\|")
                self.assertNotIn("page_0001.png", out.getvalue())
                send.assert_not_called()
                prompt.assert_not_called()

    def test_all_processed_markers_skip_single_files_and_folder_batches(self):
        markers = ("report.txt", "report.pdf.txt", "report.md", "report.pdf.md",
                   "report/", "report.pdf__pages__/")
        for marker_name in markers:
            for folder in (False, True):
                with self.subTest(marker=marker_name, folder=folder), tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    source = root / "report.pdf"
                    source.touch()
                    marker = root / marker_name
                    if marker_name.endswith("/"):
                        marker.mkdir()
                    else:
                        marker.write_text("existing output")
                    with cli_environment(api_key="") as (out, send, prompt):
                        self.assertEqual(office2llm.main([
                            "--input", str(root if folder else source), "--skip-processed",
                        ]), 0)
                        if folder:
                            self.assertNotIn(str(source.resolve()), out.getvalue())
                            self.assertRegex(out.getvalue(), r"\| Skipped\s+\| 1\s+\|")
                        else:
                            self.assertIn(f"existing={marker.resolve()}", out.getvalue())
                            self.assertIn("skipped=1", out.getvalue())
                        send.assert_not_called()
                        prompt.assert_not_called()
                    if marker.is_file():
                        self.assertEqual(marker.read_text(), "existing output")

    def test_native_word_markers_skip_before_opening_or_converting(self):
        for extension in ("doc", "docx"):
            for marker_extension in ("txt", "md"):
                for include_extension in (False, True):
                    with self.subTest(extension=extension, marker=marker_extension, full=include_extension):
                        with tempfile.TemporaryDirectory() as directory:
                            source = Path(directory) / f"report.{extension}"
                            source.touch()
                            marker = source.parent / f"{source.name if include_extension else source.stem}.{marker_extension}"
                            marker.write_text("existing")
                            with cli_environment(api_key="") as (_, send, _):
                                self.assertEqual(office2llm.main(["--input", str(source)]), 0)
                                send.assert_not_called()
                            self.assertEqual(marker.read_text(), "existing")

    def test_extension_filter_matches_case_and_optional_dot_in_dry_run_and_execution(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "nested").mkdir()
            pdf = root / "nested" / "paper.PDF"
            write_pdf(pdf, ("white",))
            Image.new("RGB", (2, 2), "black").save(root / "photo.png")
            (root / "report.docx").touch()
            with cli_environment(api_key="") as (out, send, _):
                self.assertEqual(office2llm.main([
                    "--input", str(root), "--recursive", "--extensions", ".PDF", "png", "--dry-run",
                ]), 0)
                self.assertRegex(out.getvalue(), r"\| Pending\s+\| 2\s+\|")
                self.assertNotIn("report.docx", out.getvalue())
                send.assert_not_called()
            with cli_environment() as (_, send, _):
                self.assertEqual(office2llm.main([
                    "--input", str(root), "--recursive", "--extensions", "PDF", "--fulltext-only",
                ]), 0)
                self.assertEqual(send.call_count, 1)
            self.assertTrue(pdf.with_name("paper.PDF.txt").is_file())
            self.assertFalse((root / "photo.png.txt").exists())
            self.assertFalse((root / "report.md").exists())

    def test_output_modes_work_for_single_files_and_folders_and_always_combine_pages(self):
        for folder in (False, True):
            for flags, keep in (([], True), (["--keep-artifacts"], True), (["--fulltext-only"], False)):
                with self.subTest(folder=folder, flags=flags), tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    source = root / "report.pdf"
                    write_pdf(source)
                    with cli_environment() as (out, send, _):
                        self.assertEqual(office2llm.main([
                            "--input", str(root if folder else source), "--dpi", "144", *flags,
                        ]), 0)
                        self.assertEqual(send.call_count, 2)
                        text = out.getvalue()
                        self.assertIn(f"Proccessing:\n{source.resolve()}\n", text)
                        self.assertNotIn("Processed:", text)
                        result = text.split("Results:\n", 1)[1]
                        self.assertTrue(result.startswith("+"))
                        self.assertNotIn(str(source.resolve()), result)
                        self.assertRegex(result, r"\| pages\s+\| 2\s+\|")
                        self.assertRegex(result, r"\| ocr_ok\s+\| 2\s+\|")
                        self.assertRegex(result, r"\| ocr_failed\s+\| 0\s+\|")
                        self.assertNotIn("report.pdf.txt", result)
                        self.assertNotIn("__pages__", result)
                    self.assertEqual((root / "report.pdf.txt").read_text(), "Page 255\n\nPage 0")
                    artifacts = root / "report.pdf__pages__"
                    self.assertEqual(artifacts.is_dir(), keep)
                    if keep:
                        self.assertEqual(sorted(p.name for p in artifacts.iterdir()), [
                            "page_0001.png", "page_0001.txt", "page_0002.png", "page_0002.txt",
                        ])
                        for number, text in ((1, "Page 255"), (2, "Page 0")):
                            self.assertEqual((artifacts / f"page_{number:04d}.txt").read_text(), text)
                            with Image.open(artifacts / f"page_{number:04d}.png") as image:
                                self.assertEqual(image.mode, "RGB")
                                self.assertEqual(image.size, (144, 144))

    def test_same_stem_inputs_use_separate_artifact_folders_and_text(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            write_pdf(root / "report.pdf", ("white",))
            Image.new("RGB", (2, 2), "black").save(root / "report.png")
            with cli_environment() as (_, send, _):
                self.assertEqual(office2llm.main(["--input", str(root)]), 0)
                self.assertEqual(send.call_count, 2)
            for extension, text in (("pdf", "Page 255"), ("png", "Page 0")):
                self.assertEqual((root / f"report.{extension}.txt").read_text(), text)
                self.assertEqual((root / f"report.{extension}__pages__" / "page_0001.txt").read_text(), text)
            with cli_environment(api_key="") as (_, send, prompt):
                self.assertEqual(office2llm.main(["--input", str(root), "--recursive"]), 0)
                send.assert_not_called()
                prompt.assert_not_called()

    def test_force_regenerates_every_page_and_removes_only_stale_page_artifacts(self):
        for folder in (False, True):
            with self.subTest(folder=folder), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                source = root / "report.pdf"
                write_pdf(source, ("white",))
                final = root / "report.pdf.txt"
                final.write_text("old full text")
                (root / "report.md").write_text("legacy output")
                artifacts = root / "report.pdf__pages__"
                artifacts.mkdir()
                for name in ("page_0001.png", "page_0001.txt", "page_0002.png", "page_0002.txt"):
                    (artifacts / name).write_text("old page")
                (artifacts / "notes.txt").write_text("keep me")
                with cli_environment() as (_, send, _):
                    self.assertEqual(office2llm.main([
                        "--input", str(root if folder else source), "--force-overwrite",
                    ]), 0)
                    self.assertEqual(send.call_count, 1)
                self.assertEqual(final.read_text(), "Page 255")
                self.assertEqual((artifacts / "page_0001.txt").read_text(), "Page 255")
                self.assertFalse((artifacts / "page_0002.txt").exists())
                self.assertFalse((artifacts / "page_0002.png").exists())
                self.assertEqual((artifacts / "notes.txt").read_text(), "keep me")
                self.assertEqual((root / "report.md").read_text(), "legacy output")

    def test_dry_run_reports_skips_and_force_without_mutating_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "report.pdf"
            source.touch()
            final = root / "report.pdf.txt"
            final.write_text("existing")
            (root / "fresh.pdf").touch()
            for flags, pending, skipped in (([], 1, 1), (["--force-overwrite"], 2, 0)):
                with self.subTest(flags=flags), cli_environment(api_key="") as (out, send, prompt):
                    self.assertEqual(office2llm.main([
                        "--input", str(root), "--dry-run", *flags,
                    ]), 0)
                    self.assertRegex(out.getvalue(), rf"\| Pending\s+\| {pending}\s+\|")
                    self.assertRegex(out.getvalue(), rf"\| Skipped\s+\| {skipped}\s+\|")
                    lines = out.getvalue().splitlines()
                    self.assertEqual(lines[0], "Found Paths:")
                    self.assertEqual(lines[1], lines[4])
                    self.assertEqual(lines[2:4], [str((root / "fresh.pdf").resolve()), str(source.resolve())])
                    self.assertNotIn("pending input=", out.getvalue())
                    self.assertNotIn("\033[", out.getvalue())
                    send.assert_not_called()
                    prompt.assert_not_called()
            self.assertEqual(final.read_text(), "existing")
            self.assertEqual(len(list(root.iterdir())), 3)

    def test_dry_run_colors_paths_in_terminals_and_respects_no_color(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "fresh.pdf").touch()
            (root / "done.pdf").touch()
            (root / "done.md").touch()
            for no_color in (False, True):
                with self.subTest(no_color=no_color), cli_environment() as (out, send, prompt):
                    with patch.object(out, "isatty", return_value=True), patch.dict(
                        os.environ, {"NO_COLOR": ""} if no_color else {},
                    ):
                        self.assertEqual(office2llm.main(["--input", str(root), "--dry-run"]), 0)
                    text = out.getvalue()
                    if no_color:
                        self.assertNotIn("\033[", text)
                    else:
                        self.assertIn("\033[1mFound Paths:\033[0m", text)
                        self.assertIn(f"\033[32m{(root / 'fresh.pdf').resolve()}\033[0m", text)
                        self.assertIn(f"\033[2;90m{(root / 'done.pdf').resolve()}\033[0m", text)
                        self.assertRegex(text, r"\x1b\[2;90m\| Recursive\s+\| No\s+\|\x1b\[0m")
                        self.assertRegex(text, r"\x1b\[2;90m\| Force overwrite\s+\| No\s+\|\x1b\[0m")
                    plain = re.sub(r"\x1b\[[0-9;]*m", "", text)
                    self.assertRegex(plain, r"\| OCR output\s+\| <filename.ext>.txt beside input")
                    send.assert_not_called()
                    prompt.assert_not_called()

    def test_normal_folder_preview_lists_only_pending_paths_before_confirmation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "fresh.pdf"
            source.touch()
            (root / "done.pdf").touch()
            (root / "done.md").write_text("existing")
            with cli_environment() as (out, send, prompt):
                def decline(question):
                    text = out.getvalue()
                    self.assertIn("\033[1mFound Paths:\033[0m", text)
                    self.assertIn(f"\033[32m{source.resolve()}\033[0m", text)
                    self.assertNotIn("done.pdf", text)
                    self.assertNotIn("done.md", text)
                    self.assertNotIn("skipped input=", text)
                    plain = re.sub(r"\x1b\[[0-9;]*m", "", text)
                    self.assertRegex(plain, r"\| Pending\s+\| 1\s+\|")
                    self.assertRegex(plain, r"\| Skipped\s+\| 1\s+\|")
                    self.assertIn("Process 1 documents", question)
                    send.assert_not_called()
                    return "no"

                prompt.side_effect = decline
                with patch.object(out, "isatty", return_value=True):
                    with self.assertRaisesRegex(SystemExit, "cancelled"):
                        office2llm.main(["--input", str(root)])
                prompt.assert_called_once()
                send.assert_not_called()
            self.assertFalse((root / "fresh.pdf.txt").exists())

    def test_forced_fulltext_run_overwrites_final_text_without_creating_artifacts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "report.pdf"
            write_pdf(source, ("white",))
            final = root / "report.pdf.txt"
            final.write_text("old text")
            with cli_environment() as (_, send, _):
                self.assertEqual(office2llm.main([
                    "--input", str(root), "--force-overwrite", "--fulltext-only",
                ]), 0)
                self.assertEqual(send.call_count, 1)
            self.assertEqual(final.read_text(), "Page 255")
            self.assertEqual(sorted(p.name for p in root.iterdir()), ["report.pdf", "report.pdf.txt"])

    @unittest.skipUnless(shutil.which("pandoc"), "requires Pandoc")
    def test_native_word_markdown_is_skipped_and_can_be_forcibly_regenerated(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "report.docx"
            subprocess.run([
                "pandoc", "--from=markdown", "--to=docx", "--output", str(source),
            ], input="# Fresh heading\n\nNative text.\n", text=True, check=True, capture_output=True)
            final = root / "report.md"
            final.write_text("existing text")
            with cli_environment(api_key="") as (_, send, prompt):
                self.assertEqual(office2llm.main(["--input", str(root)]), 0)
                prompt.assert_not_called()
                self.assertEqual(final.read_text(), "existing text")
                self.assertEqual(office2llm.main([
                    "--input", str(root), "--force-overwrite", "--keep-artifacts",
                ]), 0)
                send.assert_not_called()
            self.assertIn("Fresh heading", final.read_text())
            self.assertFalse((root / "report.docx__pages__").exists())

    def test_custom_artifact_directory_keeps_combined_text_beside_input_and_is_recognized(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "report.pdf"
            write_pdf(source, ("white",))
            artifacts = root / "custom"
            artifacts.mkdir()
            args = ["--input", str(source), "--outdir", str(artifacts)]
            with cli_environment() as (_, send, _):
                self.assertEqual(office2llm.main(args), 0)
                self.assertEqual(send.call_count, 1)
            final = root / "report.pdf.txt"
            self.assertEqual(final.read_text(), "Page 255")
            self.assertEqual((artifacts / "page_0001.txt").read_text(), "Page 255")
            final.unlink()
            with cli_environment(api_key="") as (out, send, _):
                self.assertEqual(office2llm.main(args), 0)
                self.assertIn(f"existing={artifacts.resolve()}", out.getvalue())
                send.assert_not_called()

    def test_custom_artifact_directory_cannot_overwrite_source_image(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "page_0001.png"
            Image.new("RGB", (2, 2), "white").save(source)
            original = source.read_bytes()
            with cli_environment() as (_, send, _):
                with self.assertRaisesRegex(SystemExit, "overwrite the input"):
                    office2llm.main([
                        "--input", str(source), "--outdir", str(root), "--force-overwrite",
                    ])
                send.assert_not_called()
            self.assertEqual(source.read_bytes(), original)

    def test_empty_or_filtered_folder_is_a_successful_noop(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "report.docx").touch()
            with cli_environment(api_key="") as (out, send, prompt):
                self.assertEqual(office2llm.main(["--input", str(root), "--extensions", "pdf"]), 0)
                self.assertRegex(out.getvalue(), r"\| Pending\s+\| 0\s+\|")
                send.assert_not_called()
                prompt.assert_not_called()

    def test_invalid_flag_combinations_fail_before_processing(self):
        cases = (
            ["--extensions", "exe"],
            ["--keep-artifacts", "--fulltext-only"],
            ["--skip-processed", "--force-overwrite"],
            ["--fulltext-only", "--outdir", "/unused"],
            ["--outdir", "/unused"],
        )
        with tempfile.TemporaryDirectory() as directory:
            for flags in cases:
                with self.subTest(flags=flags), contextlib.redirect_stderr(io.StringIO()):
                    with self.assertRaises(SystemExit) as error:
                        office2llm.main(["--input", directory, *flags])
                    self.assertEqual(error.exception.code, 2)

    def test_batch_continues_after_a_corrupt_document(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "broken.pdf").write_bytes(b"not a PDF")
            Image.new("RGB", (2, 2), "white").save(root / "valid.png")
            with cli_environment() as (out, send, _):
                self.assertEqual(office2llm.main(["--input", str(root), "--fulltext-only"]), 2)
                self.assertEqual(send.call_count, 1)
                self.assertRegex(out.getvalue(), r"\| failed\s+\| 1\s+\|")
            self.assertEqual((root / "valid.png.txt").read_text(), "Page 255")
            self.assertFalse((root / "broken.pdf.txt").exists())

    def test_failed_ocr_does_not_publish_partial_combined_text_or_replace_old_text(self):
        def reject(client, request, **kwargs):
            return httpx.Response(400, request=request, json={"error": {"message": "rejected", "code": 400}})

        for existing in (False, True):
            for fulltext in (False, True):
                with self.subTest(existing=existing, fulltext=fulltext), tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    source = root / "report.png"
                    Image.new("RGB", (2, 2), "white").save(source)
                    final = root / "report.png.txt"
                    if existing:
                        final.write_text("previous successful output")
                    flags = ["--fulltext-only"] if fulltext else []
                    with cli_environment(response=reject), patch.object(office2llm.time, "sleep"):
                        self.assertEqual(office2llm.main([
                            "--input", str(source), "--force-overwrite", *flags,
                        ]), 2)
                    self.assertEqual(final.exists(), existing)
                    if existing:
                        self.assertEqual(final.read_text(), "previous successful output")
                    self.assertEqual((root / "report.png__pages__").is_dir(), not fulltext)
