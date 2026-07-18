import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import office2llm


def write_docx(path: Path, *, with_image: bool) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", "<document />")
        if with_image:
            archive.writestr("word/media/image1.png", b"image")


class DocxRoutingTests(unittest.TestCase):
    def test_embedded_media_detection(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            plain = Path(directory) / "plain.docx"
            illustrated = Path(directory) / "illustrated.docx"
            write_docx(plain, with_image=False)
            write_docx(illustrated, with_image=True)

            self.assertFalse(office2llm.docx_has_embedded_images(plain))
            self.assertTrue(office2llm.docx_has_embedded_images(illustrated))

    def test_image_free_docx_uses_pandoc_without_api_key(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "plain.docx"
            output = source.with_suffix(".md")
            write_docx(source, with_image=False)

            with patch.dict(os.environ, {}, clear=True):
                with patch.object(
                    office2llm, "docx_to_markdown", return_value=output
                ) as convert:
                    exit_code = office2llm.process_document(
                        source,
                        outdir=None,
                        dpi=200,
                        timeout_s=120,
                        fulltext_only=False,
                    )

            self.assertEqual(exit_code, 0)
            convert.assert_called_once_with(source, output_path=output)

    def test_image_free_doc_uses_libreoffice_then_pandoc_without_api_key(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "legacy.doc"
            source.write_bytes(b"legacy")
            conversion_dir = root / "converted"
            conversion_dir.mkdir()
            converted = conversion_dir / "legacy.docx"
            write_docx(converted, with_image=False)
            output = source.with_suffix(".md")

            with patch.dict(os.environ, {}, clear=True):
                with patch.object(
                    office2llm, "office_to_docx", return_value=converted
                ) as bridge:
                    with patch.object(
                        office2llm, "docx_to_markdown", return_value=output
                    ) as convert:
                        exit_code = office2llm.process_document(
                            source,
                            outdir=None,
                            dpi=200,
                            timeout_s=120,
                            fulltext_only=False,
                        )

            self.assertEqual(exit_code, 0)
            bridge.assert_called_once_with(source, timeout_s=120)
            convert.assert_called_once_with(converted, output_path=output)
            self.assertFalse(conversion_dir.exists())

    def test_docx_with_image_requires_ocr_key_before_rendering(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "illustrated.docx"
            write_docx(source, with_image=True)

            with patch.dict(os.environ, {}, clear=True):
                with patch.object(office2llm, "office_to_pdf") as render:
                    with self.assertRaisesRegex(RuntimeError, "full-page OCR"):
                        office2llm.process_document(
                            source,
                            outdir=None,
                            dpi=200,
                            timeout_s=120,
                            fulltext_only=False,
                        )

            render.assert_not_called()

    def test_doc_with_image_requires_ocr_key_after_inspection(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "illustrated.doc"
            source.write_bytes(b"legacy")
            conversion_dir = root / "converted"
            conversion_dir.mkdir()
            converted = conversion_dir / "illustrated.docx"
            write_docx(converted, with_image=True)

            with patch.dict(os.environ, {}, clear=True):
                with patch.object(
                    office2llm, "office_to_docx", return_value=converted
                ):
                    with patch.object(office2llm, "office_to_pdf") as render:
                        with self.assertRaisesRegex(RuntimeError, "full-page OCR"):
                            office2llm.process_document(
                                source,
                                outdir=None,
                                dpi=200,
                                timeout_s=120,
                                fulltext_only=False,
                            )

            render.assert_not_called()
            self.assertFalse(conversion_dir.exists())


if __name__ == "__main__":
    unittest.main()
