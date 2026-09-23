import base64
import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import httpx
from PIL import Image

import office2llm


class DependencyIntegrationTests(unittest.TestCase):
    def test_image_normalization_and_multipage_pdf_rendering(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for mode in ("RGB", "RGBA", "L"):
                source = root / f"{mode}.png"
                Image.new(mode, (32, 24)).save(source)
                office2llm.image_to_png_page(source, outdir=root / mode)
                with Image.open(root / mode / "page_0001.png") as result:
                    self.assertEqual(result.mode, "RGB")
                    self.assertEqual(result.size, (32, 24))
                    if mode == "RGBA":
                        self.assertEqual(result.getpixel((0, 0)), (255, 255, 255))

            pdf = root / "pages.pdf"
            page = Image.new("RGB", (72, 72), "white")
            page.save(pdf, save_all=True, append_images=[page], resolution=72)
            self.assertEqual(
                office2llm.pdf_to_png_pages(pdf, outdir=root / "pages", dpi=144), 2
            )
            for number in (1, 2):
                with Image.open(root / "pages" / f"page_{number:04d}.png") as result:
                    self.assertEqual(result.mode, "RGB")
                    self.assertEqual(result.size, (144, 144))

    def test_ocr_sdk_serializes_image_and_reads_response(self):
        buffer = io.BytesIO()
        Image.new("RGB", (2, 2), "white").save(buffer, format="PNG")
        image = buffer.getvalue()
        requests = []

        def send(client, request, **kwargs):
            requests.append(request)
            return httpx.Response(
                200,
                request=request,
                json={"candidates": [{"content": {
                    "role": "model", "parts": [{"text": "Extracted text"}]
                }, "finishReason": "STOP"}]},
            )

        with patch.dict(os.environ, {"GEMINI_API_KEY": "test-key"}, clear=True):
            with patch.object(httpx.Client, "send", send):
                self.assertEqual(office2llm.run_ocr(image), "Extracted text")

        self.assertEqual(len(requests), 1)
        self.assertTrue(requests[0].url.path.endswith(":generateContent"))
        body = json.loads(requests[0].content)
        part = body["contents"][0]["parts"][0]["inlineData"]
        self.assertEqual(base64.urlsafe_b64decode(part["data"]), image)
        self.assertEqual(part.get("mimeType", part.get("mime_type")), "image/png")
        thinking = body["generationConfig"]["thinkingConfig"]
        self.assertEqual(thinking.get("thinkingLevel", thinking.get("thinking_level")), "HIGH")
        self.assertTrue(body["systemInstruction"]["parts"][0]["text"])
