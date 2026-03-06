import argparse
import io
import os
import shutil
import subprocess
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image


_EXTRACTION_PROMPT = """
Return the plain text representation of the provided image as if you were reading it naturally. Extract all visible text while preserving the document's semantic structure (headers, hierarchy, data relationships, lists, tables).

Guidelines:

- Reading Order: Follow a logical human reading order (e.g., if there are two distinct columns, process the left column completely before the right, unless the content clearly spans across).
- Multi-pages: This is likely one page out of several in the document, so be sure to preserve any sentences that come from the previous page, or continue onto the next page, exactly as they are.
- Empty pages: If there is no text at all that you think you should read, do not return anything
- Malformed text: If text is blurry or cut off, transcribe exactly what you see; do not hallucinate or auto-complete missing words.
- Placeholders: Do not use placeholders like `[Signature]` or `[Image]` unless strictly necessary for context.
- Visuals: Do not describe visual elements (e.g., do not say "There is a logo," "Image of a graph"). Ignore watermarks or noise.
- Missing text: Ensure no text is missed, including headers, footers, footnotes, references or text in margins, as long as it contains readable information.

Output Format:

- Markdown only for structure: Do not use Markdown for headers (#, ##, ###) or bold text, that is not necessary for the RAG use case.
- Tables: Represent tables using standard Markdown syntax (`| Header | ... |`). ensure row and column alignment is preserved. If a cell contains multi-line text, flatten it into a single line within the cell.
- Lists: Use proper Markdown list syntax (`-` for unordered, `1.` for ordered) rather than just newlines.
- Key-Value Pairs: Extract explicit key-value pairs only when both text elements are visible (e.g., "Invoice #: 12345"). Do not generate artificial keys or labels (such as adding "Category:", "Date:", or "Label:") if that text is not explicitly written in the image. If a value (like a tag or status) appears without a label, transcribe it simply as text on its own line or as a sub-header, preserving the visual hierarchy without adding words.
- Equations/Math: If present, represent mathematical formulas using LaTeX syntax inside `$ ... $`.
- Handwriting: Read any natural handwriting and include it.
- Output ONLY the raw extracted text: Do not include preambles (e.g., "Here is the markdown..."), code block fences (```), or concluding remarks.
""".strip()


def run_ocr(image: bytes | Path) -> str:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is required for OCR. Export it and re-run.")

    try:
        import importlib

        genai = importlib.import_module("google.genai")
        types = importlib.import_module("google.genai.types")
    except Exception as e:
        raise RuntimeError("Missing dependency: google-genai. Reinstall office2llm.") from e

    if isinstance(image, Path):
        image_bytes = image.read_bytes()
    else:
        image_bytes = image
    mime_type = "image/png"

    client = genai.Client(api_key=api_key)
    try:
        model = "gemini-3-flash-preview"
        delay_s = 1.0
        for attempt in range(5):
            try:
                response: types.GenerateContentResponse = client.models.generate_content(
                    model=model,
                    contents=[
                        types.Content(
                            role="user",
                            parts=[
                                types.Part.from_bytes(
                                    mime_type=mime_type, data=image_bytes
                                ),
                            ],
                        ),
                    ],
                    config=types.GenerateContentConfig(
                        system_instruction=_EXTRACTION_PROMPT,
                        thinking_config=types.ThinkingConfig(
                            thinking_level="HIGH",
                        ),
                    ),
                )
                return response.text or ""
            except Exception:
                if attempt >= 4:
                    raise
                time.sleep(delay_s)
                delay_s = min(delay_s * 2.0, 10.0)
        return ""
    finally:
        client.close()


def office_to_pdf(input_path: Path, *, timeout_s: int = 120) -> Path:
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice/soffice not found on $PATH")

    tmpdir = Path(tempfile.mkdtemp(prefix="office2llm_"))
    try:
        # Sane defaults for sandboxes/containers:
        # - writable HOME (LibreOffice still writes a user profile even in headless mode)
        # - UTF-8 locale (prevents exit 77: "UI language cannot be determined")
        env = os.environ.copy()
        env["HOME"] = str(tmpdir)
        if env.get("LANG", "C") in ("", "C", "POSIX"):
            env["LANG"] = "C.UTF-8"
        if env.get("LC_ALL", "") in ("", "C", "POSIX"):
            env["LC_ALL"] = env["LANG"]

        subprocess.run(
            [
                soffice,
                "--headless",
                "--nologo",
                "--norestore",
                "--nolockcheck",
                "--nofirststartwizard",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmpdir),
                str(input_path),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_s,
            env=env,
        )

        expected = tmpdir / f"{input_path.stem}.pdf"
        if expected.exists():
            return expected
        pdfs = sorted(tmpdir.glob("*.pdf"))
        if len(pdfs) == 1:
            return pdfs[0]
        raise RuntimeError(f"LibreOffice conversion succeeded but no PDF found in {tmpdir}")
    except Exception:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise


def pdf_to_png_pages(pdf_path: Path, *, outdir: Path, dpi: int) -> int:
    outdir.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        n_pages = len(doc)
        if n_pages <= 0:
            return 0
        scale = max(0.1, float(dpi) / 72.0)

        for i in range(n_pages):
            page = doc[i]
            pil: Image.Image = page.render(scale=scale).to_pil()

            # Deterministic, OCR-friendly PNGs (no alpha).
            if pil.mode not in ("RGB", "RGBA"):
                pil = pil.convert("RGB")
            elif pil.mode == "RGBA":
                bg = Image.new("RGB", pil.size, (255, 255, 255))
                bg.paste(pil, mask=pil.getchannel("A"))
                pil = bg

            buf = io.BytesIO()
            pil.save(buf, format="PNG", optimize=False)
            (outdir / f"page_{i+1:04d}.png").write_bytes(buf.getvalue())

        return n_pages
    finally:
        doc.close()


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Convert Office/PDF to per-page PNGs and LLM OCR text."
    )
    ap.add_argument(
        "--input", required=True, help="Path to input file (docx/pptx/xlsx/pdf/...)."
    )
    ap.add_argument(
        "--outdir",
        required=False,
        default=None,
        help=(
            "Output directory for page_XXXX.png and page_XXXX.txt files. "
            "Default: create a sibling folder next to the input named after the input file (e.g. ./foo.docx -> ./foo/)."
        ),
    )
    ap.add_argument("--dpi", type=int, default=200, help="Render DPI (default: 200).")
    ap.add_argument(
        "--timeout-s",
        type=int,
        default=120,
        help="LibreOffice convert timeout seconds.",
    )
    args = ap.parse_args(argv)

    if not os.environ.get("GEMINI_API_KEY"):
        raise SystemExit("GEMINI_API_KEY is required for OCR output.")

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f"input not found: {input_path}")

    if args.outdir:
        outdir = Path(args.outdir).expanduser().resolve()
    else:
        outdir = (input_path.parent / input_path.stem).resolve()

    tmp_pdf: Path | None = None
    try:
        if input_path.suffix.lower() == ".pdf":
            pdf_path = input_path
        else:
            tmp_pdf = office_to_pdf(input_path, timeout_s=args.timeout_s)
            pdf_path = tmp_pdf

        pages = pdf_to_png_pages(pdf_path, outdir=outdir, dpi=args.dpi)

        ocr_ok = 0
        ocr_skipped = 0
        ocr_failed = 0
        if pages > 0:
            max_workers = min(4, pages)
            with ThreadPoolExecutor(max_workers=max_workers) as ex:
                futures = {}
                for i in range(1, pages + 1):
                    png_path = outdir / f"page_{i:04d}.png"
                    txt_path = outdir / f"page_{i:04d}.txt"
                    if txt_path.exists():
                        ocr_skipped += 1
                        continue
                    futures[ex.submit(run_ocr, png_path)] = txt_path

                for fut in as_completed(futures):
                    txt_path = futures[fut]
                    try:
                        text = fut.result()
                        tmp_path = txt_path.with_suffix(txt_path.suffix + ".tmp")
                        tmp_path.write_text(text or "", encoding="utf-8")
                        tmp_path.replace(txt_path)
                        ocr_ok += 1
                    except Exception as e:
                        ocr_failed += 1
                        print(f"ocr failed file={txt_path.name} err={e}")

        print(
            f"ok pages={pages} ocr_ok={ocr_ok} ocr_skipped={ocr_skipped} ocr_failed={ocr_failed} outdir={outdir}"
        )
        return 0 if ocr_failed == 0 else 2
    finally:
        if tmp_pdf is not None:
            shutil.rmtree(tmp_pdf.parent, ignore_errors=True)


def cli() -> None:
    raise SystemExit(main())
