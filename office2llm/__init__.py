import argparse
import io
import os
import shutil
import subprocess
import tempfile
import time
import zipfile
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

_ELIGIBLE_EXTENSIONS = {
    ".pdf",
    ".doc",
    ".docx",
    ".gif",
    ".jpeg",
    ".jpg",
    ".ppt",
    ".pptx",
    ".png",
    ".xls",
    ".xlsx",
    ".odt",
    ".ods",
    ".odp",
    ".rtf",
    ".tif",
    ".tiff",
    ".webp",
}


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


def office_to_format(
    input_path: Path, *, output_format: str, timeout_s: int = 120
) -> Path:
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
                output_format,
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

        expected = tmpdir / f"{input_path.stem}.{output_format}"
        if expected.exists():
            return expected
        outputs = sorted(tmpdir.glob(f"*.{output_format}"))
        if len(outputs) == 1:
            return outputs[0]
        raise RuntimeError(
            f"LibreOffice conversion succeeded but no {output_format.upper()} found in {tmpdir}"
        )
    except subprocess.CalledProcessError as error:
        detail = error.stderr.decode(errors="replace").strip()
        shutil.rmtree(tmpdir, ignore_errors=True)
        message = detail or f"LibreOffice exited with status {error.returncode}"
        raise RuntimeError(
            f"LibreOffice {output_format.upper()} conversion failed: {message}"
        ) from error
    except Exception:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise


def office_to_pdf(input_path: Path, *, timeout_s: int = 120) -> Path:
    return office_to_format(input_path, output_format="pdf", timeout_s=timeout_s)


def office_to_docx(input_path: Path, *, timeout_s: int = 120) -> Path:
    return office_to_format(input_path, output_format="docx", timeout_s=timeout_s)


def docx_has_embedded_images(input_path: Path) -> bool:
    try:
        with zipfile.ZipFile(input_path) as archive:
            return any(
                name.startswith("word/media/") and not name.endswith("/")
                for name in archive.namelist()
            )
    except zipfile.BadZipFile as error:
        raise RuntimeError(f"invalid DOCX file: {input_path}") from error


def markdown_output_path(input_path: Path, outdir: Path | None) -> Path:
    if outdir is None:
        return input_path.with_suffix(".md")
    return outdir.expanduser().resolve() / f"{input_path.stem}.md"


def docx_to_markdown(input_path: Path, *, output_path: Path) -> Path:
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise RuntimeError("Pandoc not found on $PATH")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    try:
        subprocess.run(
            [
                pandoc,
                str(input_path),
                "--from=docx",
                "--to=gfm",
                "--wrap=none",
                "--markdown-headings=atx",
                f"--output={tmp_path}",
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        tmp_path.replace(output_path)
    except subprocess.CalledProcessError as error:
        tmp_path.unlink(missing_ok=True)
        detail = error.stderr.decode(errors="replace").strip()
        raise RuntimeError(f"Pandoc conversion failed: {detail or error}") from error
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
    return output_path


def word_to_markdown_if_native(
    input_path: Path, *, outdir: Path | None, timeout_s: int
) -> Path | None:
    temporary_docx = None
    if input_path.suffix.lower() == ".doc":
        temporary_docx = office_to_docx(input_path, timeout_s=timeout_s)
        docx_path = temporary_docx
    else:
        docx_path = input_path

    try:
        if docx_has_embedded_images(docx_path):
            return None
        return docx_to_markdown(
            docx_path, output_path=markdown_output_path(input_path, outdir)
        )
    finally:
        if temporary_docx is not None:
            shutil.rmtree(temporary_docx.parent, ignore_errors=True)


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


def image_to_png_page(image_path: Path, *, outdir: Path) -> int:
    outdir.mkdir(parents=True, exist_ok=True)
    with Image.open(image_path) as pil:
        if pil.mode not in ("RGB", "RGBA"):
            pil = pil.convert("RGB")
        elif pil.mode == "RGBA":
            bg = Image.new("RGB", pil.size, (255, 255, 255))
            bg.paste(pil, mask=pil.getchannel("A"))
            pil = bg

        buf = io.BytesIO()
        pil.save(buf, format="PNG", optimize=False)
        (outdir / "page_0001.png").write_bytes(buf.getvalue())
    return 1


def process_document(
    input_path: Path,
    *,
    outdir: Path | None,
    dpi: int,
    timeout_s: int,
    fulltext_only: bool,
) -> int:
    if fulltext_only and outdir is not None:
        raise SystemExit("--fulltext-only cannot be used with --outdir")

    if input_path.suffix.lower() in {".doc", ".docx"}:
        output_path = word_to_markdown_if_native(
            input_path, outdir=outdir, timeout_s=timeout_s
        )
        if output_path is not None:
            mode = (
                "libreoffice+pandoc"
                if input_path.suffix.lower() == ".doc"
                else "pandoc"
            )
            print(f"ok input={input_path} mode={mode} output={output_path}")
            return 0

    if not os.environ.get("GEMINI_API_KEY"):
        raise RuntimeError(
            "GEMINI_API_KEY is required because this input needs full-page OCR."
        )

    if outdir is not None:
        resolved_outdir = outdir.expanduser().resolve()
    elif fulltext_only:
        resolved_outdir = Path(tempfile.mkdtemp(prefix="office2llm_pages_"))
    else:
        resolved_outdir = (input_path.parent / input_path.stem).resolve()

    tmp_pdf: Path | None = None
    final_txt_path = input_path.parent / f"{input_path.name}.txt"
    try:
        if input_path.suffix.lower() == ".pdf":
            pages = pdf_to_png_pages(input_path, outdir=resolved_outdir, dpi=dpi)
        elif input_path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".gif", ".tif", ".tiff"}:
            pages = image_to_png_page(input_path, outdir=resolved_outdir)
        else:
            tmp_pdf = office_to_pdf(input_path, timeout_s=timeout_s)
            pages = pdf_to_png_pages(tmp_pdf, outdir=resolved_outdir, dpi=dpi)

        ocr_ok = 0
        ocr_skipped = 0
        ocr_failed = 0
        page_texts = [""] * pages
        if pages > 0:
            max_workers = min(4, pages)
            with ThreadPoolExecutor(max_workers=max_workers) as ex:
                futures = {}
                for i in range(1, pages + 1):
                    png_path = resolved_outdir / f"page_{i:04d}.png"
                    txt_path = resolved_outdir / f"page_{i:04d}.txt"
                    if not fulltext_only and txt_path.exists():
                        ocr_skipped += 1
                        continue
                    futures[ex.submit(run_ocr, png_path)] = (i - 1, txt_path)

                for fut in as_completed(futures):
                    page_idx, txt_path = futures[fut]
                    try:
                        text = fut.result()
                        page_texts[page_idx] = text or ""
                        if not fulltext_only:
                            tmp_path = txt_path.with_suffix(txt_path.suffix + ".tmp")
                            tmp_path.write_text(text or "", encoding="utf-8")
                            tmp_path.replace(txt_path)
                        ocr_ok += 1
                    except Exception as e:
                        ocr_failed += 1
                        print(f"ocr failed file={txt_path.name} err={e}")

        if fulltext_only and ocr_failed == 0:
            tmp_path = final_txt_path.with_suffix(final_txt_path.suffix + ".tmp")
            tmp_path.write_text("\n\n".join(page_texts), encoding="utf-8")
            tmp_path.replace(final_txt_path)

        print(
            f"ok input={input_path} pages={pages} ocr_ok={ocr_ok} ocr_skipped={ocr_skipped} "
            f"ocr_failed={ocr_failed} "
            f"{'output=' + str(final_txt_path) if fulltext_only else 'outdir=' + str(resolved_outdir)}"
        )
        return 0 if ocr_failed == 0 else 2
    finally:
        if fulltext_only:
            shutil.rmtree(resolved_outdir, ignore_errors=True)
        if tmp_pdf is not None:
            shutil.rmtree(tmp_pdf.parent, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Convert native Word files to Markdown or use full-page OCR when needed."
    )
    ap.add_argument(
        "--input", required=True, help="Path to input file (doc/docx/pptx/xlsx/pdf/...)."
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
    ap.add_argument(
        "--fulltext-only",
        action="store_true",
        help=(
            "Write one sibling output and remove OCR intermediates. Native Word uses "
            ".md; OCR-routed inputs use <input-filename>.<ext>.txt."
        ),
    )
    args = ap.parse_args(argv)

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f"input not found: {input_path}")

    if input_path.is_dir():
        if args.outdir:
            raise SystemExit("--outdir cannot be used when --input points to a directory")
        inputs = sorted(
            path
            for path in input_path.iterdir()
            if path.is_file() and path.suffix.lower() in _ELIGIBLE_EXTENSIONS
        )
        if not inputs:
            raise SystemExit(f"no eligible documents found in: {input_path}")
        try:
            answer = input(
                f"Process {len(inputs)} documents in {input_path} and write sibling Markdown or OCR text files? [y/N] "
            )
        except EOFError:
            raise SystemExit("confirmation required for directory input")
        if answer.strip().lower() not in {"y", "yes"}:
            raise SystemExit("cancelled")

        failures = 0
        for doc_path in inputs:
            try:
                exit_code = process_document(
                    doc_path,
                    outdir=None,
                    dpi=args.dpi,
                    timeout_s=args.timeout_s,
                    fulltext_only=True,
                )
            except Exception as e:
                failures += 1
                print(f"failed input={doc_path} err={e}")
                continue
            if exit_code != 0:
                failures += 1
        print(f"batch processed={len(inputs)} failed={failures}")
        return 0 if failures == 0 else 2

    try:
        return process_document(
            input_path,
            outdir=Path(args.outdir) if args.outdir else None,
            dpi=args.dpi,
            timeout_s=args.timeout_s,
            fulltext_only=args.fulltext_only,
        )
    except RuntimeError as error:
        raise SystemExit(str(error)) from error


def cli() -> None:
    raise SystemExit(main())
