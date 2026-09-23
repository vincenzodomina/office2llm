import argparse
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
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

_PAGES_SUFFIX = "__pages__"


def processed_output(input_path: Path, outdir: Path | None = None) -> Path | None:
    for name in (input_path.stem, input_path.name):
        for extension in (".txt", ".md"):
            output = input_path.parent / f"{name}{extension}"
            if output.is_file():
                return output
    if outdir is not None:
        markdown = markdown_output_path(input_path, outdir)
        if markdown.is_file():
            return markdown
        if outdir.is_dir() and any(
            re.fullmatch(r"page_\d{4,}\.(png|txt)", path.name) and path.is_file()
            for path in outdir.iterdir()
        ):
            return outdir
    for directory in (
        input_path.parent / f"{input_path.name}{_PAGES_SUFFIX}",
        input_path.with_suffix(""),
    ):
        if directory.is_dir():
            return directory
    return None


def discover_documents(
    root: Path, *, recursive: bool, extensions: set[str]
) -> list[Path]:
    def raise_walk_error(error: OSError) -> None:
        raise error

    inputs = []
    if root.name.endswith(_PAGES_SUFFIX):
        return inputs
    for directory, subdirs, filenames in os.walk(root, onerror=raise_walk_error):
        legacy_outputs = {
            Path(name).stem
            for name in filenames
            if Path(name).suffix.lower() in _ELIGIBLE_EXTENSIONS
        }
        subdirs[:] = [
            name for name in subdirs
            if recursive
            and not name.endswith(_PAGES_SUFFIX)
            and name not in legacy_outputs
        ]
        inputs.extend(
            Path(directory) / name
            for name in filenames
            if Path(name).suffix.lower() in extensions
            and (Path(directory) / name).is_file()
        )
    return sorted(inputs)


def styled(text: str, code: str) -> str:
    if sys.stdout.isatty() and "NO_COLOR" not in os.environ:
        return f"\033[{code}m{text}\033[0m"
    return text


def print_path_status(title: str, path: Path, *, failed: bool = False) -> None:
    print(styled(title, "1"))
    print(styled(str(path), "31" if failed else "32"), flush=True)


def print_scan_summary(paths: list[tuple[Path, Path | None]], metadata: dict[str, str]) -> None:
    rule = "─" * min(
        max([len("Found Paths:"), *(len(str(path)) for path, _ in paths)]),
        shutil.get_terminal_size().columns,
    )
    print(styled("Found Paths:", "1"))
    print(rule)
    for path, existing in paths:
        print(styled(str(path), "2;90" if existing is not None else "32"))
    print(rule)
    print()

    print_table(metadata)


def print_table(metadata: dict[str, str]) -> None:
    key_width = max(len(key) for key in metadata)
    value_width = min(
        max(len(value) for value in metadata.values()),
        max(20, shutil.get_terminal_size().columns - key_width - 7),
    )
    border = f"+-{'-' * key_width}-+-{'-' * value_width}-+"
    print(border)
    for key, value in metadata.items():
        for index, line in enumerate(textwrap.wrap(value, width=value_width, break_on_hyphens=False)):
            row = f"| {key if index == 0 else '':<{key_width}} | {line:<{value_width}} |"
            if value.casefold() in {"false", "no", "not retained"} or key == "Skipped":
                row = styled(row, "2;90")
            elif key == "Pending":
                row = styled(row, "32")
            print(row)
    print(border)


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
                        automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True),
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
    force_overwrite: bool = False,
) -> int:
    if fulltext_only and outdir is not None:
        raise SystemExit("--fulltext-only cannot be used with --outdir")

    if not force_overwrite:
        existing = processed_output(input_path, outdir)
        if existing is not None:
            print(f"skipped input={input_path} existing={existing}")
            return 0

    print_path_status("Proccessing:", input_path)
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
            print(styled("Results:", "1"))
            print_table({"mode": mode})
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
        resolved_outdir = input_path.parent / f"{input_path.name}{_PAGES_SUFFIX}"

    tmp_pdf: Path | None = None
    final_txt_path = input_path.parent / f"{input_path.name}.txt"
    try:
        if resolved_outdir.resolve() == input_path.parent.resolve() and re.fullmatch(
            r"page_\d{4,}\.(png|txt)", input_path.name
        ):
            raise RuntimeError("page output would overwrite the input file")
        if not fulltext_only and force_overwrite and resolved_outdir.is_dir():
            for artifact in resolved_outdir.iterdir():
                if artifact.is_file() and re.fullmatch(r"page_\d{4,}\.(png|txt)", artifact.name):
                    artifact.unlink()
        if input_path.suffix.lower() == ".pdf":
            pages = pdf_to_png_pages(input_path, outdir=resolved_outdir, dpi=dpi)
        elif input_path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".gif", ".tif", ".tiff"}:
            pages = image_to_png_page(input_path, outdir=resolved_outdir)
        else:
            tmp_pdf = office_to_pdf(input_path, timeout_s=timeout_s)
            pages = pdf_to_png_pages(tmp_pdf, outdir=resolved_outdir, dpi=dpi)

        ocr_ok = 0
        ocr_failed = 0
        page_texts = [""] * pages
        if pages > 0:
            max_workers = min(4, pages)
            with ThreadPoolExecutor(max_workers=max_workers) as ex:
                futures = {}
                for i in range(1, pages + 1):
                    png_path = resolved_outdir / f"page_{i:04d}.png"
                    txt_path = resolved_outdir / f"page_{i:04d}.txt"
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
                        print_table({"page": str(page_idx + 1), "error": str(e)})

        if ocr_failed == 0:
            tmp_path = final_txt_path.with_suffix(final_txt_path.suffix + ".tmp")
            tmp_path.write_text("\n\n".join(page_texts), encoding="utf-8")
            tmp_path.replace(final_txt_path)

        print(styled("Results:", "1"))
        ok_result = styled(f"ocr_ok | {ocr_ok}", "32" if ocr_ok > 0 else "2;90")
        failed_result = styled(f"ocr_failed | {ocr_failed}", "31" if ocr_failed > 0 else "2;90")
        print(f"| page: {pages} | {ok_result} | {failed_result} |")
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
        "--input", required=True, help="Path to an input file or folder."
    )
    ap.add_argument(
        "--outdir",
        required=False,
        default=None,
        help=(
            "Output directory for page_XXXX.png and page_XXXX.txt files. "
            "Single-file input only. Default: <filename.ext>__pages__/ beside the input. "
            "Combined OCR text is always written beside the input."
        ),
    )
    ap.add_argument("--dpi", type=int, default=200, help="Render DPI (default: 200).")
    ap.add_argument(
        "--timeout-s",
        type=int,
        default=120,
        help="LibreOffice convert timeout seconds.",
    )
    output_mode = ap.add_mutually_exclusive_group()
    output_mode.add_argument(
        "--fulltext-only",
        action="store_true",
        help=(
            "Write one sibling output and remove OCR intermediates. Native Word uses "
            ".md; OCR-routed inputs use <input-filename>.<ext>.txt."
        ),
    )
    output_mode.add_argument(
        "--keep-artifacts", dest="fulltext_only", action="store_false",
        help="Keep page images and page OCR in addition to combined text (default).",
    )
    ap.set_defaults(fulltext_only=False)
    ap.add_argument(
        "--recursive", action="store_true", help="Include subfolders in a folder scan."
    )
    ap.add_argument(
        "--extensions", nargs="+", metavar="EXT",
        help="Only select these supported extensions, e.g. pdf .DOCX (default: all).",
    )
    ap.add_argument(
        "--dry-run", action="store_true",
        help="List pending and skipped inputs without writing files or calling OCR.",
    )
    overwrite = ap.add_mutually_exclusive_group()
    overwrite.add_argument(
        "--skip-processed", dest="force_overwrite", action="store_false",
        help="Skip inputs with matching text, Markdown or artifact outputs (default).",
    )
    overwrite.add_argument(
        "--force-overwrite", action="store_true",
        help="Regenerate selected inputs, overwriting their outputs and page artifacts.",
    )
    ap.set_defaults(force_overwrite=False)
    args = ap.parse_args(argv)

    extensions = _ELIGIBLE_EXTENSIONS
    if args.extensions is not None:
        extensions = {"." + value.lstrip(".").lower() for value in args.extensions}
        unsupported = extensions - _ELIGIBLE_EXTENSIONS
        if unsupported:
            ap.error(f"unsupported extensions: {', '.join(sorted(unsupported))}")
    if args.fulltext_only and args.outdir:
        ap.error("--fulltext-only cannot be used with --outdir")

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f"input not found: {input_path}")

    is_directory = input_path.is_dir()
    if is_directory and args.outdir:
        ap.error("--outdir cannot be used when --input points to a directory")
    outdir = Path(args.outdir).expanduser().resolve() if args.outdir else None
    inputs = (
        discover_documents(input_path, recursive=args.recursive, extensions=extensions)
        if is_directory else [input_path]
    )
    pending = []
    skipped = 0
    found = []
    for doc_path in inputs:
        if doc_path.suffix.lower() not in extensions:
            if not args.dry_run:
                print(f"excluded input={doc_path} reason=extension")
            continue
        existing = None if args.force_overwrite else processed_output(doc_path, outdir)
        if args.dry_run:
            found.append((doc_path, existing))
        if existing is not None:
            skipped += 1
            if not args.dry_run and not is_directory:
                print(f"skipped input={doc_path} existing={existing}")
        else:
            pending.append(doc_path)
    if args.dry_run or is_directory:
        print_scan_summary(found if args.dry_run else [(path, None) for path in pending], {
            "Mode": "Dry run" if args.dry_run else "Process",
            "Input": str(input_path),
            "Recursive": "Yes" if args.recursive else "No",
            "Extensions": ", ".join(sorted(extensions)) if args.extensions else "All supported",
            "Force overwrite": "Yes" if args.force_overwrite else "No",
            "OCR output": "<filename.ext>.txt beside input",
            "Native output": str(outdir / "<stem>.md") if outdir else "<stem>.md beside input",
            "Artifacts": "Not retained" if args.fulltext_only else str(outdir or "<filename.ext>__pages__/ beside input"),
            "Pending": str(len(pending)),
            "Skipped": str(skipped),
        })
    if args.dry_run:
        return 0
    if not pending:
        if not is_directory:
            print(f"batch pending={len(pending)} skipped={skipped}")
        return 0

    if is_directory:
        try:
            answer = input(
                f"Process {len(pending)} documents in {input_path} and write sibling Markdown or OCR text files? [y/N] "
            )
        except EOFError:
            raise SystemExit("confirmation required for directory input")
        if answer.strip().lower() not in {"y", "yes"}:
            raise SystemExit("cancelled")

        failures = 0
        for doc_path in pending:
            try:
                exit_code = process_document(
                    doc_path,
                    outdir=None,
                    dpi=args.dpi,
                    timeout_s=args.timeout_s,
                    fulltext_only=args.fulltext_only,
                    force_overwrite=args.force_overwrite,
                )
            except Exception as e:
                failures += 1
                print_path_status("Failed:", doc_path, failed=True)
                print_table({"error": str(e)})
                continue
            if exit_code != 0:
                failures += 1
        print(styled("Batch results:", "1"))
        print_table({"selected": str(len(pending)), "skipped": str(skipped), "failed": str(failures)})
        return 0 if failures == 0 else 2

    try:
        return process_document(
            input_path,
            outdir=outdir,
            dpi=args.dpi,
            timeout_s=args.timeout_s,
            fulltext_only=args.fulltext_only,
            force_overwrite=args.force_overwrite,
        )
    except RuntimeError as error:
        raise SystemExit(str(error)) from error


def cli() -> None:
    raise SystemExit(main())
