# office2llm

Convert native Word documents into Markdown and use full-page OCR when visual context is required.

- **DOCX → Markdown**: image-free DOCX files use Pandoc and do not require an API key
- **DOC → DOCX → Markdown**: image-free legacy Word files use LibreOffice followed by Pandoc
- **Image-bearing DOCX → OCR**: any embedded image routes the entire document through the full-page OCR pipeline
- **Office → PDF**: uses LibreOffice (`libreoffice` / `soffice`) in headless mode
- **PDF → PNG**: renders pages via `pypdfium2` and writes `page_0001.png`, `page_0002.png`, …
- **Image → PNG**: normalizes supported image inputs into a single OCR-ready page image
- **PNG → OCR text**: uses LLM based API to extract text from each page, preserving the document's semantic structure (headers, hierarchy, data relationships, lists, tables).

## Requirements

- **Python**: 3.10+
- **macOS**: 13+ for the bundled PDFium binary
- **Pandoc**: required for image-free Word conversion
- **Gemini API key**: required only when the input routes to OCR
  - Export `GEMINI_API_KEY` before running.
- **LibreOffice**: required for legacy `.doc` and inputs routed to OCR (`.docx` with images, `.pptx`, `.xlsx`, …)
  - The binary must be discoverable as `libreoffice` or `soffice` on `PATH`.
  - `office2llm` sets a **writable temporary `HOME`** and **UTF-8 locale defaults** for the subprocess
    to avoid common headless failures in sandboxes/containers (e.g. exit code 77 / “UI language cannot be determined”).

## Install (recommended: one command)

```bash
bash ./install.sh
```

This will:
- install Pandoc, LibreOffice, and common fonts (macOS + Linux, best-effort)
- create a venv at `~/.office2llm/.venv`
- install this package into the venv
- symlink `office2llm` into `~/.local/bin/office2llm`

If `office2llm` is not found afterwards:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

## Dependency maintenance

Runtime dependencies are pinned with hashes in `requirements.txt`, used by both
the installer and Docker. After editing `pyproject.toml`, regenerate the lock with
uv 0.8.22:

```bash
uv pip compile pyproject.toml --universal --python-version 3.10 --generate-hashes -o requirements.txt
```

Use `--upgrade-package PACKAGE` for a targeted transitive update. Verify with
`uvx pip-audit==2.10.1 -r requirements.txt --require-hashes` and
`python -m unittest discover -s tests -v` in the installed environment.
CI runs the audit and tests; Dependabot checks Python and GitHub Actions weekly.
The locked cryptography release no longer provides Intel macOS support; use the
Linux Docker image on Intel Macs. Apple Silicon macOS and Linux are supported.

## Quick start

```bash
office2llm --input /path/to/report.docx
```

An image-free `.docx` or `.doc` produces `/path/to/report.md`. Legacy `.doc` first passes through a temporary LibreOffice DOCX conversion. If the converted Word document contains an embedded image, `office2llm` uses full-page OCR and requires `GEMINI_API_KEY`.

## Usage

Show help:

```bash
office2llm --help
```

### Convert an image-free Word document to Markdown

```bash
office2llm --input /path/to/report.docx
```

The output is written beside the input:

- `/path/to/report.docx` → `/path/to/report.md`

Pandoc preserves native headings, lists, emphasis, and tables. An embedded image causes the complete document to use the existing OCR pipeline so the image remains in its rendered page context.

Legacy `.doc` input uses the same output and routing rules:

```bash
office2llm --input /path/to/legacy-report.doc
```

### Convert with a custom output folder

```bash
office2llm --input /path/to/report.docx --outdir /tmp/report
```

For an image-free DOCX, this writes `/tmp/report/report.md`. Other inputs use the OCR output structure.

### Convert a PDF (LibreOffice not used)

```bash
office2llm --input /path/to/file.pdf --outdir ./out --dpi 250
```

### Convert an image

```bash
office2llm --input /path/to/image.png --fulltext-only
```

### Convert to a single sibling fulltext file only

```bash
office2llm --input /path/to/example.pdf --fulltext-only
```

This writes:

- `/path/to/example.pdf.txt`

No output folder is kept in this mode.

### Scan or convert a folder

```bash
office2llm --input /path/to/folder --recursive --extensions pdf docx --dry-run
yes | office2llm --input /path/to/folder --recursive --extensions pdf docx --fulltext-only
```

Folder scans are non-recursive by default. Add `--recursive` to include subfolders.
`--extensions` accepts one or more supported extensions, case-insensitively and with
or without a dot (`pdf`, `.PDF`, `docx`). Without it, all supported formats are selected.
Dry runs list pending and skipped files without conversion, API calls, writes, or confirmation.
The **Found Paths:** list shows one path per line between horizontal rules. In a
terminal, pending paths are green and skipped paths are dim gray; a key/value table
shows counts, filters, and output settings. Redirected output is plain text, and
`NO_COLOR` disables terminal styling.
Actual folder processing asks for confirmation; an empty or fully processed selection exits successfully.

Output modes work identically for files and folders:

- `--keep-artifacts` (default) keeps page images and per-page OCR in
  `<filename.ext>__pages__/`, and writes combined text beside the source.
- `--fulltext-only` writes the combined result and removes this run's temporary
  artifacts. It does not remove artifacts retained by earlier runs.

Sibling outputs include:

- `/path/to/folder/example.pdf.txt`
- `/path/to/folder/report.md` for image-free DOCX or DOC
- `/path/to/folder/illustrated.docx.txt` for Word files routed to OCR
- `/path/to/folder/photo.jpg.txt`

### Skip processed documents or regenerate them

`--skip-processed` is the default. For `report.pdf`, any of these existing outputs
causes a skip before conversion or credential checks:

- `report.txt`, `report.pdf.txt`, `report.md`, or `report.pdf.md`
- `report.pdf__pages__/` or the legacy `report/` directory

An empty output file or matching directory also counts as processed. To retry an
incomplete run or regenerate output, use `--force-overwrite`. This bypasses every
processed marker and regenerates every page; it removes stale numbered page images
and text in the selected artifact directory while preserving unrelated files.
The combined text is replaced only after every page succeeds.

```bash
office2llm --input /path/to/folder --recursive --extensions pdf --force-overwrite --keep-artifacts
```

Recursive scans always exclude directories ending in `__pages__`, plus legacy
directories whose names match the stem of a supported sibling input. They do not
follow directory symlinks. The forgiving legacy rules can also match unrelated
same-stem outputs; `--force-overwrite` bypasses document skips, but artifact
directories remain excluded from discovery.

For single-file `--outdir`, existing page artifacts and native Markdown outputs
are also recognized. `--outdir` is unavailable for folder input and cannot be
combined with `--fulltext-only`.

### Tune timeouts (Office → PDF step)

```bash
office2llm --input ./big.xlsx --timeout-s 300
```

## Output

Image-free Word files produce one Markdown file, without OCR artifacts in either
output mode. Successful OCR runs always write `<filename.ext>.txt` beside the input.
With artifacts retained, `report.pdf` additionally produces:

```text
report.pdf.txt
report.pdf__pages__/
  page_0001.png
  page_0001.txt
  page_0002.png
  page_0002.txt
```

`report.png` uses its own `report.png__pages__/` directory. A custom `--outdir`
changes only the OCR artifact location; combined OCR text stays beside the input.

PNG output is deterministic and “OCR-friendly” (no alpha channel).

## Troubleshooting

### LibreOffice fails with exit code 77 (UI language cannot be determined)

This is typically caused by missing/empty locale environment variables (e.g. `LANG`, `LC_ALL`) in minimal
environments. `office2llm` runs LibreOffice with a writable temporary `HOME` and sets UTF-8 locale defaults
for the LibreOffice subprocess. If your base image/OS provides *no* UTF-8 locales at all, you may need to
install/generate one system-wide.

## Docker (optional)

Build + run locally:

```bash
docker build -t office2llm .
docker run --rm -v "$PWD:/work" office2llm --input /work/in.docx --outdir /work/out --dpi 200
```

Or via the included compose file (uses `./data` mounted to `/data`):

```bash
docker compose -f docker-compose.yml run --rm office2llm --input /data/in.docx --outdir /data/out --dpi 200
```
