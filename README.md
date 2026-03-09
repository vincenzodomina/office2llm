# office2llm

Convert Office documents and PDFs into **OCR-friendly, per-page PNGs** and **per-page OCR text** for advanced document understanding.

- **Office → PDF**: uses LibreOffice (`libreoffice` / `soffice`) in headless mode
- **PDF → PNG**: renders pages via `pypdfium2` and writes `page_0001.png`, `page_0002.png`, …
- **PNG → OCR text**: uses LLM based API to extract text from each page, preserving the document's semantic structure (headers, hierarchy, data relationships, lists, tables).

## Requirements

- **Python**: 3.10+
- **Gemini API key**: required for OCR output
  - Export `GEMINI_API_KEY` before running.
- **LibreOffice**: required for non-PDF inputs (`.docx`, `.pptx`, `.xlsx`, …)
  - The binary must be discoverable as `libreoffice` or `soffice` on `PATH`.
  - `office2llm` sets a **writable temporary `HOME`** and **UTF-8 locale defaults** for the subprocess
    to avoid common headless failures in sandboxes/containers (e.g. exit code 77 / “UI language cannot be determined”).

## Install (recommended: one command)

```bash
bash ./install.sh
```

This will:
- install LibreOffice + common fonts (macOS + Linux, best-effort)
- create a venv at `~/.office2llm/.venv`
- install this package into the venv
- symlink `office2llm` into `~/.local/bin/office2llm`

If `office2llm` is not found afterwards:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

## Quick start

```bash
export GEMINI_API_KEY="your-api-key-here"
office2llm --input /path/to/report.docx
```

## Usage

Show help:

```bash
office2llm --help
```

### Convert an Office file (default output folder)

```bash
office2llm --input /path/to/report.docx
```

If you omit `--outdir`, outputs go to a sibling folder next to the input named after the input file:

- `/path/to/report.docx` → `/path/to/report/page_0001.png`, …

### Convert an Office file (custom output folder)

```bash
office2llm --input /path/to/report.pptx --outdir /tmp/report-pages --dpi 200
```

### Convert a PDF (LibreOffice not used)

```bash
office2llm --input /path/to/file.pdf --outdir ./out --dpi 250
```

### Convert to a single sibling fulltext file only

```bash
office2llm --input /path/to/example.pdf --fulltext-only
```

This writes:

- `/path/to/example.pdf.txt`

No output folder is kept in this mode.

### Tune timeouts (Office → PDF step)

```bash
office2llm --input ./big.xlsx --timeout-s 300
```

## Output

The command writes:
- `page_0001.png`
- `page_0002.png`
- …
- `page_0001.txt`
- `page_0002.txt`
- …

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

