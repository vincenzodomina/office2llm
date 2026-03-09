---
name: office2llm
description: Convert any Office document or PDF into per-page PNGs and LLM-powered OCR text extraction for advanced document understanding.
---

## What it does

```
Office doc (.docx, .pptx, .xlsx, …)
  └─ LibreOffice (headless) ─▶ PDF ─▶ pypdfium2 ─▶ page_0001.png, page_0002.png, …
                                                  └─ LLM OCR ─▶ page_0001.txt, page_0002.txt, …

PDF (.pdf)
  └─ pypdfium2 (direct, no LibreOffice) ─▶ page_0001.png, page_0002.png, …
                                          └─ LLM OCR ─▶ page_0001.txt, page_0002.txt, …
```

All PNGs are deterministic, RGB (no alpha), and optimized for downstream consumption.

Each page is sent to an OCR-capable LLMfor OCR text extraction. The LLM preserves the document's semantic structure — headers, hierarchy, data relationships, lists, tables, key-value pairs, equations, and handwriting — and outputs clean plain text with minimal Markdown for structure (tables, lists).

## Supported input formats

| Category | Extensions |
| --- | --- |
| Word | `.docx`, `.doc`, `.odt`, `.rtf` |
| Slides | `.pptx`, `.ppt`, `.odp` |
| Sheets | `.xlsx`, `.xls`, `.ods` |
| PDF | `.pdf` (rendered directly, LibreOffice not needed) |

Anything LibreOffice can open will work — the list above covers the most common cases.

## Quick start

```bash
# 1. Install (macOS or Linux — installs LibreOffice + creates a venv)
bash ./install.sh

# 2. Export your LLM API key (required for OCR)
export GEMINI_API_KEY="your-api-key-here"

# 3. Convert a document (produces PNGs + OCR text files)
office2llm --input report.docx
```

If `office2llm` is not found after install, add `~/.local/bin` to your PATH:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

## CLI reference

```
office2llm --input <file> [--outdir <dir>] [--dpi <int>] [--timeout-s <int>]
```

| Flag | Default | Description |
| --- | --- | --- |
| `--input` | *(required)* | Path to input file |
| `--outdir` | sibling folder named after input | Where to write PNGs and text files |
| `--dpi` | `200` | Render resolution (higher = sharper but larger files) |
| `--timeout-s` | `120` | Max seconds for the LibreOffice conversion step |

## Use cases

### Convert a Word doc (auto-named output folder)

```bash
export GEMINI_API_KEY="your-api-key-here"
office2llm --input /path/to/report.docx
# -> /path/to/report/page_0001.png, page_0001.txt, page_0002.png, page_0002.txt, …
```

### Convert a PowerPoint deck to a custom folder

```bash
office2llm --input slides.pptx --outdir ./slide-images --dpi 300
```

### Convert a PDF (no LibreOffice needed)

```bash
office2llm --input paper.pdf --outdir ./pages --dpi 250
```

### Convert to a single sibling fulltext file only

```bash
office2llm --input /path/to/example.pdf --fulltext-only
```

This writes:

- `/path/to/example.pdf.txt`

No output folder is kept in this mode.

### Convert a large spreadsheet (increase timeout)

```bash
office2llm --input financials.xlsx --timeout-s 300
```

### Run via Docker (no local install)

```bash
docker build -t office2llm .
docker run --rm -e GEMINI_API_KEY -v "$PWD:/work" office2llm --input /work/in.docx --outdir /work/out
```

Or with Docker Compose:

```bash
docker compose run --rm office2llm --input /data/in.pptx --outdir /data/out --dpi 200
```

## Output

The output directory will contain sequentially numbered PNGs and corresponding OCR text files:

```
page_0001.png
page_0001.txt
page_0002.png
page_0002.txt
page_0003.png
page_0003.txt
…
```

The CLI prints a summary when done:

```
ok pages=8 ocr_ok=8 ocr_skipped=0 ocr_failed=0 outdir=/path/to/output
```

- **ocr_ok**: pages successfully OCR'd in this run
- **ocr_skipped**: pages whose `.txt` already existed (incremental/resumable)
- **ocr_failed**: pages where OCR failed (exit code 2 if any failures)

## Requirements

- **Python** 3.10+
- **Gemini API key** — export `GEMINI_API_KEY` before running
- **LibreOffice** on `PATH` (as `libreoffice` or `soffice`) — only needed for non-PDF inputs
- Python deps (`pypdfium2`, `Pillow`, `google-genai`) are installed automatically
