---
name: office2llm
description: Convert image-free Word files to Markdown and route image-bearing documents and other supported formats through full-page OCR.
---

## What it does

```
Image-free DOCX
  └─ Pandoc ─▶ document.md

Image-free DOC
  └─ LibreOffice ─▶ temporary DOCX ─▶ Pandoc ─▶ document.md

Image-bearing DOCX or other Office doc (.pptx, .xlsx, …)
  └─ LibreOffice (headless) ─▶ PDF ─▶ pypdfium2 ─▶ page_0001.png, page_0002.png, …
                                                  └─ LLM OCR ─▶ page_0001.txt, page_0002.txt, …

PDF (.pdf)
  └─ pypdfium2 (direct, no LibreOffice) ─▶ page_0001.png, page_0002.png, …
                                          └─ LLM OCR ─▶ page_0001.txt, page_0002.txt, …
```

Image-free Word files do not require an API key. Any embedded image routes the complete document through OCR so each image remains in its rendered page context.

All OCR-path PNGs are deterministic, RGB (no alpha), and optimized for downstream consumption.

Each page is sent to an OCR-capable LLMfor OCR text extraction. The LLM preserves the document's semantic structure — headers, hierarchy, data relationships, lists, tables, key-value pairs, equations, and handwriting — and outputs clean plain text with minimal Markdown for structure (tables, lists).

## Supported input formats

| Category | Extensions |
| --- | --- |
| Word | `.docx`, `.doc`, `.odt`, `.rtf` |
| Slides | `.pptx`, `.ppt`, `.odp` |
| Sheets | `.xlsx`, `.xls`, `.ods` |
| PDF | `.pdf` (rendered directly, LibreOffice not needed) |
| Images | `.png`, `.jpg`, `.jpeg`, `.gif`, `.webp`, `.tif`, `.tiff` |

Anything LibreOffice can open will work — the list above covers the most common cases.

## Quick start

```bash
# 1. Install (macOS or Linux — installs Pandoc, LibreOffice, and creates a venv)
bash ./install.sh

# 2. Convert an image-free Word file without an API key
office2llm --input report.docx
```

Export `GEMINI_API_KEY` before processing a Word file with embedded images or another input that requires OCR.

If `office2llm` is not found after install, add `~/.local/bin` to your PATH:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

## CLI reference

```
office2llm --input <file-or-folder> [--outdir <dir>] [--dpi <int>] [--timeout-s <int>] [--fulltext-only]
```

| Flag | Default | Description |
| --- | --- | --- |
| `--input` | *(required)* | Path to input file |
| `--outdir` | input-dependent | Where to write Markdown or OCR page artifacts |
| `--dpi` | `200` | Render resolution (higher = sharper but larger files) |
| `--timeout-s` | `120` | Max seconds for the LibreOffice conversion step |
| `--fulltext-only` | `false` | Write one sibling output: `.md` for native Word or `.txt` for OCR |

## Use cases

### Convert an image-free Word document

```bash
office2llm --input /path/to/report.docx
# -> /path/to/report.md
```

If the DOCX contains an embedded image, the same command requires `GEMINI_API_KEY` and uses the full-page OCR outputs instead.

Legacy `.doc` files use LibreOffice as a temporary bridge and otherwise follow the same routing:

```bash
office2llm --input /path/to/legacy-report.doc
```

### Convert a PowerPoint deck to a custom folder

```bash
office2llm --input slides.pptx --outdir ./slide-images --dpi 300
```

### Convert a PDF (no LibreOffice needed)

```bash
office2llm --input paper.pdf --outdir ./pages --dpi 250
```

### Convert an image

```bash
office2llm --input /path/to/photo.jpg --fulltext-only
```

### Convert to a single sibling fulltext file only

```bash
office2llm --input /path/to/example.pdf --fulltext-only
```

This writes:

- `/path/to/example.pdf.txt`

No output folder is kept in this mode.

### Convert every eligible document in a folder

```bash
yes | office2llm --input /path/to/folder
```

This confirms the interactive prompt automatically and writes sibling outputs like:

- `/path/to/folder/example.pdf.txt`
- `/path/to/folder/report.md` for an image-free DOCX or DOC
- `/path/to/folder/illustrated.docx.txt` for an OCR-routed Word file
- `/path/to/folder/photo.jpg.txt`

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

An image-free Word file produces one Markdown file. OCR-routed inputs produce sequentially numbered PNGs and corresponding OCR text files:

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
- **Pandoc** on `PATH` — required for image-free Word conversion
- **Gemini API key** — required only for OCR-routed inputs
- **LibreOffice** on `PATH` (as `libreoffice` or `soffice`) — required for Office inputs routed to OCR
- Python deps (`pypdfium2`, `Pillow`, `google-genai`) are installed automatically
