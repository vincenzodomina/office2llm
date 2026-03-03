---
name: office2llm
description: Turn any Office document or PDF into one PNG per page for OCR / Vision model input for advanced document understanding.
---

## What it does

```
Office doc (.docx, .pptx, .xlsx, …)
  └─ LibreOffice (headless) ─▶ PDF ─▶ pypdfium2 ─▶ page_0001.png, page_0002.png, …

PDF (.pdf)
  └─ pypdfium2 (direct, no LibreOffice) ─▶ page_0001.png, page_0002.png, …
```

All PNGs are deterministic, RGB (no alpha), and optimized for downstream OCR / vision-model consumption.

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

# 2. Convert a document
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
| `--outdir` | sibling folder named after input | Where to write the PNGs |
| `--dpi` | `200` | Render resolution (higher = sharper but larger files) |
| `--timeout-s` | `120` | Max seconds for the LibreOffice conversion step |

## Use cases

### Convert a Word doc (auto-named output folder)

```bash
office2llm --input /path/to/report.docx
# -> /path/to/report/page_0001.png, page_0002.png, …
```

### Convert a PowerPoint deck to a custom folder

```bash
office2llm --input slides.pptx --outdir ./slide-images --dpi 300
```

### Convert a PDF (no LibreOffice needed)

```bash
office2llm --input paper.pdf --outdir ./pages --dpi 250
```

### Convert a large spreadsheet (increase timeout)

```bash
office2llm --input financials.xlsx --timeout-s 300
```

### Run via Docker (no local install)

```bash
docker build -t office2llm .
docker run --rm -v "$PWD:/work" office2llm --input /work/in.docx --outdir /work/out
```

Or with Docker Compose:

```bash
docker compose run --rm office2llm --input /data/in.pptx --outdir /data/out --dpi 200
```

## Output

The output directory will contain sequentially numbered PNGs:

```
page_0001.png
page_0002.png
page_0003.png
…
```

The CLI prints a summary when done:

```
ok pages=8 outdir=/path/to/output
```

## Requirements

- **Python** 3.10+
- **LibreOffice** on `PATH` (as `libreoffice` or `soffice`) — only needed for non-PDF inputs
- Python deps (`pypdfium2`, `Pillow`) are installed automatically
