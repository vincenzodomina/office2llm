## PRD: office2llm (Per-page Images + OCR Text)

### Summary
`office2llm` converts image-free DOCX files directly to Markdown. DOCX files with embedded images and all other supported inputs use per-page rendering and OCR.

### Goals
- **One-command batch conversion** from a single input document to per-page outputs.
- **Native DOCX extraction** through Pandoc without an OCR API key when no images are embedded.
- **High-fidelity OCR text** that captures all readable text and preserves semantic structure (lists, tables, key/value lines).
- **Deterministic, resumable outputs** so repeated runs are safe and predictable.

### Non-goals
- **Interactive UI** for editing or reviewing OCR results.
- **Semantic interpretation** (classification, summarization, or entity extraction).
- **Layout recreation** beyond text structure needed for RAG ingestion.

### Primary users
- **Data/ML engineers** preparing documents for RAG ingestion.
- **Automation scripts** converting document batches in CI or scheduled jobs.

### User experience (happy path)
- The user runs the CLI with an input document path and optionally an output directory.
- An image-free DOCX produces a single Markdown file.
- An OCR-routed input produces an output directory containing:
  - `page_0001.png`, `page_0002.png`, …
  - `page_0001.txt`, `page_0002.txt`, …
- The user ingests the `.txt` files (and optionally keeps `.png` files for traceability/auditing).

### Functional requirements
- **Inputs**
  - The tool shall accept a single local file path as input.
  - The tool shall support common Office document formats and PDFs.
  - The tool shall fail with a clear error message if the input file does not exist or cannot be read.

- **Outputs**
  - Image-free DOCX files shall produce a single sibling Markdown file.
  - The tool shall create an output directory for results (defaulting to a predictable location when not specified).
  - The tool shall produce exactly one page image per page, named `page_XXXX.png` (zero-padded).
  - The tool shall produce exactly one OCR text file per page, named `page_XXXX.txt` (zero-padded).
  - Each `page_XXXX.txt` shall be written alongside its corresponding `page_XXXX.png`.

- **OCR behavior**
  - A DOCX file with any embedded image shall use the full-page OCR path for the complete document.
  - For each page image, the tool shall request OCR text from an external OCR service.
  - The OCR output shall follow a natural human reading order and preserve semantic structure (including tables and lists).
  - The OCR output shall include all readable text visible on the page (including headers/footers/marginalia when readable).
  - The OCR output shall not add invented text, labels, preambles, or commentary.
  - If a page contains no readable text, the tool shall write an empty text file for that page.

- **Resumability**
  - By default, if `page_XXXX.txt` already exists, the tool shall skip generating OCR for that page.

- **Reliability & error handling**
  - The tool shall continue processing remaining pages if OCR fails for a subset of pages.
  - The tool shall print a summary indicating how many pages were produced, processed for OCR, skipped, and failed.
  - The tool shall exit with a non-zero status code if any page failed OCR.

- **Performance**
  - The tool shall process multiple pages concurrently, with a bounded default concurrency level suitable for typical API quotas.

- **Configuration**
  - The tool shall allow users to tune page rendering quality.
  - The tool shall require credentials only when the selected path uses the OCR service.

### Acceptance criteria
- Running on an image-free DOCX without OCR credentials produces a Markdown file through Pandoc.
- Running on an OCR-routed multi-page input produces matching sets of `page_XXXX.png` and `page_XXXX.txt` for every page.
- Re-running the tool on the same output directory does not re-generate existing `page_XXXX.txt` files by default.
- Missing OCR credentials causes a clear, actionable failure only for inputs routed to OCR.
- When OCR requests fail for some pages, the tool completes the rest of the pages, reports failures, and exits non-zero.
