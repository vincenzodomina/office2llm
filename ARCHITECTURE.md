## Architecture: office2llm

### System overview
`office2llm` is a batch-oriented CLI that converts a single input document into per-page artifacts for downstream retrieval workflows:
- **Page images**: `page_XXXX.png`
- **OCR text**: `page_XXXX.txt`

It supports Office formats by converting them into a PDF first, then rendering pages to images and extracting OCR text per page via an external LLM-powered OCR service.

### Tech stack (current)
- **Language/runtime**: Python 3.10+
- **CLI**: Python entrypoint (`office2llm`)
- **Office → PDF conversion**: LibreOffice/soffice (external binary)
- **PDF rendering**: PDFium via `pypdfium2`
- **Image handling**: `Pillow`
- **OCR (LLM)**: Google GenAI SDK (`google-genai`) to an external Gemini OCR-capable model
- **Packaging**: `pyproject.toml` (setuptools)
- **Container option**: Dockerfile (optional runtime)

### High-level data flow

```mermaid
flowchart TD
  U[User] -->|--input FILE| CLI[office2llm CLI]
  CLI --> T{Input type?}
  T -->|Office doc| C[Convert to PDF]
  T -->|PDF| P[Open PDF]
  C --> P
  P --> R[Render each page to PNG]
  R --> W1[Write page_XXXX.png]
  W1 --> O[Request OCR for each page image]
  O --> W2[Write page_XXXX.txt]
  W2 --> S[Print summary + exit status]
```

### Page-level processing flow

```mermaid
flowchart LR
  A[page_XXXX.png exists] --> B{page_XXXX.txt exists?}
  B -->|Yes| C[Skip OCR for this page]
  B -->|No| D[Send image to OCR service]
  D --> E[Receive extracted text]
  E --> F[Atomically write page_XXXX.txt]
```

### Sequence diagram (happy path)

```mermaid
sequenceDiagram
  participant U as User
  participant CLI as office2llm
  participant LO as Office Converter
  participant PDF as PDF Renderer
  participant OCR as LLM OCR Service
  participant FS as Filesystem

  U->>CLI: Run with input path
  alt Input is Office format
    CLI->>LO: Convert to PDF
    LO-->>CLI: PDF
  end
  CLI->>PDF: Load PDF
  loop For each page
    CLI->>PDF: Render page
    PDF-->>CLI: Page image
    CLI->>FS: Write page_XXXX.png
    CLI->>OCR: OCR(page_XXXX.png)
    OCR-->>CLI: Extracted text
    CLI->>FS: Write page_XXXX.txt
  end
  CLI-->>U: Summary + exit status
```

### Key architectural properties
- **Deterministic outputs**: stable naming (`page_XXXX.*`) enables downstream indexing and predictable diffs.
- **Resumability**: existing `page_XXXX.txt` can be treated as “already processed” to make reruns cheap and safe.
- **Batch resilience**: partial OCR failures do not prevent other pages from being processed; failures are surfaced in a final summary and exit status.
- **Bounded parallelism**: OCR is performed concurrently but with a small cap to reduce the risk of quota/rate-limit issues.
- **Atomic writes**: text outputs are written in a way that avoids leaving partially-written files on interruption.

### Runtime dependencies
- **Local execution**
  - Requires LibreOffice installed and available on `PATH` for non-PDF inputs.
  - Requires credentials for the OCR service.
- **Container execution**
  - Docker image bundles LibreOffice and Python dependencies; OCR still requires credentials at runtime.

### External interfaces
- **CLI inputs**
  - Input file path
  - Optional output directory
  - Optional rendering-quality controls
- **CLI outputs**
  - `page_XXXX.png` and `page_XXXX.txt` per page
  - A single-line summary and a process exit code indicating success/partial failure

