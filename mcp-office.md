# MCP Server for Reading Microsoft Office Files and PDFs

## Overview

Build a fully functional **Model Context Protocol (MCP) server** in **Node.js (TypeScript)** that exposes tools for reading and extracting structured content from Microsoft Office files and PDFs — specifically `.docx` (Word), `.pptx` (PowerPoint), `.xlsx` (Excel), and `.pdf`. The server integrates with an accompanying **web UI** (React + Express) that lets users browse files, chat with an LLM, and receive streamed, tool-augmented responses.

A core design goal is handling **long documents without exceeding LLM context limits**: every tool supports selective extraction, and the backend enforces per-result truncation and conversation pruning automatically.

---

## Goals

- Expose MCP tools that an LLM can call to read Office files and PDFs from disk
- Extract rich, structured content from each file type (text, tables, metadata, slide content, cell values, PDF pages)
- Return clean, well-structured JSON or Markdown output suitable for LLM consumption
- Handle both native (text-based) PDFs and scanned (image-based) PDFs via OCR fallback
- Handle errors gracefully (missing files, unsupported formats, corrupt files, password-protected files)
- **Overcome LLM context-length limits** via selective extraction parameters and automatic context guards
- Provide a web UI with file browsing, streaming chat, and **print-to-PDF** conversation export

---

## Tech Stack

| Layer | Choice |
|---|---|
| Language | TypeScript (Node.js 18+) |
| MCP SDK | `@modelcontextprotocol/sdk` |
| Word (.docx) | `mammoth` (text + HTML extraction) + `turndown` (HTML → Markdown) |
| PowerPoint (.pptx) | Direct XML parsing via `adm-zip` + `fast-xml-parser` |
| Excel (.xlsx) | `xlsx` (SheetJS) |
| PDF (native) | `pdf-parse` (text + metadata extraction) |
| PDF (tables) | `pdfplumber` via Python subprocess |
| PDF (scanned/OCR) | `tesseract.js` with `pdf2pic` for image conversion |
| Validation | `zod` |
| Web UI | React + TypeScript (Vite) |
| Backend | Express (Node.js) with SSE streaming |
| LLM client | OpenAI SDK (`openai`) |
| Build | `tsx` for dev, `tsc` for prod |

---

## Project Structure

```
MCP_microsoft/
├── mcp-office-reader/              # Standalone MCP server
│   └── src/
│       ├── index.ts                # MCP server entry point
│       ├── tools/
│       │   ├── readWord.ts         # .docx reader tool
│       │   ├── readPowerPoint.ts   # .pptx reader tool
│       │   ├── readExcel.ts        # .xlsx reader tool
│       │   ├── readPdf.ts          # .pdf reader tool
│       │   └── listOfficeFiles.ts  # File lister tool
│       ├── utils/
│       │   ├── fileUtils.ts        # Path resolution, file validation
│       │   ├── xmlUtils.ts         # XML parsing helpers
│       │   └── ocrUtils.ts         # Tesseract OCR helpers
│       └── types.ts                # Shared TypeScript types
│
└── mcp-office-ui/                  # Web UI + Express backend
    ├── backend/
    │   └── src/
    │       ├── index.ts            # Express server setup
    │       ├── llm.ts              # OpenAI client, tool definitions, agentic loop, context guards
    │       ├── tools/
    │       │   └── executor.ts     # Plain async tool implementations (mirrors MCP tools)
    │       └── routes/
    │           ├── chat.ts         # SSE streaming endpoint
    │           ├── files.ts        # File listing endpoint
    │           └── models.ts       # Model selection endpoint
    └── frontend/
        └── src/
            ├── App.tsx             # Main layout: topbar, sidebar, chat, print button
            ├── hooks/useChat.ts    # React hook for SSE stream parsing
            ├── types.ts            # Frontend TypeScript types
            ├── styles/index.css    # App styles + @media print rules
            └── components/
                ├── MessageBubble.tsx
                ├── Sidebar.tsx
                └── ToolCallCard.tsx
```

---

## MCP Tools

### 1. `read_word_document`

**Description:** Extracts text content, headings, tables, and metadata from a `.docx` file. Supports selective section extraction to stay within LLM context limits.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string(),
  output_format: z.enum(["markdown", "plain_text", "json"]).default("markdown"),
  include_metadata: z.boolean().default(true),
  outline_only: z.boolean().default(false)
    .describe("Return only the numbered section-heading outline + word count. Use this first on long documents."),
  sections: z.array(z.number().int().positive()).optional()
    .describe("1-indexed section numbers to extract (sections delimited by any heading level). Call outline_only first to discover sections."),
})
```

**Selective extraction workflow:**
1. Call with `outline_only: true` → receive list of numbered sections + word count
2. Call with `sections: [2, 4]` → receive only those sections' content

**Expected output (normal mode):**
```json
{
  "metadata": {
    "title": "Project Report",
    "author": "Jane Doe",
    "word_count": 1520,
    "section_count": 8
  },
  "content": "# Introduction\n\nThis report covers...",
  "tables": [{ "index": 0, "rows": [["Header 1", "Header 2"], ["A", "B"]] }]
}
```

**Expected output (`outline_only: true`):**
```json
{
  "metadata": { "word_count": 12400, "section_count": 8 },
  "content": "",
  "sections": [
    { "index": 0, "heading": "Introduction" },
    { "index": 1, "heading": "Background" },
    ...
  ]
}
```

**Implementation notes:**
- Primary extraction via `mammoth` → HTML → `turndown` → Markdown
- Section splitting: post-process Markdown by splitting on any `#`–`######` heading line
- Sections are 1-indexed in all parameters and error messages
- Tables extracted directly from `word/document.xml` via `fast-xml-parser` (`<w:tbl>` elements)
- Metadata from `docProps/core.xml`

---

### 2. `read_powerpoint`

**Description:** Extracts slide content, speaker notes, and structure from a `.pptx` file.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string(),
  slides: z.array(z.number().int().positive()).optional()
    .describe("1-indexed slide numbers. Fetch slide 1 first to see total slide_count in metadata."),
  include_notes: z.boolean().default(true),
  include_shapes: z.boolean().default(false),
})
```

**Selective extraction workflow:**
1. Call with `slides: [1]` → receive first slide + `metadata.slide_count`
2. Call with `slides: [3, 5, 7]` → receive only those slides

**Expected output:**
```json
{
  "metadata": { "title": "Q1 Strategy", "slide_count": 12, "author": "John Smith" },
  "slides": [
    { "index": 1, "title": "Executive Summary", "content": "Key achievements:\n- Revenue up 15%", "notes": "Emphasise Acme Corp." }
  ]
}
```

**Implementation notes:**
- `.pptx` is a ZIP archive; unzip with `adm-zip`
- Slide order from `ppt/_rels/presentation.xml.rels` — sort by rId
- Title from `<p:ph type="title"/>` placeholders; body from all other `<a:t>` nodes
- Speaker notes from `ppt/notesSlides/notesSlide*.xml` (resolved via slide `.rels` file)

---

### 3. `read_excel`

**Description:** Extracts cell data, sheet names, and formulas from a `.xlsx` file. Supports sheet selection, row windowing, and column filtering to handle large workbooks.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string(),
  sheets: z.array(z.string()).optional()
    .describe("Sheet names to read. Omit for all sheets."),
  max_rows: z.number().int().positive().default(1000)
    .describe("Maximum data rows per sheet. Used as the default window size when start_row/end_row are omitted."),
  start_row: z.number().int().positive().optional()
    .describe("1-indexed starting data row (after the header). Use with end_row to page through large sheets."),
  end_row: z.number().int().positive().optional()
    .describe("1-indexed ending data row (after the header)."),
  columns: z.array(z.string()).optional()
    .describe("Column names to include (requires header_row: true). Preview with max_rows: 5 first."),
  include_formulas: z.boolean().default(false),
  header_row: z.boolean().default(true),
})
```

**Selective extraction workflow:**
1. Call with `max_rows: 5` → preview headers and first 5 rows across all sheets
2. Call with `sheets: ["Sales"], columns: ["Date", "Revenue"], start_row: 1, end_row: 500` → targeted slice

**Expected output:**
```json
{
  "metadata": { "sheet_names": ["Summary", "Raw Data"], "file_size_bytes": 48230 },
  "sheets": {
    "Summary": {
      "headers": ["Region", "Q1", "Q2"],
      "rows": [["APAC", 120000, 135000]],
      "row_count": 1,
      "col_count": 3
    }
  }
}
```

**Implementation notes:**
- SheetJS (`xlsx`): `XLSX.readFile()` then manual cell-by-cell iteration to support row windowing
- Row range: `start_row`/`end_row` are 1-indexed data rows; header row is always read first when `header_row: true`
- Column filter: match column names from the header row and include only matching column indices
- If `columns` lists a name not found in headers, it is silently ignored (no error)

---

### 4. `read_pdf`

**Description:** Extracts text, tables, and metadata from a `.pdf` file. Handles both native (text-based) and scanned (image-based) PDFs with automatic OCR fallback.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string(),
  pages: z.array(z.number().int().positive()).optional()
    .describe("1-indexed page numbers. Fetch page 1 first to see total page_count in metadata."),
  output_format: z.enum(["markdown", "plain_text", "json"]).default("markdown"),
  extract_tables: z.boolean().default(false),
  ocr_fallback: z.boolean().default(true),
  include_metadata: z.boolean().default(true),
})
```

**Selective extraction workflow:**
1. Call with `pages: [1], include_metadata: true` → receive page 1 + `metadata.page_count`
2. Call with `pages: [5, 6, 7]` → receive only those pages

**Expected output:**
```json
{
  "metadata": {
    "title": "Annual Report 2024",
    "page_count": 24,
    "file_size_bytes": 1245184,
    "is_scanned": false,
    "encrypted": false
  },
  "pages": [
    { "index": 1, "content": "## Executive Summary\n\n...", "word_count": 312, "ocr_used": false }
  ]
}
```

**Implementation notes:**
- `pdf-parse` with custom `pagerender` callback to capture per-page text
- Scanned PDF detection: average text length < 50 chars/page → trigger OCR
- OCR: `pdf2pic` → PNG → `tesseract.js`, processed in batches of 4 pages
- Table extraction: Python subprocess running `pdfplumber` via `scripts/extract_pdf_tables.py`

---

### 5. `list_office_files` (utility)

**Description:** Lists all `.docx`, `.pptx`, `.xlsx`, and `.pdf` files in a given directory.

**Input schema:**
```ts
z.object({
  directory: z.string(),
  recursive: z.boolean().default(false),
  file_types: z.array(z.enum(["docx", "pptx", "xlsx", "pdf"])).default(["docx", "pptx", "xlsx", "pdf"]),
})
```

---

## Context-Length Protection

### Problem
LLM APIs have a fixed context window (e.g., 131 072 tokens for GPT-4o). Large documents or multi-turn conversations with repeated tool results can easily exceed this limit.

### Solution: two-layer defence in `llm.ts`

#### Layer 1 — Per-result truncation (`truncateToolResult`)
- Every tool result is capped at **120 000 characters** (~30 K tokens)
- If a result exceeds this, it is sliced at the nearest newline boundary and a notice is appended:
  ```
  [RESULT TRUNCATED — Use selective parameters to reduce the payload:
   outline_only/sections (Word), pages (PDF), slides (PowerPoint),
   sheets/start_row/end_row/columns/max_rows (Excel).]
  ```

#### Layer 2 — Conversation pruning (`pruneConversation`)
- Before each API call, estimates total character size of `conversationMessages`
- If size exceeds **360 000 characters** (~90 K tokens), replaces content of the oldest `tool` messages with a stub: `[Tool result omitted from context to stay within token limit]`
- The system message and the 4 most recent messages are always preserved

### System prompt
A `SYSTEM_PROMPT` constant is prepended to every conversation, instructing the LLM to:
1. Call `outline_only: true` (Word) or fetch slide/page 1 (PowerPoint/PDF) first to discover document size
2. Then use selective parameters (`sections`, `pages`, `slides`, `start_row`/`end_row`, `columns`) to fetch only the relevant portion
3. Work iteratively through long documents to stay within context limits

---

## Error Handling Requirements

Every tool must handle and return descriptive errors for:

- File not found at the given path
- File extension mismatch (e.g., a `.pdf` passed to `read_word_document`)
- Corrupt or unreadable ZIP archive (Office files are ZIPs)
- Password-protected files — return a clear message; do not attempt to crack
- Empty files or files with no extractable content
- Sections / slides / pages / sheets out of range — include total count in the error message
- Scanned PDFs with OCR disabled — return a warning suggesting `ocr_fallback: true`
- Python/pdfplumber not available when `extract_tables: true` — return actionable install message

All errors are returned as MCP tool errors (not thrown exceptions).

---

## Web UI Requirements

### Features
- **File browser sidebar**: directory path input, recursive toggle, sort by type/date/size, multi-select with checkboxes
- **Chat interface**: streaming assistant responses via SSE, tool call cards (expandable, show arguments + result)
- **Model selector**: dynamically fetches available models from `/api/models`
- **Send / Stop generation**: abort controller for cancelling in-flight requests
- **Clear chat**: resets the conversation history
- **Print conversation to PDF**: a "🖨 Print" button in the top bar calls `window.print()`, which opens the browser's native print/save-as-PDF dialog

### Print-to-PDF behaviour
- Triggered by the **🖨 Print** button in the topbar (`window.print()`)
- `@media print` CSS rules:
  - Hide topbar, sidebar, input area, suggestions, stop button
  - Remove scroll constraints (`overflow: visible`, `height: auto`) so the full conversation renders
  - Force a light colour scheme (white background, dark text) for readability on paper
  - Prepend a "MCP Office Reader — Conversation" page heading via `::before`
  - Expand all tool-call cards (show body regardless of toggle state)
  - Remove code-block max-height so all content is visible
  - Hide streaming cursors and thinking dots
  - Page-break-inside: avoid on each message wrapper

### Backend API endpoints
| Endpoint | Method | Description |
|---|---|---|
| `POST /api/chat` | POST | SSE stream; accepts `{ messages, model }` |
| `GET /api/files` | GET | Lists Office/PDF files; accepts `?directory=&recursive=` |
| `GET /api/models` | GET | Returns available LLM model IDs |

---

## `package.json` Dependencies

### `mcp-office-reader`
```json
{
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.0.0",
    "adm-zip": "^0.5.10",
    "fast-xml-parser": "^4.3.0",
    "mammoth": "^1.7.0",
    "pdf-parse": "^1.1.1",
    "pdf2pic": "^3.1.1",
    "tesseract.js": "^5.0.0",
    "turndown": "^7.1.3",
    "xlsx": "^0.18.5",
    "zod": "^3.22.0"
  }
}
```

### `mcp-office-ui/backend`
```json
{
  "dependencies": {
    "adm-zip": "^0.5.10",
    "express": "^4.18.0",
    "fast-xml-parser": "^4.3.0",
    "mammoth": "^1.7.0",
    "openai": "^4.0.0",
    "pdf-parse": "^1.1.1",
    "pdf2pic": "^3.1.1",
    "tesseract.js": "^5.0.0",
    "turndown": "^7.1.3",
    "xlsx": "^0.18.5"
  }
}
```

**Python dependency** (for PDF table extraction):
```bash
pip install pdfplumber
```

---

## Claude Desktop Integration (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "office-reader": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-office-reader/dist/index.js"]
    }
  }
}
```

---

## Testing Checklist

### Core extraction
- [ ] `read_word_document` returns correct headings and table content from a sample `.docx`
- [ ] `read_word_document` correctly extracts metadata (author, dates, word count)
- [ ] `read_powerpoint` returns all slides in correct order
- [ ] `read_powerpoint` correctly extracts speaker notes
- [ ] `read_excel` reads multi-sheet workbooks correctly
- [ ] `read_excel` respects `max_rows` limit
- [ ] `read_pdf` extracts text and metadata from a native PDF
- [ ] `read_pdf` correctly handles page-range selection
- [ ] `read_pdf` detects a scanned PDF and triggers OCR when `ocr_fallback: true`
- [ ] `read_pdf` with `extract_tables: true` returns table data via pdfplumber subprocess
- [ ] `read_pdf` returns a clear error for password-protected PDFs
- [ ] All tools return structured errors for missing/corrupt files
- [ ] `list_office_files` finds all file types recursively

### Selective extraction (context-length features)
- [ ] `read_word_document` with `outline_only: true` returns section list without body content
- [ ] `read_word_document` with `sections: [2, 4]` returns only those sections
- [ ] `read_word_document` with out-of-range section returns error including total section count
- [ ] `read_excel` with `start_row: 100, end_row: 200` returns only that row window
- [ ] `read_excel` with `columns: ["Revenue"]` returns only that column
- [ ] `read_powerpoint` with `slides: [1]` returns slide_count in metadata
- [ ] `read_pdf` with `pages: [1]` returns page_count in metadata

### Context guards
- [ ] Tool result > 120 000 chars is truncated with hint message
- [ ] Conversation > 360 000 chars causes old tool results to be pruned before API call
- [ ] No 400 context-length errors when reading large documents

### Web UI
- [ ] File sidebar lists and selects `.docx`, `.pptx`, `.xlsx`, `.pdf` files
- [ ] Streaming responses render incrementally with tool call cards
- [ ] Print button opens browser print dialog showing conversation only (no sidebar/input)
- [ ] Print output uses light colour scheme with page header
- [ ] All tool cards are expanded in print output
- [ ] Clear chat resets conversation history

---

## Stretch Goals (Optional)

- Add a `convert_to_pdf` tool using LibreOffice headless (`soffice --headless --convert-to pdf`)
- Add image extraction from Word/PowerPoint (images in `media/image*.png` inside the ZIP)
- Add a `summarize_document` tool that pipes extracted content into an LLM call
- Support `.xls` (legacy Excel) via SheetJS compatibility layer
- Support `.doc` (legacy Word) via LibreOffice conversion to `.docx` first
- Add a `merge_and_read` tool that accepts a list of mixed file paths and returns unified content
- Add client-side PDF generation (e.g., using `jsPDF` + `html2canvas`) as an alternative to browser print
