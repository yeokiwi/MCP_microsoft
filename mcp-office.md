# MCP Server for Reading Microsoft Office Files and PDFs

## Overview

Build a fully functional **Model Context Protocol (MCP) server** in **Node.js (TypeScript)** that exposes tools for reading and extracting structured content from Microsoft Office files and PDFs — specifically `.docx` (Word), `.pptx` (PowerPoint), `.xlsx` (Excel), and `.pdf`. The server should integrate seamlessly with any MCP-compatible client (e.g., Claude Desktop, Cursor, or a custom LLM app).

---

## Goals

- Expose MCP tools that an LLM can call to read Office files and PDFs from disk
- Extract rich, structured content from each file type (text, tables, metadata, slide content, cell values, PDF pages)
- Return clean, well-structured JSON or Markdown output suitable for LLM consumption
- Handle both native (text-based) PDFs and scanned (image-based) PDFs via OCR fallback
- Handle errors gracefully (missing files, unsupported formats, corrupt files, password-protected PDFs)

---

## Tech Stack

| Layer | Choice |
|---|---|
| Language | TypeScript (Node.js 18+) |
| MCP SDK | `@modelcontextprotocol/sdk` |
| Word (.docx) | `mammoth` (text + HTML extraction) |
| PowerPoint (.pptx) | Direct XML parsing via `adm-zip` + `fast-xml-parser` |
| Excel (.xlsx) | `xlsx` (SheetJS) |
| PDF (native) | `pdf-parse` (text + metadata extraction) |
| PDF (tables) | `pdfplumber` via Python subprocess (best-in-class table detection) |
| PDF (scanned/OCR) | `tesseract.js` with `pdf2pic` for image conversion |
| Validation | `zod` |
| Build | `tsx` for dev, `tsc` for prod |

---

## Project Structure

```
mcp-office-reader/
├── src/
│   ├── index.ts              # MCP server entry point
│   ├── tools/
│   │   ├── readWord.ts       # .docx reader tool
│   │   ├── readPowerPoint.ts # .pptx reader tool
│   │   ├── readExcel.ts      # .xlsx reader tool
│   │   └── readPdf.ts        # .pdf reader tool
│   ├── utils/
│   │   ├── fileUtils.ts      # Path resolution, MIME detection
│   │   ├── xmlUtils.ts       # XML parsing helpers for pptx/docx
│   │   └── ocrUtils.ts       # Tesseract OCR helpers for scanned PDFs
│   └── types.ts              # Shared TypeScript types
├── scripts/
│   └── extract_pdf_tables.py # Python pdfplumber script for table extraction
├── package.json
├── tsconfig.json
└── README.md
```

---

## MCP Tools to Implement

### 1. `read_word_document`

**Description:** Extracts text content, headings, tables, and metadata from a `.docx` file.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string().describe("Absolute or relative path to the .docx file"),
  output_format: z.enum(["markdown", "plain_text", "json"]).default("markdown")
    .describe("Format of the extracted content"),
  include_metadata: z.boolean().default(true)
    .describe("Whether to include document metadata (author, dates, word count)"),
})
```

**Expected output:**
```json
{
  "metadata": {
    "title": "Project Report",
    "author": "Jane Doe",
    "created": "2024-01-15T10:30:00Z",
    "modified": "2024-03-01T14:22:00Z",
    "word_count": 1520
  },
  "content": "# Introduction\n\nThis report covers...",
  "tables": [
    {
      "index": 0,
      "rows": [
        ["Header 1", "Header 2"],
        ["Value A", "Value B"]
      ]
    }
  ]
}
```

**Implementation notes:**
- Use `mammoth` for primary text/HTML extraction, then convert HTML to Markdown using `turndown`
- For raw XML access (e.g., tracked changes), unzip the `.docx` with `adm-zip` and parse `word/document.xml` with `fast-xml-parser`
- Extract metadata from `docProps/core.xml` inside the zip
- Tables should be parsed from `<w:tbl>` XML elements

---

### 2. `read_powerpoint`

**Description:** Extracts slide content, speaker notes, and structure from a `.pptx` file.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string().describe("Absolute or relative path to the .pptx file"),
  slides: z.array(z.number().int().positive()).optional()
    .describe("Optional: specific 1-indexed slide numbers to extract. Omit for all slides."),
  include_notes: z.boolean().default(true)
    .describe("Whether to include speaker notes for each slide"),
  include_shapes: z.boolean().default(false)
    .describe("Whether to include shape/element metadata per slide"),
})
```

**Expected output:**
```json
{
  "metadata": {
    "title": "Q1 Strategy Deck",
    "slide_count": 12,
    "author": "John Smith"
  },
  "slides": [
    {
      "index": 1,
      "title": "Executive Summary",
      "content": "Key achievements this quarter:\n- Revenue up 15%\n- 3 new partnerships",
      "notes": "Emphasise the partnership with Acme Corp.",
      "shapes": []
    }
  ]
}
```

**Implementation notes:**
- `.pptx` files are ZIP archives. Unzip with `adm-zip` and iterate over `ppt/slides/slide*.xml`
- Slide order is determined by `ppt/slides/_rels/` relationship files — respect this ordering
- Extract title from `<p:sp>` elements with `<p:ph type="title"/>` placeholder
- Extract body text from all other `<a:t>` (text run) nodes
- Speaker notes are in `ppt/notesSlides/notesSlide*.xml`
- Parse with `fast-xml-parser`, configured with `ignoreAttributes: false`

---

### 3. `read_excel`

**Description:** Extracts cell data, sheet names, and optionally named ranges from a `.xlsx` file.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string().describe("Absolute or relative path to the .xlsx file"),
  sheets: z.array(z.string()).optional()
    .describe("Optional: specific sheet names to read. Omit for all sheets."),
  max_rows: z.number().int().positive().default(1000)
    .describe("Maximum rows to read per sheet (default 1000)"),
  include_formulas: z.boolean().default(false)
    .describe("If true, return raw formulas instead of computed values where available"),
  header_row: z.boolean().default(true)
    .describe("Whether to treat the first row as column headers"),
})
```

**Expected output:**
```json
{
  "metadata": {
    "sheet_names": ["Summary", "Raw Data", "Charts"],
    "file_size_bytes": 48230
  },
  "sheets": {
    "Summary": {
      "headers": ["Region", "Q1", "Q2", "Q3", "Q4"],
      "rows": [
        ["APAC", 120000, 135000, 142000, 160000],
        ["EMEA", 98000, 102000, 115000, 121000]
      ],
      "row_count": 2,
      "col_count": 5
    }
  }
}
```

**Implementation notes:**
- Use `SheetJS` (`xlsx` package): `XLSX.readFile(path)` then `XLSX.utils.sheet_to_json()`
- Pass `{ header: 1 }` to `sheet_to_json` for raw 2D array output, then split headers from rows
- Use `{ raw: false }` to get formatted cell values, or `{ raw: true }` for raw numbers/dates
- For formulas, read `ws[cell].f` from the cell object directly
- Detect merged cells via `ws['!merges']` and note them in the output if present

---

### 4. `read_pdf`

**Description:** Extracts text, tables, and metadata from a `.pdf` file. Handles both native (text-based) and scanned (image-based) PDFs, with automatic detection and OCR fallback.

**Input schema (Zod):**
```ts
z.object({
  file_path: z.string().describe("Absolute or relative path to the .pdf file"),
  pages: z.array(z.number().int().positive()).optional()
    .describe("Optional: specific 1-indexed page numbers to extract. Omit for all pages."),
  output_format: z.enum(["markdown", "plain_text", "json"]).default("markdown")
    .describe("Format of the extracted content"),
  extract_tables: z.boolean().default(false)
    .describe("Whether to extract tables (uses Python pdfplumber subprocess — requires Python 3 + pdfplumber installed)"),
  ocr_fallback: z.boolean().default(true)
    .describe("If text extraction yields empty results (scanned PDF), automatically run Tesseract OCR"),
  include_metadata: z.boolean().default(true)
    .describe("Whether to include document metadata (title, author, page count, etc.)"),
})
```

**Expected output:**
```json
{
  "metadata": {
    "title": "Annual Report 2024",
    "author": "Finance Team",
    "page_count": 24,
    "file_size_bytes": 1245184,
    "pdf_version": "1.6",
    "is_scanned": false,
    "encrypted": false
  },
  "pages": [
    {
      "index": 1,
      "content": "## Executive Summary\n\nThis report presents the consolidated financial results...",
      "word_count": 312,
      "ocr_used": false
    },
    {
      "index": 2,
      "content": "## Revenue Breakdown\n\nRevenue grew 18% year-on-year...",
      "word_count": 275,
      "ocr_used": false
    }
  ],
  "tables": [
    {
      "page": 3,
      "index": 0,
      "rows": [
        ["Region", "Q1", "Q2", "Q3", "Q4"],
        ["APAC",   "120K", "135K", "142K", "160K"]
      ]
    }
  ]
}
```

**Implementation notes:**

*Text extraction (primary path):*
- Use `pdf-parse`: `const data = await pdfParse(fs.readFileSync(filePath))`
- `data.text` gives full concatenated text; `data.numpages` gives page count
- `data.info` contains metadata (Title, Author, CreationDate, etc.)
- `data.version` gives PDF spec version
- To extract text per page, use `pdf-parse` with a custom `pagerender` callback:
  ```ts
  const pageTexts: string[] = [];
  await pdfParse(buffer, {
    pagerender: (pageData) => {
      return pageData.getTextContent().then((content: any) => {
        pageTexts.push(content.items.map((i: any) => i.str).join(" "));
        return "";
      });
    }
  });
  ```

*Scanned PDF detection and OCR fallback:*
- After extraction, check if total text length is below a threshold (e.g., fewer than 50 characters per page on average) — this signals a scanned PDF
- If `ocr_fallback: true` and the PDF appears scanned, convert pages to images using `pdf2pic`:
  ```ts
  import { fromPath } from "pdf2pic";
  const convert = fromPath(filePath, { density: 300, format: "png", width: 2480, height: 3508 });
  const pageImage = await convert(pageNumber);
  ```
- Run Tesseract OCR on each image using `tesseract.js`:
  ```ts
  import Tesseract from "tesseract.js";
  const { data: { text } } = await Tesseract.recognize(pageImage.path, "eng");
  ```
- Set `ocr_used: true` on any page processed this way
- Note: OCR is slow (~2–5 seconds per page). For large documents, process pages in parallel with a concurrency limit of 4

*Table extraction (subprocess path):*
- If `extract_tables: true`, spawn a Python subprocess running `scripts/extract_pdf_tables.py`:
  ```ts
  import { execFile } from "child_process";
  const result = await new Promise((resolve, reject) => {
    execFile("python3", ["scripts/extract_pdf_tables.py", filePath], (err, stdout) => {
      if (err) reject(err);
      else resolve(JSON.parse(stdout));
    });
  });
  ```
- The Python script uses `pdfplumber` — the most accurate open-source table extractor available:
  ```python
  # scripts/extract_pdf_tables.py
  import sys, json, pdfplumber

  with pdfplumber.open(sys.argv[1]) as pdf:
      output = []
      for i, page in enumerate(pdf.pages):
          for table in page.extract_tables():
              output.append({ "page": i + 1, "rows": table })
      print(json.dumps(output))
  ```
- Dependency check: verify `python3` and `pdfplumber` are available before calling; return a descriptive error if not

*Password-protected PDFs:*
- `pdf-parse` will throw if the PDF is encrypted with a user password
- Catch this and return: `{ error: "PDF is password-protected. Provide the password to decrypt." }`
- Do not attempt to crack or bypass encryption

---

### 5. `list_office_files` (utility tool)

**Description:** Lists all `.docx`, `.pptx`, `.xlsx`, and `.pdf` files in a given directory.

**Input schema:**
```ts
z.object({
  directory: z.string().describe("Directory path to scan"),
  recursive: z.boolean().default(false),
  file_types: z.array(z.enum(["docx", "pptx", "xlsx", "pdf"])).default(["docx", "pptx", "xlsx", "pdf"])
})
```

---

## MCP Server Entry Point (`src/index.ts`)

```ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerReadWordTool } from "./tools/readWord.js";
import { registerReadPowerPointTool } from "./tools/readPowerPoint.js";
import { registerReadExcelTool } from "./tools/readExcel.js";
import { registerReadPdfTool } from "./tools/readPdf.js";

const server = new McpServer({
  name: "office-reader",
  version: "1.0.0",
});

registerReadWordTool(server);
registerReadPowerPointTool(server);
registerReadExcelTool(server);
registerReadPdfTool(server);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("Office Reader MCP server running on stdio");
```

---

## Error Handling Requirements

Every tool must handle and return descriptive errors for:

- File not found at the given path
- File extension mismatch (e.g., a `.pdf` passed to `read_word_document`)
- Corrupt or unreadable ZIP archive (Office files are ZIPs)
- Password-protected files — detect and return a clear message for both Office files and PDFs (do not attempt to crack)
- Empty files or files with no extractable content
- Sheets not found (for Excel) or slide index out of range (for PowerPoint)
- Scanned PDFs where OCR is disabled but no text could be extracted — return a warning suggesting `ocr_fallback: true`
- Python/pdfplumber not available when `extract_tables: true` is requested on a PDF — return an actionable install message

All errors should be returned as MCP tool errors using `server.tool()` error response convention, not thrown exceptions.

---

## `package.json` Dependencies

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
  },
  "devDependencies": {
    "@types/adm-zip": "^0.5.5",
    "@types/node": "^20.0.0",
    "@types/pdf-parse": "^1.1.4",
    "@types/turndown": "^5.0.4",
    "tsx": "^4.0.0",
    "typescript": "^5.3.0"
  },
  "scripts": {
    "dev": "tsx src/index.ts",
    "build": "tsc",
    "start": "node dist/index.js"
  },
  "type": "module"
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

- [ ] `read_word_document` returns correct headings and table content from a sample `.docx`
- [ ] `read_word_document` correctly extracts metadata (author, dates)
- [ ] `read_powerpoint` returns all slides in correct order
- [ ] `read_powerpoint` correctly extracts speaker notes
- [ ] `read_excel` reads multi-sheet workbooks correctly
- [ ] `read_excel` respects `max_rows` limit
- [ ] `read_pdf` extracts text and metadata from a native (text-based) PDF
- [ ] `read_pdf` correctly handles page-range selection
- [ ] `read_pdf` detects a scanned PDF and triggers OCR when `ocr_fallback: true`
- [ ] `read_pdf` with `extract_tables: true` returns table data via pdfplumber subprocess
- [ ] `read_pdf` returns a clear error for password-protected PDFs
- [ ] All tools return structured errors for missing/corrupt files
- [ ] Server registers and connects cleanly via stdio transport
- [ ] `list_office_files` finds all `.docx`, `.pptx`, `.xlsx`, and `.pdf` files recursively

---

## Stretch Goals (Optional)

- Add a `convert_to_pdf` tool using LibreOffice headless (`soffice --headless --convert-to pdf`) to convert any Office file to PDF
- Add image extraction from Word/PowerPoint (images stored as `media/image*.png` inside the ZIP) and from PDFs (using `pdfimages` from poppler-utils)
- Add a `summarize_document` tool that pipes extracted content into a Claude API call
- Support `.xls` (legacy Excel) via SheetJS's built-in compatibility layer
- Support `.doc` (legacy Word) via LibreOffice conversion to `.docx` first
- Add a `merge_and_read` tool that accepts a list of mixed file paths (`.docx`, `.pdf`, etc.) and returns a unified content response — useful for LLM context assembly
