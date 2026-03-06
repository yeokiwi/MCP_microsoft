# MCP Office Reader

A fully functional **Model Context Protocol (MCP) server** built in TypeScript (Node.js) that exposes tools for reading and extracting structured content from Microsoft Office files and PDFs.

## Supported File Types

- `.docx` — Microsoft Word documents
- `.pptx` — Microsoft PowerPoint presentations
- `.xlsx` — Microsoft Excel workbooks
- `.pdf` — PDF documents (native text + scanned/OCR)

## MCP Tools

### `read_word_document`
Extracts text content, headings, tables, and metadata from a `.docx` file.

**Parameters:**
- `file_path` — Path to the `.docx` file
- `output_format` — `"markdown"` | `"plain_text"` | `"json"` (default: `"markdown"`)
- `include_metadata` — Include author, dates, word count (default: `true`)

### `read_powerpoint`
Extracts slide content, speaker notes, and structure from a `.pptx` file.

**Parameters:**
- `file_path` — Path to the `.pptx` file
- `slides` — Optional: specific 1-indexed slide numbers to extract
- `include_notes` — Include speaker notes (default: `true`)
- `include_shapes` — Include shape/element metadata (default: `false`)

### `read_excel`
Extracts cell data, sheet names, and optionally formulas from a `.xlsx` file.

**Parameters:**
- `file_path` — Path to the `.xlsx` file
- `sheets` — Optional: specific sheet names to read
- `max_rows` — Maximum rows per sheet (default: `1000`)
- `include_formulas` — Return raw formulas instead of values (default: `false`)
- `header_row` — Treat first row as column headers (default: `true`)

### `read_pdf`
Extracts text, tables, and metadata from a `.pdf` file. Handles native and scanned PDFs.

**Parameters:**
- `file_path` — Path to the `.pdf` file
- `pages` — Optional: specific 1-indexed page numbers
- `output_format` — `"markdown"` | `"plain_text"` | `"json"` (default: `"markdown"`)
- `extract_tables` — Extract tables via Python/pdfplumber (default: `false`)
- `ocr_fallback` — Auto-run Tesseract OCR on scanned PDFs (default: `true`)
- `include_metadata` — Include PDF metadata (default: `true`)

### `list_office_files`
Lists all Office files and PDFs in a given directory.

**Parameters:**
- `directory` — Directory path to scan
- `recursive` — Scan subdirectories (default: `false`)
- `file_types` — Types to include: `["docx", "pptx", "xlsx", "pdf"]`

## Installation

```bash
npm install
npm run build
```

### Python dependency (for PDF table extraction)
```bash
pip install pdfplumber
```

## Usage

### Development
```bash
npm run dev
```

### Production
```bash
npm run build
npm start
```

## Claude Desktop Integration

Add to `claude_desktop_config.json`:
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

## Project Structure

```
mcp-office-reader/
├── src/
│   ├── index.ts              # MCP server entry point
│   ├── tools/
│   │   ├── readWord.ts       # .docx reader tool
│   │   ├── readPowerPoint.ts # .pptx reader tool
│   │   ├── readExcel.ts      # .xlsx reader tool
│   │   ├── readPdf.ts        # .pdf reader tool
│   │   └── listOfficeFiles.ts # File listing utility
│   ├── utils/
│   │   ├── fileUtils.ts      # Path resolution, file info
│   │   ├── xmlUtils.ts       # XML parsing helpers
│   │   └── ocrUtils.ts       # Tesseract OCR helpers
│   └── types.ts              # Shared TypeScript types
├── scripts/
│   └── extract_pdf_tables.py # Python pdfplumber table extractor
├── dist/                     # Compiled JavaScript (after build)
├── package.json
└── tsconfig.json
```

## Tech Stack

| Layer | Choice |
|---|---|
| Language | TypeScript (Node.js 18+) |
| MCP SDK | `@modelcontextprotocol/sdk` |
| Word (.docx) | `mammoth` + `turndown` |
| PowerPoint (.pptx) | `adm-zip` + `fast-xml-parser` |
| Excel (.xlsx) | `xlsx` (SheetJS) |
| PDF (native) | `pdf-parse` |
| PDF (tables) | `pdfplumber` via Python subprocess |
| PDF (OCR) | `tesseract.js` + `pdf2pic` |
| Validation | `zod` |

## Error Handling

All tools return descriptive errors for:
- File not found
- Extension mismatch
- Corrupt/invalid archives
- Password-protected files
- Empty files with no content
- Out-of-range page/slide/sheet indices
- Missing Python/pdfplumber dependency
