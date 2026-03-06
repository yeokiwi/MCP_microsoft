# MCP Office UI — DeepSeek Chat Interface

A full-stack web application that integrates the **MCP Office Reader** tools with **DeepSeek LLM**, providing a chat interface for reading and analysing Office files and PDFs.

## Architecture

```
mcp-office-ui/
├── backend/         Express API server — DeepSeek integration + tool execution
└── frontend/        React + Vite chat UI
```

```
Browser (React)
    │  SSE stream
    ▼
Express Backend (port 3001)
    │  OpenAI-compatible API  ──→  DeepSeek LLM
    │  Tool calls
    ▼
Office Tool Executor
    ├── mammoth       (.docx)
    ├── adm-zip + fast-xml-parser  (.pptx)
    ├── SheetJS       (.xlsx)
    ├── pdf-parse     (.pdf native)
    ├── tesseract.js  (.pdf scanned/OCR)
    └── pdfplumber    (.pdf tables via Python)
```

## Quick Start

### 1. Configure DeepSeek API key

```bash
cp backend/.env.example backend/.env
# Edit backend/.env and set your DEEPSEEK_API_KEY
```

Get a key at https://platform.deepseek.com

### 2. Install dependencies

```bash
cd backend && npm install
cd ../frontend && npm install
```

### 3. Start backend

```bash
cd backend
npm run dev        # development (tsx watch)
# or
npm run build && npm start   # production
```

### 4. Start frontend

```bash
cd frontend
npm run dev        # http://localhost:5173
```

### Optional: PDF table extraction

```bash
pip install pdfplumber
```

## Features

- **Chat with DeepSeek** about any Office file or PDF
- **Agentic tool use** — DeepSeek automatically calls the right tool based on your question
- **Streaming responses** via Server-Sent Events
- **File browser sidebar** — scan directories for Office files and click to analyse
- **Tool call inspector** — expandable cards show what tools were called and the raw results
- **Auto OCR** — scanned PDFs are automatically processed with Tesseract
- **Dark / light theme** — follows system preference

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/chat` | Stream a chat completion with tool use |
| GET | `/api/files?directory=&recursive=` | List Office files in a directory |
| GET | `/api/health` | Health check |

## Tech Stack

| Layer | Tech |
|-------|------|
| LLM | DeepSeek (OpenAI-compatible API) |
| Backend | Express + TypeScript |
| Frontend | React 18 + Vite + TypeScript |
| Streaming | Server-Sent Events (SSE) |
| Markdown | react-markdown + remark-gfm |
