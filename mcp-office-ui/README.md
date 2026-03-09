# MCP Office UI — OpenAI-Powered Chat Interface

A full-stack web application that integrates the **MCP Office Reader** tools with any **OpenAI-compatible API**, providing a chat interface for reading and analysing Office files and PDFs.

## Architecture

```
mcp-office-ui/
├── backend/         Express API server — LLM routing + tool execution
└── frontend/        React + Vite chat UI
```

```
Browser (React)
    │  SSE stream
    ▼
Express Backend (port 3001)
    │  GET /api/models  — returns model list from .env
    │  POST /api/chat   — streams LLM response with tool calls
    │  GET /api/files   — lists Office files in a directory
    │
    │  OpenAI API  ──→  gpt-4o / gpt-4o-mini / gpt-4-turbo / gpt-3.5-turbo
    │              ──→  any OpenAI-compatible endpoint (Azure, local, proxy)
    │  Tool calls
    ▼
Office Tool Executor
    ├── mammoth                    (.docx)
    ├── adm-zip + fast-xml-parser  (.pptx)
    ├── SheetJS                    (.xlsx)
    ├── pdf-parse                  (.pdf native)
    ├── tesseract.js               (.pdf scanned / OCR)
    └── pdfplumber                 (.pdf tables via Python)
```

---

## Features

### LLM

- **OpenAI API** — uses the standard OpenAI SDK; works with any OpenAI-compatible endpoint
- **Configurable base URL** — set `OPENAI_BASE_URL` to point at Azure OpenAI, a local server (Ollama, LM Studio), or any other compatible proxy
- **Model selector** — switch between models at any time from the top-bar dropdown
- **Configurable model list** — set `MODELS` in `.env` to expose any models your endpoint supports
- **Streaming responses** via Server-Sent Events (SSE)
- **Agentic tool use** — the LLM automatically calls the right Office/PDF tool based on your question and iterates until it has a final answer
- **Tool call inspector** — expandable cards show each tool call name, arguments, and raw result
- **Context shift** — toggle in the topbar; when enabled, old messages are dropped (sliding window) so the conversation always fits within the context size limit
- **Context size** — number input in the topbar (active when context shift is on); sets the maximum number of messages retained; the system message is always preserved

### File browser

- **Directory scanner** — enter any path and click Scan to list Office files (`.docx`, `.pptx`, `.xlsx`, `.pdf`)
- **Recursive scan** — optional checkbox to include sub-directories
- **Multi-file selection** — click individual files to toggle them on/off; selected files are highlighted
- **Folder selection** — "Select all / Deselect all" button selects every file returned by the current scan
- **Sorting** — sort by **Type**, **Date** (last modified), or **Size**; click the active button again to reverse direction
- **Selected-files panel** — chosen files appear as removable chips above the chat input

### Chat

- **Manual query initiation** — selecting files never triggers the LLM automatically
- **File context injection** — on Send, selected file paths are prepended to the message
- **Adaptive placeholder** — textarea prompt updates to reflect whether files are selected
- **Stop generation** — cancel a streaming response mid-flight
- **Clear chat** — reset the conversation at any time
- **Auto OCR** — scanned PDFs are automatically processed with Tesseract when native text extraction yields no content
- **Dark / light theme** — follows the system colour-scheme preference

---

## Supported Models

The default list is `gpt-4o`, `gpt-4o-mini`, `gpt-4-turbo`, `gpt-3.5-turbo`. Override it with `MODELS` in `.env` to use any model your endpoint supports (see [Configuration](#configuration)).

| Model ID | Notes |
|----------|-------|
| `gpt-4o` | Default — best capability and speed |
| `gpt-4o-mini` | Lightweight, cost-effective |
| `gpt-4-turbo` | High capability, large context |
| `gpt-3.5-turbo` | Fastest, most cost-effective |

---

## Prerequisites

- **Node.js** 18 or higher — [nodejs.org](https://nodejs.org)
- **npm** 8 or higher (bundled with Node.js)
- **OpenAI API key** — [platform.openai.com](https://platform.openai.com)
- **Python 3** + **pdfplumber** *(optional)* — only required for PDF table extraction

Check your versions:
```bash
node --version   # must be >= 18
npm --version
python3 --version
```

---

## Quick Start

All commands below assume you are inside the `mcp-office-ui/` directory.

### 1. Configure environment variables

```bash
cp backend/.env.example backend/.env
```

Open `backend/.env` and configure it for your chosen provider:

**OpenAI (default)**
```dotenv
OPENAI_API_KEY=sk-...
OPENAI_BASE_URL=https://api.openai.com/v1
MODELS=gpt-4o,gpt-4o-mini,gpt-4-turbo,gpt-3.5-turbo
PORT=3001
```

**DeepSeek**
```dotenv
OPENAI_API_KEY=your_deepseek_api_key
OPENAI_BASE_URL=https://api.deepseek.com/v1
MODELS=deepseek-chat,deepseek-reasoner
PORT=3001
```

> **Important:** `MODELS` must match the model IDs accepted by your endpoint.
> If you change the endpoint but leave `MODELS` set to OpenAI model IDs (or vice versa),
> the API will return a "model not found" error.

### 2. Install dependencies

```bash
# Backend
cd backend
npm install

# Frontend
cd ../frontend
npm install
```

### 3. Start the backend

Open a terminal in `mcp-office-ui/backend/` and run:

```bash
# Development — auto-restarts on file changes
npm run dev

# Production
npm run build
npm start
```

The backend listens on **http://localhost:3001** by default.

### 4. Start the frontend

Open a second terminal in `mcp-office-ui/frontend/` and run:

```bash
npm run dev
```

Open **http://localhost:5173** in your browser.

> **Note:** The Vite dev server proxies all `/api/*` requests to `localhost:3001`. Both servers must be running at the same time.

### Optional: PDF table extraction

```bash
pip install pdfplumber
# or on some systems:
pip3 install pdfplumber
```

---

## Configuration

All configuration lives in `backend/.env`. Copy `backend/.env.example` as a starting point.

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `OPENAI_API_KEY` | **Yes** | — | API key from platform.openai.com |
| `OPENAI_BASE_URL` | No | `https://api.openai.com/v1` | Override for Azure, local, or proxy endpoints |
| `PORT` | No | `3001` | Port for the Express backend |
| `MODELS` | No | built-in list | Comma-separated model IDs for the UI dropdown |
| `CONTEXT_SHIFT` | No | `false` | Server-side default for the context shift toggle |
| `CONTEXT_SIZE` | No | `20` | Server-side default for the max messages to retain |

### Using an alternative or self-hosted endpoint

Set `OPENAI_BASE_URL` and `MODELS` to match your provider:

| Provider | `OPENAI_BASE_URL` | Example `MODELS` |
|----------|-------------------|-----------------|
| OpenAI (default) | `https://api.openai.com/v1` | `gpt-4o,gpt-4o-mini` |
| DeepSeek | `https://api.deepseek.com/v1` | `deepseek-chat,deepseek-reasoner` |
| Azure OpenAI | `https://<resource>.openai.azure.com/openai/deployments/<deployment>` | deployment name |
| Ollama (local) | `http://localhost:11434/v1` | `llama3,mistral` |
| LM Studio (local) | `http://localhost:1234/v1` | model name shown in LM Studio |

> Always restart the backend after editing `.env`.

---

## How to Use

1. **Scan a directory** — enter a folder path in the sidebar and click **Scan** (check *Recursive* to include sub-folders)
2. **Sort the results** — use the **Type / Date / Size** sort buttons; click the active button again to reverse direction
3. **Select files** — click individual file rows to toggle selection, or use **Select all**
4. **Ask a question** — type a question in the chat input, or leave it blank to get a default summarise prompt
5. **Send** — press **Send ↵** or hit **Enter**; the LLM receives the selected file paths as context and reads them via tool calls
6. **Review tool calls** — expand the tool-call cards to see which files were read and what was extracted
7. **Deselect / clear** — remove individual files with × on their chip, or click **Clear all**; switch models anytime from the top-bar dropdown

---

## API Reference

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/api/models` | Returns the configured model list |
| `POST` | `/api/chat` | Streams a chat completion with agentic tool use |
| `GET` | `/api/files?directory=&recursive=` | Lists Office files in a directory |
| `GET` | `/api/health` | Health check |

### `POST /api/chat`

```json
{
  "messages": [
    { "role": "user", "content": "Selected files:\n- /path/to/report.docx\n\nSummarise this document." }
  ],
  "model": "gpt-4o"
}
```

If `model` is omitted, it defaults to `gpt-4o`.

Response is a stream of `text/event-stream` events:

| Event type | Payload | Description |
|------------|---------|-------------|
| `text` | `{ content: string }` | Streamed text chunk |
| `tool_call` | `{ tool: { id, name, args } }` | LLM is calling a tool |
| `tool_result` | `{ tool: { id, name, args }, result }` | Tool execution result |
| `done` | — | Stream complete |
| `error` | `{ error: string }` | Error during generation |

### `GET /api/models`

```json
{ "models": ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo"] }
```

### `GET /api/files`

| Query param | Type | Description |
|-------------|------|-------------|
| `directory` | string (required) | Absolute path to scan |
| `recursive` | `true` / `false` | Include sub-directories (default `false`) |
| `file_types` | comma-separated | Filter by extension: `docx,pptx,xlsx,pdf` |

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| LLM | OpenAI API (any OpenAI-compatible endpoint) |
| Backend | Express + TypeScript |
| Frontend | React 18 + Vite + TypeScript |
| Streaming | Server-Sent Events (SSE) |
| Markdown | react-markdown + remark-gfm |
| Word | mammoth |
| PowerPoint | adm-zip + fast-xml-parser |
| Excel | SheetJS (xlsx) |
| PDF | pdf-parse + tesseract.js + pdfplumber |
