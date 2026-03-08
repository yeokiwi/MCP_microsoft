# MCP Office UI — Multi-LLM Chat Interface

A full-stack web application that integrates the **MCP Office Reader** tools with a configurable LLM (DeepSeek, Mistral, or GLM), providing a chat interface for reading and analysing Office files and PDFs.

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
    │  OpenAI-compatible API  ──→  DeepSeek LLM
    │                         ──→  Mistral LLM
    │                         ──→  GLM (Zhipu AI) LLM
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

- **Multi-provider support** — DeepSeek, Mistral, and GLM (Zhipu AI) via OpenAI-compatible APIs
- **Model selector** — switch between models at any time from the top-bar dropdown; optgroups are built dynamically from the model list
- **Configurable model list** — set `MODELS` in `.env` to expose exactly the models you want; no frontend code changes needed
- **Streaming responses** via Server-Sent Events (SSE)
- **Agentic tool use** — the LLM automatically calls the right Office/PDF tool based on your question and iterates until it has a final answer
- **Tool call inspector** — expandable cards show each tool call name, arguments, and raw result

### File browser

- **Directory scanner** — enter any path and click Scan to list Office files (`.docx`, `.pptx`, `.xlsx`, `.pdf`)
- **Recursive scan** — optional checkbox to include sub-directories
- **Multi-file selection** — click individual files to toggle them on/off (☐/☑); selected files are highlighted
- **Folder selection** — "Select all / Deselect all" button selects every file returned by the current scan in one click
- **Sorting** — sort the file list by **Type** (grouped by extension), **Date** (last modified), or **Size**; click the active sort button to toggle ascending ↑ / descending ↓
- **Selected-files panel** — chosen files appear as removable chips above the chat input; individual × buttons or "Clear all" to deselect

### Chat

- **Manual query initiation** — selecting files never triggers the LLM automatically; the user types a question (or leaves the textarea empty for a default summarise prompt) and presses **Send**
- **File context injection** — on Send, the selected file paths are prepended to the outgoing message so the LLM knows which files to read
- **Adaptive placeholder** — the textarea prompt updates to reflect whether files are currently selected
- **Stop generation** — cancel a streaming response mid-flight
- **Clear chat** — reset the conversation at any time
- **Auto OCR** — scanned PDFs are automatically processed with Tesseract when native text extraction yields no content
- **Dark / light theme** — follows the system colour-scheme preference

---

## Supported Models

The models listed here are the built-in defaults. You can replace or extend this list via the `MODELS` environment variable (see [Configuration](#configuration)).

| Provider | Model ID | Notes |
|----------|----------|-------|
| DeepSeek | `deepseek-chat` | Default model |
| DeepSeek | `deepseek-reasoner` | DeepSeek-R1 — strong reasoning |
| Mistral | `mistral-small-latest` | Fast, cost-effective |
| Mistral | `mistral-large-latest` | Most capable Mistral model |
| Mistral | `open-mistral-7b` | Open-weight, lightweight |
| Mistral | `open-mixtral-8x7b` | Open-weight MoE model |
| GLM | `glm-4` | Zhipu AI flagship model |
| GLM | `glm-4-flash` | Fast, low-latency variant |
| GLM | `glm-4-air` | Lightweight, cost-effective |

Provider routing is automatic based on the model ID prefix:

| Prefix | Provider | API key required |
|--------|----------|-----------------|
| `deepseek-*` | DeepSeek | `DEEPSEEK_API_KEY` |
| `mistral-*`, `open-mistral-*`, `open-mixtral-*` | Mistral | `MISTRAL_API_KEY` |
| `glm-*` | GLM / Zhipu AI | `GLM_API_KEY` |

---

## Prerequisites

- **Node.js** 18 or higher — [nodejs.org](https://nodejs.org)
- **npm** 8 or higher (bundled with Node.js)
- **API key** for at least one provider:
  - DeepSeek — [platform.deepseek.com](https://platform.deepseek.com)
  - Mistral — [console.mistral.ai](https://console.mistral.ai)
  - GLM (Zhipu AI) — [open.bigmodel.cn](https://open.bigmodel.cn)
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

Open `backend/.env` and fill in the values:

```dotenv
# API keys — only set the key(s) for the provider(s) you intend to use
DEEPSEEK_API_KEY=your_deepseek_api_key_here
MISTRAL_API_KEY=your_mistral_api_key_here
GLM_API_KEY=your_glm_api_key_here

PORT=3001

# Comma-separated list of model IDs to show in the UI dropdown.
# Provider is inferred from the prefix — see the routing table above.
# Omit this line to use the built-in default list.
MODELS=deepseek-chat,deepseek-reasoner,mistral-small-latest,mistral-large-latest,glm-4
```

> The backend throws a clear error if you select a model whose API key is missing.

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

| Variable | Required | Description |
|----------|----------|-------------|
| `DEEPSEEK_API_KEY` | If using DeepSeek models | API key from platform.deepseek.com |
| `MISTRAL_API_KEY` | If using Mistral models | API key from console.mistral.ai |
| `GLM_API_KEY` | If using GLM models | API key from open.bigmodel.cn |
| `PORT` | No (default `3001`) | Port for the Express backend |
| `MODELS` | No | Comma-separated model IDs for the UI dropdown |

### Adding a new model

1. Add the model ID to `MODELS` in `.env` (e.g., `MODELS=deepseek-chat,my-new-model`)
2. If it is from a new provider, add a new `if` branch in `backend/src/deepseek.ts` → `createLLMClient()` with the provider's base URL and API key variable
3. Restart the backend — the frontend picks up the new model automatically on next load

---

## How to Use

1. **Scan a directory** — enter a folder path in the sidebar and click **Scan** (check *Recursive* to include sub-folders)
2. **Sort the results** — use the **Type / Date / Size** sort buttons; click the active button again to reverse direction
3. **Select files** — click individual file rows to toggle selection, or use **Select all** to select the entire scanned folder
4. **Ask a question** — type a question in the chat input (e.g. *"Summarise the key points"*), or leave it blank to get a default summarise prompt
5. **Send** — press **Send ↵** or hit **Enter**; the LLM receives the selected file paths as context and reads them automatically via tool calls
6. **Review tool calls** — expand the tool-call cards that appear in the assistant message to see exactly which files were read and what was extracted
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
  "model": "mistral-large-latest"
}
```

If `model` is omitted, it defaults to `deepseek-chat`.

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
{ "models": ["deepseek-chat", "mistral-large-latest", "glm-4"] }
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
| LLM | DeepSeek / Mistral / GLM via OpenAI-compatible APIs |
| Backend | Express + TypeScript |
| Frontend | React 18 + Vite + TypeScript |
| Streaming | Server-Sent Events (SSE) |
| Markdown | react-markdown + remark-gfm |
| Word | mammoth |
| PowerPoint | adm-zip + fast-xml-parser |
| Excel | SheetJS (xlsx) |
| PDF | pdf-parse + tesseract.js + pdfplumber |
