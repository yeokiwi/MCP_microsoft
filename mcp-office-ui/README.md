# MCP Office UI — Multi-LLM Chat Interface

A full-stack web application that integrates the **MCP Office Reader** tools with your choice of **DeepSeek**, **Mistral**, or **GLM** LLM, providing a chat interface for reading and analysing Office files and PDFs.

## Architecture

```
mcp-office-ui/
├── backend/         Express API server — LLM integration + tool execution
└── frontend/        React + Vite chat UI
```

```
Browser (React)
    │  SSE stream
    ▼
Express Backend (port 3001)
    │  OpenAI-compatible API  ──→  DeepSeek LLM
    │                         ──→  Mistral LLM
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

## Supported Models

| Provider | Model ID | Notes |
|----------|----------|-------|
| DeepSeek | `deepseek-chat` | Default |
| DeepSeek | `deepseek-reasoner` | DeepSeek-R1, strong reasoning |
| Mistral | `mistral-small-latest` | Fast, cost-effective |
| Mistral | `mistral-large-latest` | Most capable Mistral model |
| Mistral | `open-mistral-7b` | Open-weight, lightweight |
| Mistral | `open-mixtral-8x7b` | Open-weight MoE model |
| GLM | `glm-4` | Zhipu AI flagship model |
| GLM | `glm-4-flash` | Fast, low-latency variant |
| GLM | `glm-4-air` | Lightweight, cost-effective |

Switch between models at any time using the dropdown in the top bar.

## Prerequisites

- **Node.js** 18 or higher — [nodejs.org](https://nodejs.org)
- **npm** 8 or higher (bundled with Node.js)
- **API key** for at least one provider:
  - DeepSeek — [platform.deepseek.com](https://platform.deepseek.com)
  - Mistral — [console.mistral.ai](https://console.mistral.ai)
  - GLM (Zhipu AI) — [open.bigmodel.cn](https://open.bigmodel.cn)
- **Python 3** + **pdfplumber** — only required for PDF table extraction

Check your versions:
```bash
node --version   # must be >= 18
npm --version
python3 --version
```

## Quick Start

All commands below assume you are inside the `mcp-office-ui/` directory.

### 1. Configure API keys

```bash
cp backend/.env.example backend/.env
```

Open `backend/.env` and set the key(s) for the provider(s) you want to use:
```
DEEPSEEK_API_KEY=your_deepseek_api_key_here
MISTRAL_API_KEY=your_mistral_api_key_here
GLM_API_KEY=your_glm_api_key_here
PORT=3001
```

You only need to set the key for the provider you intend to use. The backend will throw a clear error if you select a model whose API key is missing.

### 2. Install dependencies

```bash
# Install backend dependencies
cd backend
npm install

# Install frontend dependencies
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

> **Note:** The frontend dev server proxies all `/api/*` requests to `localhost:3001`, so both servers must be running.

### Optional: PDF table extraction

```bash
pip install pdfplumber
# or, on some systems:
pip3 install pdfplumber
```

## Features

- **Multi-LLM support** — switch between DeepSeek and Mistral models from the top bar
- **Chat with any supported model** about any Office file or PDF
- **Agentic tool use** — the model automatically calls the right tool based on your question
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

The `/api/chat` request body accepts:
```json
{
  "messages": [...],
  "model": "mistral-large-latest"
}
```

If `model` is omitted, it defaults to `deepseek-chat`.

## Tech Stack

| Layer | Tech |
|-------|------|
| LLM | DeepSeek / Mistral / GLM (OpenAI-compatible APIs) |
| Backend | Express + TypeScript |
| Frontend | React 18 + Vite + TypeScript |
| Streaming | Server-Sent Events (SSE) |
| Markdown | react-markdown + remark-gfm |
