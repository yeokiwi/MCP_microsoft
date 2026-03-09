import OpenAI from "openai";
import type { ChatCompletionMessageParam, ChatCompletionTool } from "openai/resources/chat/completions.js";
import {
  readWordDocument,
  readPowerPoint,
  readExcel,
  readPdf,
  listOfficeFiles,
} from "./tools/executor.js";

function createLLMClient(): OpenAI {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error("OPENAI_API_KEY environment variable is not set");
  return new OpenAI({
    apiKey,
    baseURL: process.env.OPENAI_BASE_URL ?? "https://api.openai.com/v1",
  });
}

export const SYSTEM_PROMPT = `You are an intelligent document assistant with access to tools for reading Office documents and PDFs.

## Handling Long Documents

To avoid exceeding the context window, always use selective extraction when a document may be large. Follow these strategies:

### Word Documents (.docx)
1. First call \`read_word_document\` with \`outline_only: true\` to get the list of sections (headings) and total word count.
2. Based on the outline and the user's question, call again with \`sections: [n, m, ...]\` (1-indexed) to fetch only the relevant sections.
3. Repeat for additional sections as needed.

### PDF Files (.pdf)
1. First call \`read_pdf\` with \`pages: [1]\` and \`include_metadata: true\` to see the total page count and a sample of content.
2. Then call with \`pages: [n, m, ...]\` (1-indexed) to fetch only the pages relevant to the user's question.
3. Work through pages incrementally for very large PDFs.

### PowerPoint Files (.pptx)
1. First call \`read_powerpoint\` with \`slides: [1]\` to see the total slide count and first slide content.
2. Then call with \`slides: [n, m, ...]\` (1-indexed) to fetch only the relevant slides.

### Excel Files (.xlsx)
1. First call \`read_excel\` with \`max_rows: 5\` to preview column headers and data structure.
2. Use \`sheets\` to read only the relevant sheets.
3. Use \`start_row\` and \`end_row\` (1-indexed data rows, after header) to page through large datasets in chunks.
4. Use \`columns\` (array of column names) to select only the columns relevant to the user's question.

Always prefer selective extraction over reading entire documents when the document may be large.`;

export const TOOL_DEFINITIONS: ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "read_word_document",
      description: "Extracts text content, headings, tables, and metadata from a .docx file. Supports selective section extraction to handle long documents that exceed context limits.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .docx file" },
          output_format: { type: "string", enum: ["markdown", "plain_text", "json"], default: "markdown" },
          include_metadata: { type: "boolean", default: true },
          outline_only: { type: "boolean", default: false, description: "If true, return only the document outline (numbered section headings) without full content. Use this first on long documents to discover structure." },
          sections: { type: "array", items: { type: "number" }, description: "Optional: specific 1-indexed section numbers to extract. Sections are delimited by headings. Call with outline_only: true first to see available sections." },
        },
        required: ["file_path"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_powerpoint",
      description: "Extracts slide content, speaker notes, and structure from a .pptx file. Use slides parameter to fetch only specific slides from long presentations.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .pptx file" },
          slides: { type: "array", items: { type: "number" }, description: "Optional: specific 1-indexed slide numbers to extract. Omit for all slides. Fetch slide 1 first to see total slide count in metadata." },
          include_notes: { type: "boolean", default: true },
          include_shapes: { type: "boolean", default: false },
        },
        required: ["file_path"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_excel",
      description: "Extracts cell data, sheet names, and formulas from a .xlsx file. Use sheets, start_row/end_row, and columns to selectively read large workbooks.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .xlsx file" },
          sheets: { type: "array", items: { type: "string" }, description: "Optional: specific sheet names to read. Omit for all sheets." },
          max_rows: { type: "number", default: 1000, description: "Maximum number of data rows to read per sheet. Use with start_row/end_row for pagination." },
          start_row: { type: "number", description: "Optional: 1-indexed starting data row (after the header). Use with end_row to read a specific range." },
          end_row: { type: "number", description: "Optional: 1-indexed ending data row (after the header). Use with start_row to read a specific range." },
          columns: { type: "array", items: { type: "string" }, description: "Optional: column names to include (requires header_row: true). Preview with max_rows: 5 first to see column names." },
          include_formulas: { type: "boolean", default: false },
          header_row: { type: "boolean", default: true },
        },
        required: ["file_path"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_pdf",
      description: "Extracts text, tables, and metadata from a .pdf file. Handles native and scanned PDFs with OCR fallback. Use pages parameter to fetch only specific pages from long PDFs.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .pdf file" },
          pages: { type: "array", items: { type: "number" }, description: "Optional: specific 1-indexed page numbers to extract. Omit for all pages. Fetch page 1 first to see total page count in metadata." },
          output_format: { type: "string", enum: ["markdown", "plain_text", "json"], default: "markdown" },
          extract_tables: { type: "boolean", default: false },
          ocr_fallback: { type: "boolean", default: true },
          include_metadata: { type: "boolean", default: true },
        },
        required: ["file_path"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_office_files",
      description: "Lists all .docx, .pptx, .xlsx, and .pdf files in a given directory.",
      parameters: {
        type: "object",
        properties: {
          directory: { type: "string", description: "Directory path to scan" },
          recursive: { type: "boolean", default: false },
          file_types: { type: "array", items: { type: "string", enum: ["docx", "pptx", "xlsx", "pdf"] } },
        },
        required: ["directory"],
      },
    },
  },
];

export async function executeTool(name: string, args: Record<string, unknown>): Promise<string> {
  try {
    let result: unknown;
    switch (name) {
      case "read_word_document":
        result = await readWordDocument(args as Parameters<typeof readWordDocument>[0]);
        break;
      case "read_powerpoint":
        result = await readPowerPoint(args as Parameters<typeof readPowerPoint>[0]);
        break;
      case "read_excel":
        result = await readExcel(args as Parameters<typeof readExcel>[0]);
        break;
      case "read_pdf":
        result = await readPdf(args as Parameters<typeof readPdf>[0]);
        break;
      case "list_office_files":
        result = await listOfficeFiles(args as Parameters<typeof listOfficeFiles>[0]);
        break;
      default:
        return JSON.stringify({ error: `Unknown tool: ${name}` });
    }
    return truncateToolResult(JSON.stringify(result, null, 2));
  } catch (e) {
    return JSON.stringify({ error: (e as Error).message });
  }
}

// ── Context-length guards ──────────────────────────────────────────────────────

/**
 * Max characters for a single tool result (~30K tokens at 4 chars/token).
 * Leaves ample room in the 131K-token context window for conversation + completion.
 */
const TOOL_RESULT_MAX_CHARS = 120_000;

/**
 * Max total characters across all conversation messages before we start pruning
 * old tool results (~90K tokens, giving ~40K tokens headroom for completion).
 */
const CONVERSATION_MAX_CHARS = 360_000;

/** Truncate a tool result that is too large, appending a hint to use selective params. */
function truncateToolResult(result: string): string {
  if (result.length <= TOOL_RESULT_MAX_CHARS) return result;
  const truncated = result.slice(0, TOOL_RESULT_MAX_CHARS);
  // Try to end at a clean JSON boundary
  const lastNewline = truncated.lastIndexOf("\n");
  const safe = lastNewline > TOOL_RESULT_MAX_CHARS * 0.9 ? truncated.slice(0, lastNewline) : truncated;
  return (
    safe +
    "\n\n[RESULT TRUNCATED — the output exceeded the context limit. " +
    "Use selective parameters to reduce the payload: " +
    "outline_only/sections (Word), pages (PDF), slides (PowerPoint), " +
    "sheets/start_row/end_row/columns/max_rows (Excel).]"
  );
}

/** Rough character-count estimate of all message content. */
function estimateSize(messages: ChatCompletionMessageParam[]): number {
  return messages.reduce((total, msg) => {
    if (typeof msg.content === "string") return total + msg.content.length;
    if (Array.isArray(msg.content)) {
      return total + (msg.content as { text?: string }[]).reduce((s, c) => s + (c.text?.length ?? 0), 0);
    }
    return total;
  }, 0);
}

/**
 * When the accumulated conversation exceeds CONVERSATION_MAX_CHARS,
 * replace the content of the oldest tool-result messages with a short stub,
 * keeping at least the 4 most recent messages intact.
 */
function pruneConversation(messages: ChatCompletionMessageParam[]): ChatCompletionMessageParam[] {
  if (estimateSize(messages) <= CONVERSATION_MAX_CHARS) return messages;

  const result = messages.map((m) => ({ ...m }));
  // Leave system message (index 0) and the last 4 messages alone
  for (let i = 1; i < result.length - 4; i++) {
    if (estimateSize(result) <= CONVERSATION_MAX_CHARS) break;
    const msg = result[i];
    if (msg.role === "tool" && typeof msg.content === "string" && msg.content.length > 200) {
      result[i] = { ...msg, content: "[Tool result omitted from context to stay within token limit]" };
    }
  }
  return result;
}

export interface StreamEvent {
  type: "text" | "tool_call" | "tool_result" | "done" | "error";
  content?: string;
  tool?: { id: string; name: string; args: string };
  result?: string;
  error?: string;
}

export async function chatWithLLM(
  messages: ChatCompletionMessageParam[],
  onEvent: (event: StreamEvent) => void,
  model = "gpt-4o"
): Promise<void> {
  const client = createLLMClient();
  const systemMessage: ChatCompletionMessageParam = { role: "system", content: SYSTEM_PROMPT };
  const conversationMessages: ChatCompletionMessageParam[] = [systemMessage, ...messages];

  // Agentic loop: keep running until model produces a final text response
  while (true) {
    const stream = await client.chat.completions.create({
      model,
      messages: pruneConversation(conversationMessages),
      tools: TOOL_DEFINITIONS,
      tool_choice: "auto",
      stream: true,
    });

    let currentText = "";
    const toolCalls: { id: string; name: string; arguments: string }[] = [];

    for await (const chunk of stream) {
      const delta = chunk.choices[0]?.delta;
      if (!delta) continue;

      if (delta.content) {
        currentText += delta.content;
        onEvent({ type: "text", content: delta.content });
      }

      if (delta.tool_calls) {
        for (const tc of delta.tool_calls) {
          const idx = tc.index ?? 0;
          if (!toolCalls[idx]) {
            toolCalls[idx] = { id: tc.id ?? "", name: tc.function?.name ?? "", arguments: "" };
          }
          if (tc.function?.arguments) toolCalls[idx].arguments += tc.function.arguments;
          if (tc.id) toolCalls[idx].id = tc.id;
          if (tc.function?.name) toolCalls[idx].name = tc.function.name;
        }
      }

      if (chunk.choices[0]?.finish_reason === "stop") {
        onEvent({ type: "done" });
        return;
      }

      if (chunk.choices[0]?.finish_reason === "tool_calls") {
        break;
      }
    }

    if (toolCalls.length === 0) {
      onEvent({ type: "done" });
      return;
    }

    conversationMessages.push({
      role: "assistant",
      content: currentText || null,
      tool_calls: toolCalls.map((tc) => ({
        id: tc.id,
        type: "function" as const,
        function: { name: tc.name, arguments: tc.arguments },
      })),
    });

    for (const tc of toolCalls) {
      onEvent({ type: "tool_call", tool: { id: tc.id, name: tc.name, args: tc.arguments } });

      let parsedArgs: Record<string, unknown> = {};
      try { parsedArgs = JSON.parse(tc.arguments); } catch { /* ignore */ }

      const result = await executeTool(tc.name, parsedArgs);
      onEvent({ type: "tool_result", tool: { id: tc.id, name: tc.name, args: tc.arguments }, result });

      conversationMessages.push({
        role: "tool",
        tool_call_id: tc.id,
        content: result,
      });
    }
  }
}
