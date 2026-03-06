import OpenAI from "openai";
import type { ChatCompletionMessageParam, ChatCompletionTool } from "openai/resources/chat/completions.js";
import {
  readWordDocument,
  readPowerPoint,
  readExcel,
  readPdf,
  listOfficeFiles,
} from "./tools/executor.js";

export function createDeepSeekClient(): OpenAI {
  const apiKey = process.env.DEEPSEEK_API_KEY;
  if (!apiKey) throw new Error("DEEPSEEK_API_KEY environment variable is not set");
  return new OpenAI({ apiKey, baseURL: "https://api.deepseek.com" });
}

export const TOOL_DEFINITIONS: ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "read_word_document",
      description: "Extracts text content, headings, tables, and metadata from a .docx file.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .docx file" },
          output_format: { type: "string", enum: ["markdown", "plain_text", "json"], default: "markdown" },
          include_metadata: { type: "boolean", default: true },
        },
        required: ["file_path"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_powerpoint",
      description: "Extracts slide content, speaker notes, and structure from a .pptx file.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .pptx file" },
          slides: { type: "array", items: { type: "number" }, description: "Optional: specific 1-indexed slide numbers" },
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
      description: "Extracts cell data, sheet names, and formulas from a .xlsx file.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .xlsx file" },
          sheets: { type: "array", items: { type: "string" }, description: "Optional: specific sheet names" },
          max_rows: { type: "number", default: 1000 },
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
      description: "Extracts text, tables, and metadata from a .pdf file. Handles native and scanned PDFs with OCR fallback.",
      parameters: {
        type: "object",
        properties: {
          file_path: { type: "string", description: "Absolute or relative path to the .pdf file" },
          pages: { type: "array", items: { type: "number" }, description: "Optional: specific 1-indexed page numbers" },
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

// Execute a tool call by name
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
    return JSON.stringify(result, null, 2);
  } catch (e) {
    return JSON.stringify({ error: (e as Error).message });
  }
}

export interface StreamEvent {
  type: "text" | "tool_call" | "tool_result" | "done" | "error";
  content?: string;
  tool?: { id: string; name: string; args: string };
  result?: string;
  error?: string;
}

// Chat with DeepSeek using tool calling, streaming events via callback
export async function chatWithDeepSeek(
  messages: ChatCompletionMessageParam[],
  onEvent: (event: StreamEvent) => void,
  model = "deepseek-chat"
): Promise<void> {
  const client = createDeepSeekClient();
  const conversationMessages: ChatCompletionMessageParam[] = [...messages];

  // Agentic loop: keep running until model produces a final text response
  while (true) {
    const stream = await client.chat.completions.create({
      model,
      messages: conversationMessages,
      tools: TOOL_DEFINITIONS,
      tool_choice: "auto",
      stream: true,
    });

    let currentText = "";
    const toolCalls: { id: string; name: string; arguments: string }[] = [];

    for await (const chunk of stream) {
      const delta = chunk.choices[0]?.delta;
      if (!delta) continue;

      // Accumulate text
      if (delta.content) {
        currentText += delta.content;
        onEvent({ type: "text", content: delta.content });
      }

      // Accumulate tool calls
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

    // Execute tool calls
    if (toolCalls.length === 0) {
      onEvent({ type: "done" });
      return;
    }

    // Add assistant message with tool calls
    conversationMessages.push({
      role: "assistant",
      content: currentText || null,
      tool_calls: toolCalls.map((tc) => ({
        id: tc.id,
        type: "function" as const,
        function: { name: tc.name, arguments: tc.arguments },
      })),
    });

    // Execute each tool call and add results
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
    // Continue loop to get next model response
  }
}
