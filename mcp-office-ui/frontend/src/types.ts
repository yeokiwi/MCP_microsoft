export type MessageRole = "user" | "assistant" | "system";

export interface ToolCall {
  id: string;
  name: string;
  args: string;
  result?: string;
  loading?: boolean;
}

export interface Message {
  id: string;
  role: MessageRole;
  content: string;
  toolCalls?: ToolCall[];
  isStreaming?: boolean;
  timestamp: number;
}

export interface OfficeFile {
  path: string;
  name: string;
  extension: string;
  size_bytes: number;
  modified?: string;
}

export interface StreamEvent {
  type: "text" | "tool_call" | "tool_result" | "done" | "error";
  content?: string;
  tool?: { id: string; name: string; args: string };
  result?: string;
  error?: string;
}

export const TOOL_ICONS: Record<string, string> = {
  read_word_document: "📝",
  read_powerpoint: "📊",
  read_excel: "📈",
  read_pdf: "📄",
  list_office_files: "📁",
};

export const FILE_ICONS: Record<string, string> = {
  docx: "📝",
  pptx: "📊",
  xlsx: "📈",
  pdf: "📄",
};
