import { useState } from "react";
import type { ToolCall } from "../types.js";
import { TOOL_ICONS } from "../types.js";

function formatToolName(name: string): string {
  return name.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
}

function tryParseArgs(args: string): string {
  try {
    const parsed = JSON.parse(args);
    return JSON.stringify(parsed, null, 2);
  } catch {
    return args;
  }
}

function tryFormatResult(result: string): string {
  try {
    const parsed = JSON.parse(result);
    return JSON.stringify(parsed, null, 2);
  } catch {
    return result;
  }
}

interface Props {
  toolCall: ToolCall;
}

export function ToolCallCard({ toolCall }: Props) {
  const [expanded, setExpanded] = useState(false);
  const icon = TOOL_ICONS[toolCall.name] ?? "🔧";
  const isLoading = toolCall.loading;
  const hasResult = toolCall.result != null;

  return (
    <div className="tool-card">
      <button
        className="tool-card-header"
        onClick={() => setExpanded((v) => !v)}
        aria-expanded={expanded}
      >
        <span className="tool-icon">{icon}</span>
        <span className="tool-name">{formatToolName(toolCall.name)}</span>
        {isLoading && <span className="tool-status loading">Running…</span>}
        {!isLoading && hasResult && <span className="tool-status done">Done</span>}
        <span className="tool-chevron">{expanded ? "▲" : "▼"}</span>
      </button>

      {expanded && (
        <div className="tool-card-body">
          <div className="tool-section">
            <span className="tool-section-label">Arguments</span>
            <pre className="tool-code">{tryParseArgs(toolCall.args)}</pre>
          </div>
          {hasResult && (
            <div className="tool-section">
              <span className="tool-section-label">Result</span>
              <pre className="tool-code tool-result">{tryFormatResult(toolCall.result!)}</pre>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
