import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import type { Message } from "../types.js";
import { ToolCallCard } from "./ToolCallCard.js";

interface Props {
  message: Message;
}

function formatTime(ts: number): string {
  return new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

export function MessageBubble({ message }: Props) {
  const isUser = message.role === "user";
  const hasTools = message.toolCalls && message.toolCalls.length > 0;

  return (
    <div className={`message-wrapper ${isUser ? "message-user" : "message-assistant"}`}>
      <div className="message-avatar">
        {isUser ? "👤" : "🤖"}
      </div>
      <div className="message-content">
        {hasTools && (
          <div className="tool-calls-container">
            {message.toolCalls!.map((tc) => (
              <ToolCallCard key={tc.id} toolCall={tc} />
            ))}
          </div>
        )}

        {message.content && (
          <div className={`message-bubble ${isUser ? "bubble-user" : "bubble-assistant"}`}>
            {isUser ? (
              <p className="message-text">{message.content}</p>
            ) : (
              <div className="markdown-body">
                <ReactMarkdown remarkPlugins={[remarkGfm]}>
                  {message.content}
                </ReactMarkdown>
                {message.isStreaming && <span className="cursor-blink">▍</span>}
              </div>
            )}
          </div>
        )}

        {!message.content && message.isStreaming && (
          <div className="bubble-assistant message-bubble">
            <span className="thinking-dots">
              <span>.</span><span>.</span><span>.</span>
            </span>
          </div>
        )}

        <span className="message-time">{formatTime(message.timestamp)}</span>
      </div>
    </div>
  );
}
