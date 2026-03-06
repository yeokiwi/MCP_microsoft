import { useEffect, useRef, useCallback, useState } from "react";
import { useChat } from "./hooks/useChat.js";
import { MessageBubble } from "./components/MessageBubble.js";
import { Sidebar } from "./components/Sidebar.js";

const MODELS = [
  { value: "deepseek-chat", label: "DeepSeek Chat", provider: "DeepSeek" },
  { value: "deepseek-reasoner", label: "DeepSeek Reasoner", provider: "DeepSeek" },
  { value: "mistral-small-latest", label: "Mistral Small", provider: "Mistral" },
  { value: "mistral-large-latest", label: "Mistral Large", provider: "Mistral" },
  { value: "open-mistral-7b", label: "Mistral 7B (open)", provider: "Mistral" },
  { value: "open-mixtral-8x7b", label: "Mixtral 8x7B (open)", provider: "Mistral" },
];

export default function App() {
  const [selectedModel, setSelectedModel] = useState(MODELS[0].value);
  const { messages, isLoading, sendMessage, stopGeneration, clearChat } = useChat(selectedModel);
  const [input, setInput] = useState("");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const currentModel = MODELS.find((m) => m.value === selectedModel) ?? MODELS[0];

  // Auto-scroll on new messages
  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const handleSubmit = useCallback(() => {
    const text = input.trim();
    if (!text || isLoading) return;
    setInput("");
    sendMessage(text);
    if (textareaRef.current) textareaRef.current.style.height = "auto";
  }, [input, isLoading, sendMessage]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  const handleInput = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    // Auto-resize
    const ta = e.target;
    ta.style.height = "auto";
    ta.style.height = `${Math.min(ta.scrollHeight, 200)}px`;
  };

  const handleFileSelect = useCallback((path: string) => {
    sendMessage(`Please read and summarise the file: ${path}`);
  }, [sendMessage]);

  const SUGGESTIONS = [
    "List all Office files in /tmp",
    "Read and summarise a Word document at /path/to/file.docx",
    "Extract all slides from /path/to/presentation.pptx",
    "Show me the data from /path/to/spreadsheet.xlsx",
    "Extract text from /path/to/document.pdf",
  ];

  return (
    <div className="app-root">
      {/* Top bar */}
      <header className="topbar">
        <div className="topbar-left">
          <button className="btn-icon" onClick={() => setSidebarOpen((v) => !v)} title="Toggle sidebar">
            ☰
          </button>
          <h1 className="app-title">
            <span className="app-logo">📄</span>
            MCP Office Reader
            <span className="app-subtitle">powered by {currentModel.provider}</span>
          </h1>
        </div>
        <div className="topbar-right">
          <select
            className="model-selector"
            value={selectedModel}
            onChange={(e) => setSelectedModel(e.target.value)}
            disabled={isLoading}
            title="Select LLM model"
          >
            <optgroup label="DeepSeek">
              {MODELS.filter((m) => m.provider === "DeepSeek").map((m) => (
                <option key={m.value} value={m.value}>{m.label}</option>
              ))}
            </optgroup>
            <optgroup label="Mistral">
              {MODELS.filter((m) => m.provider === "Mistral").map((m) => (
                <option key={m.value} value={m.value}>{m.label}</option>
              ))}
            </optgroup>
          </select>
          <button className="btn btn-ghost" onClick={clearChat} title="Clear chat">
            🗑 Clear
          </button>
        </div>
      </header>

      <div className="app-body">
        {/* Sidebar */}
        {sidebarOpen && <Sidebar onFileSelect={handleFileSelect} />}

        {/* Main chat area */}
        <main className="chat-main">
          <div className="messages-container">
            {messages.map((msg) => (
              <MessageBubble key={msg.id} message={msg} />
            ))}
            <div ref={bottomRef} />
          </div>

          {/* Suggestions (shown when only greeting is present) */}
          {messages.length <= 1 && (
            <div className="suggestions">
              <p className="suggestions-label">Try asking:</p>
              <div className="suggestions-grid">
                {SUGGESTIONS.map((s) => (
                  <button
                    key={s}
                    className="suggestion-chip"
                    onClick={() => sendMessage(s)}
                  >
                    {s}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Input area */}
          <div className="input-area">
            <div className="input-wrapper">
              <textarea
                ref={textareaRef}
                className="chat-input"
                value={input}
                onChange={handleInput}
                onKeyDown={handleKeyDown}
                placeholder="Ask about an Office file or PDF… (Enter to send, Shift+Enter for newline)"
                rows={1}
                disabled={isLoading}
              />
              <div className="input-actions">
                {isLoading ? (
                  <button className="btn btn-stop" onClick={stopGeneration} title="Stop generation">
                    ⏹ Stop
                  </button>
                ) : (
                  <button
                    className="btn btn-primary"
                    onClick={handleSubmit}
                    disabled={!input.trim()}
                    title="Send message"
                  >
                    Send ↵
                  </button>
                )}
              </div>
            </div>
            <p className="input-hint">
              {currentModel.label} will automatically read and analyse Office files when you share a file path.
            </p>
          </div>
        </main>
      </div>
    </div>
  );
}
