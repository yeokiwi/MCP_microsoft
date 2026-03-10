import { useEffect, useRef, useCallback, useState } from "react";
import { useChat } from "./hooks/useChat.js";
import { MessageBubble } from "./components/MessageBubble.js";
import { Sidebar } from "./components/Sidebar.js";

interface ModelOption {
  value: string;
  label: string;
}

export default function App() {
  const [models, setModels] = useState<ModelOption[]>([]);
  const [selectedModel, setSelectedModel] = useState("");
  const { messages, isLoading, sendMessage, stopGeneration, clearChat } = useChat(selectedModel);
  const [input, setInput] = useState("");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [selectedPaths, setSelectedPaths] = useState<Set<string>>(new Set());
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Fetch available models from backend on mount
  useEffect(() => {
    fetch("/api/models")
      .then((r) => r.json())
      .then((data: { models: string[] }) => {
        const opts: ModelOption[] = data.models.map((id) => ({ value: id, label: id }));
        setModels(opts);
        if (opts.length > 0) setSelectedModel(opts[0].value);
      })
      .catch(() => {
        setModels([{ value: "gpt-4o", label: "gpt-4o" }]);
        setSelectedModel("gpt-4o");
      });
  }, []);

  // Auto-scroll on new messages
  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const handleToggleFile = useCallback((path: string) => {
    setSelectedPaths((prev) => {
      const next = new Set(prev);
      if (next.has(path)) next.delete(path);
      else next.add(path);
      return next;
    });
  }, []);

  const handleSelectAll = useCallback((paths: string[]) => {
    if (paths.length === 0) {
      setSelectedPaths(new Set());
    } else {
      setSelectedPaths((prev) => {
        const next = new Set(prev);
        paths.forEach((p) => next.add(p));
        return next;
      });
    }
  }, []);

  const handleRemovePath = useCallback((path: string) => {
    setSelectedPaths((prev) => {
      const next = new Set(prev);
      next.delete(path);
      return next;
    });
  }, []);

  const handleSubmit = useCallback(() => {
    const text = input.trim();
    if ((!text && selectedPaths.size === 0) || isLoading) return;

    let message = text;
    if (selectedPaths.size > 0) {
      const fileList = [...selectedPaths].map((p) => `- ${p}`).join("\n");
      const fileContext = `Selected files:\n${fileList}`;
      message = text
        ? `${fileContext}\n\n${text}`
        : `${fileContext}\n\nPlease read and summarise the selected file${selectedPaths.size > 1 ? "s" : ""}.`;
    }

    setInput("");
    sendMessage(message);
    if (textareaRef.current) textareaRef.current.style.height = "auto";
  }, [input, isLoading, sendMessage, selectedPaths]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  const handleInput = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    const ta = e.target;
    ta.style.height = "auto";
    ta.style.height = `${Math.min(ta.scrollHeight, 200)}px`;
  };

  const SUGGESTIONS = [
    "List all Office files in /tmp",
    "Read and summarise a Word document at /path/to/file.docx",
    "Extract all slides from /path/to/presentation.pptx",
    "Show me the data from /path/to/spreadsheet.xlsx",
    "Extract text from /path/to/document.pdf",
  ];

  const canSend = (input.trim().length > 0 || selectedPaths.size > 0) && !isLoading;

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
            <span className="app-subtitle">powered by OpenAI</span>
          </h1>
        </div>
        <div className="topbar-right">
          {models.length > 0 && (
            <select
              className="model-selector"
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value)}
              disabled={isLoading}
              title="Select model"
            >
              {models.map((m) => (
                <option key={m.value} value={m.value}>{m.label}</option>
              ))}
            </select>
          )}
          <button className="btn btn-ghost" onClick={clearChat} title="Clear chat">
            🗑 Clear
          </button>
          <button className="btn btn-ghost" onClick={() => window.print()} title="Print conversation to PDF">
            🖨 Print
          </button>
        </div>
      </header>

      <div className="app-body">
        {/* Sidebar */}
        {sidebarOpen && (
          <Sidebar
            selectedPaths={selectedPaths}
            onToggleFile={handleToggleFile}
            onSelectAll={handleSelectAll}
          />
        )}

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
                  <button key={s} className="suggestion-chip" onClick={() => sendMessage(s)}>
                    {s}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Input area */}
          <div className="input-area">
            {/* Selected files chips */}
            {selectedPaths.size > 0 && (
              <div className="selected-files">
                <span className="selected-files-label">
                  {selectedPaths.size} file{selectedPaths.size > 1 ? "s" : ""} selected:
                </span>
                <div className="selected-files-chips">
                  {[...selectedPaths].map((path) => {
                    const name = path.split("/").pop() ?? path;
                    return (
                      <span key={path} className="file-chip" title={path}>
                        {name}
                        <button
                          className="file-chip-remove"
                          onClick={() => handleRemovePath(path)}
                          title={`Remove ${path}`}
                        >
                          ×
                        </button>
                      </span>
                    );
                  })}
                  <button className="btn-link" onClick={() => setSelectedPaths(new Set())}>
                    Clear all
                  </button>
                </div>
              </div>
            )}

            <div className="input-wrapper">
              <textarea
                ref={textareaRef}
                className="chat-input"
                value={input}
                onChange={handleInput}
                onKeyDown={handleKeyDown}
                placeholder={
                  selectedPaths.size > 0
                    ? "Ask something about the selected files, or just press Send to summarise…"
                    : "Ask about an Office file or PDF… (Enter to send, Shift+Enter for newline)"
                }
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
                    disabled={!canSend}
                    title="Send message"
                  >
                    Send ↵
                  </button>
                )}
              </div>
            </div>
            <p className="input-hint">
              Select files from the sidebar, then type your question and press Send.
            </p>
          </div>
        </main>
      </div>
    </div>
  );
}
