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
  { value: "glm-4", label: "GLM-4", provider: "GLM" },
  { value: "glm-4-flash", label: "GLM-4 Flash", provider: "GLM" },
  { value: "glm-4-air", label: "GLM-4 Air", provider: "GLM" },
];

export default function App() {
  const [selectedModel, setSelectedModel] = useState(MODELS[0].value);
  const { messages, isLoading, sendMessage, stopGeneration, clearChat } = useChat(selectedModel);
  const [input, setInput] = useState("");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [selectedPaths, setSelectedPaths] = useState<Set<string>>(new Set());
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const currentModel = MODELS.find((m) => m.value === selectedModel) ?? MODELS[0];

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

  // Called with a list of paths to add, or empty array to deselect all current scanned files
  const handleSelectAll = useCallback((paths: string[]) => {
    if (paths.length === 0) {
      // Sidebar signals "deselect all" — we clear only the paths it knows about,
      // but since we don't have them here, clear everything.
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
      message = text ? `${fileContext}\n\n${text}` : `${fileContext}\n\nPlease read and summarise the selected file${selectedPaths.size > 1 ? "s" : ""}.`;
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
            <optgroup label="GLM (Zhipu AI)">
              {MODELS.filter((m) => m.provider === "GLM").map((m) => (
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
