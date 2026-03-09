import { useState, useCallback, useRef } from "react";
import type { Message, StreamEvent, ToolCall } from "../types.js";

const API_BASE = "/api";

function uid(): string {
  return Math.random().toString(36).slice(2, 10);
}

export function useChat(model: string) {
  const [messages, setMessages] = useState<Message[]>([
    {
      id: uid(),
      role: "assistant",
      content: "Hello! I can read and analyse **Word documents**, **PowerPoint presentations**, **Excel workbooks**, and **PDF files** for you. Just tell me which file you'd like to explore, or ask me to list files in a directory.",
      timestamp: Date.now(),
    },
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const abortRef = useRef<AbortController | null>(null);

  const sendMessage = useCallback(async (userText: string) => {
    if (!userText.trim() || isLoading) return;

    const userMsg: Message = { id: uid(), role: "user", content: userText, timestamp: Date.now() };
    setMessages((prev) => [...prev, userMsg]);

    const assistantId = uid();
    const assistantMsg: Message = { id: assistantId, role: "assistant", content: "", toolCalls: [], isStreaming: true, timestamp: Date.now() };
    setMessages((prev) => [...prev, assistantMsg]);
    setIsLoading(true);

    const ctrl = new AbortController();
    abortRef.current = ctrl;

    // Build conversation for API (skip initial assistant greeting from history if desired)
    const historyMessages = [...messages, userMsg].map((m) => ({
      role: m.role,
      content: m.content || (m.toolCalls?.map((tc) => `[Tool: ${tc.name}] ${tc.result ?? ""}`).join("\n") ?? ""),
    }));

    try {
      const res = await fetch(`${API_BASE}/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: historyMessages, model }),
        signal: ctrl.signal,
      });

      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (!res.body) throw new Error("No response body");

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() ?? "";

        for (const line of lines) {
          if (!line.startsWith("data: ")) continue;
          const raw = line.slice(6).trim();
          if (!raw) continue;

          let event: StreamEvent;
          try { event = JSON.parse(raw); } catch { continue; }

          setMessages((prev) =>
            prev.map((msg) => {
              if (msg.id !== assistantId) return msg;

              if (event.type === "text" && event.content) {
                return { ...msg, content: msg.content + event.content };
              }

              if (event.type === "tool_call" && event.tool) {
                const newTc: ToolCall = { id: event.tool.id, name: event.tool.name, args: event.tool.args, loading: true };
                return { ...msg, toolCalls: [...(msg.toolCalls ?? []), newTc] };
              }

              if (event.type === "tool_result" && event.tool) {
                return {
                  ...msg,
                  toolCalls: (msg.toolCalls ?? []).map((tc) =>
                    tc.id === event.tool!.id ? { ...tc, result: event.result, loading: false } : tc
                  ),
                };
              }

              if (event.type === "done") {
                return { ...msg, isStreaming: false };
              }

              if (event.type === "error") {
                return { ...msg, content: msg.content + `\n\n**Error:** ${event.error}`, isStreaming: false };
              }

              return msg;
            })
          );
        }
      }
    } catch (e) {
      if ((e as Error).name === "AbortError") return;
      setMessages((prev) =>
        prev.map((msg) =>
          msg.id === assistantId
            ? { ...msg, content: `Sorry, something went wrong: ${(e as Error).message}`, isStreaming: false }
            : msg
        )
      );
    } finally {
      setIsLoading(false);
      abortRef.current = null;
    }
  }, [messages, isLoading]);

  const stopGeneration = useCallback(() => {
    abortRef.current?.abort();
    setIsLoading(false);
    setMessages((prev) =>
      prev.map((msg) => msg.isStreaming ? { ...msg, isStreaming: false } : msg)
    );
  }, []);

  const clearChat = useCallback(() => {
    setMessages([{
      id: uid(),
      role: "assistant",
      content: "Chat cleared. How can I help you with your Office files?",
      timestamp: Date.now(),
    }]);
  }, []);

  return { messages, isLoading, sendMessage, stopGeneration, clearChat };
}
