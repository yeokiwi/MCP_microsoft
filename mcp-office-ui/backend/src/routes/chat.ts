import { Router, Request, Response } from "express";
import type { ChatCompletionMessageParam } from "openai/resources/chat/completions.js";
import { chatWithLLM } from "../deepseek.js";

const router = Router();

router.post("/", async (req: Request, res: Response) => {
  const { messages, model } = req.body as {
    messages?: ChatCompletionMessageParam[];
    model?: string;
  };

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: "messages array is required" });
    return;
  }

  // SSE headers
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();

  const send = (data: object) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  try {
    await chatWithLLM(
      messages,
      (event) => send(event),
      model ?? "deepseek-chat"
    );
  } catch (e) {
    send({ type: "error", error: (e as Error).message });
  } finally {
    res.end();
  }
});

export default router;
