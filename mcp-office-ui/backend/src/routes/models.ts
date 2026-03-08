import { Router, Request, Response } from "express";

const router = Router();

const DEFAULT_MODELS = [
  "deepseek-chat",
  "deepseek-reasoner",
  "mistral-small-latest",
  "mistral-large-latest",
  "open-mistral-7b",
  "open-mixtral-8x7b",
  "glm-4",
  "glm-4-flash",
  "glm-4-air",
];

router.get("/", (_req: Request, res: Response) => {
  const raw = process.env.MODELS;
  const models = raw
    ? raw.split(",").map((m) => m.trim()).filter(Boolean)
    : DEFAULT_MODELS;
  res.json({ models });
});

export default router;
