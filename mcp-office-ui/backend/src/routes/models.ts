import { Router, Request, Response } from "express";

const router = Router();

const DEFAULT_MODELS = [
  "gpt-4o",
  "gpt-4o-mini",
  "gpt-4-turbo",
  "gpt-3.5-turbo",
];

router.get("/", (_req: Request, res: Response) => {
  const raw = process.env.MODELS;
  const models = raw
    ? raw.split(",").map((m) => m.trim()).filter(Boolean)
    : DEFAULT_MODELS;
  res.json({ models });
});

export default router;
