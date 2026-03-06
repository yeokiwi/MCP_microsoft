import { Router, Request, Response } from "express";
import { listOfficeFiles } from "../tools/executor.js";

const router = Router();

router.get("/", async (req: Request, res: Response) => {
  const { directory, recursive, file_types } = req.query;

  if (!directory || typeof directory !== "string") {
    res.status(400).json({ error: "directory query parameter is required" });
    return;
  }

  const result = await listOfficeFiles({
    directory,
    recursive: recursive === "true",
    file_types: file_types ? String(file_types).split(",") : undefined,
  });

  res.json(result);
});

export default router;
