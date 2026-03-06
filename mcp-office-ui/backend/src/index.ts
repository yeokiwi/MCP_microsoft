import "dotenv/config";
import express from "express";
import cors from "cors";
import chatRouter from "./routes/chat.js";
import filesRouter from "./routes/files.js";

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT) : 3001;

app.use(cors());
app.use(express.json({ limit: "10mb" }));

app.use("/api/chat", chatRouter);
app.use("/api/files", filesRouter);

app.get("/api/health", (_req, res) => {
  res.json({ status: "ok", version: "1.0.0" });
});

app.listen(PORT, () => {
  console.log(`MCP Office UI backend running on http://localhost:${PORT}`);
});
