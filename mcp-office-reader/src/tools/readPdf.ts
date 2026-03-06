import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { readFileSync } from "fs";
import { execFile } from "child_process";
import { promisify } from "util";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";
import pdfParse from "pdf-parse";
import { validateFilePath, getFileSize } from "../utils/fileUtils.js";
import { isScannedPdf, ocrPdfPages } from "../utils/ocrUtils.js";
import type { PdfResult, PdfPage, PdfTable, PdfMetadata } from "../types.js";

const execFileAsync = promisify(execFile);

const inputSchema = z.object({
  file_path: z
    .string()
    .describe("Absolute or relative path to the .pdf file"),
  pages: z
    .array(z.number().int().positive())
    .optional()
    .describe(
      "Optional: specific 1-indexed page numbers to extract. Omit for all pages."
    ),
  output_format: z
    .enum(["markdown", "plain_text", "json"])
    .default("markdown")
    .describe("Format of the extracted content"),
  extract_tables: z
    .boolean()
    .default(false)
    .describe(
      "Whether to extract tables (uses Python pdfplumber subprocess — requires Python 3 + pdfplumber installed)"
    ),
  ocr_fallback: z
    .boolean()
    .default(true)
    .describe(
      "If text extraction yields empty results (scanned PDF), automatically run Tesseract OCR"
    ),
  include_metadata: z
    .boolean()
    .default(true)
    .describe(
      "Whether to include document metadata (title, author, page count, etc.)"
    ),
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

function getScriptsDir(): string {
  // Navigate from src/tools/ to scripts/
  return resolve(__dirname, "../../scripts");
}

async function checkPythonAndPdfplumber(): Promise<{
  available: boolean;
  error?: string;
}> {
  try {
    await execFileAsync("python3", ["-c", "import pdfplumber"]);
    return { available: true };
  } catch {
    try {
      // Try 'python' as fallback
      await execFileAsync("python", ["-c", "import pdfplumber"]);
      return { available: true };
    } catch {
      return {
        available: false,
        error:
          "Python 3 with pdfplumber is required for table extraction. " +
          "Install with: pip install pdfplumber",
      };
    }
  }
}

async function extractTablesWithPdfplumber(
  filePath: string
): Promise<PdfTable[]> {
  const scriptPath = resolve(getScriptsDir(), "extract_pdf_tables.py");

  let stdout: string;
  try {
    // Try python3 first
    try {
      const result = await execFileAsync("python3", [scriptPath, filePath]);
      stdout = result.stdout;
    } catch {
      const result = await execFileAsync("python", [scriptPath, filePath]);
      stdout = result.stdout;
    }
  } catch (error) {
    throw new Error(
      `Failed to run PDF table extraction script: ${
        error instanceof Error ? error.message : String(error)
      }`
    );
  }

  try {
    const tables = JSON.parse(stdout) as Array<{
      page: number;
      rows: (string | null)[][];
    }>;
    return tables.map((t, i) => ({
      page: t.page,
      index: i,
      rows: t.rows,
    }));
  } catch {
    return [];
  }
}

function parseDateFromPdfInfo(dateStr: string | undefined): string | undefined {
  if (!dateStr) return undefined;
  // PDF dates are often in format: D:YYYYMMDDHHmmSSOHH'mm'
  const match = dateStr.match(
    /D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/
  );
  if (match) {
    const [, year, month, day, hour, min, sec] = match;
    return `${year}-${month}-${day}T${hour}:${min}:${sec}Z`;
  }
  return dateStr;
}

export function registerReadPdfTool(server: McpServer): void {
  server.tool(
    "read_pdf",
    "Extracts text, tables, and metadata from a .pdf file. Handles both native and scanned PDFs with OCR fallback.",
    inputSchema.shape,
    async (args) => {
      const {
        file_path,
        pages: pageFilter,
        output_format,
        extract_tables,
        ocr_fallback,
        include_metadata,
      } = inputSchema.parse(args);

      const validation = validateFilePath(file_path, ".pdf");
      if (!validation.valid) {
        return {
          isError: true,
          content: [{ type: "text", text: `Error: ${validation.error}` }],
        };
      }

      const resolvedPath = validation.resolvedPath!;

      // Check pdfplumber availability if needed
      if (extract_tables) {
        const pythonCheck = await checkPythonAndPdfplumber();
        if (!pythonCheck.available) {
          return {
            isError: true,
            content: [{ type: "text", text: `Error: ${pythonCheck.error}` }],
          };
        }
      }

      const buffer = readFileSync(resolvedPath);
      const fileSize = getFileSize(resolvedPath);

      let pdfData: Awaited<ReturnType<typeof pdfParse>>;
      let isEncrypted = false;

      // Per-page text storage
      const pageTexts: string[] = [];

      try {
        pdfData = await pdfParse(buffer, {
          pagerender: (pageData: {
            getTextContent: () => Promise<{
              items: Array<{ str: string }>;
            }>;
          }) => {
            return pageData.getTextContent().then((content) => {
              pageTexts.push(
                content.items.map((i) => i.str).join(" ")
              );
              return "";
            });
          },
        });
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        if (
          msg.toLowerCase().includes("password") ||
          msg.toLowerCase().includes("encrypt") ||
          msg.toLowerCase().includes("encrypted")
        ) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: "Error: PDF is password-protected. Provide the password to decrypt.",
              },
            ],
          };
        }
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Error: Failed to parse PDF: ${msg}`,
            },
          ],
        };
      }

      const totalPages = pdfData.numpages;
      const pdfInfo = pdfData.info as Record<string, unknown>;

      // Detect scanned PDF
      const scanned = isScannedPdf(pdfData.text, totalPages);

      // Validate page filter
      if (pageFilter) {
        const invalidPages = pageFilter.filter(
          (p) => p > totalPages || p < 1
        );
        if (invalidPages.length > 0) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: `Error: Page numbers out of range: ${invalidPages.join(", ")}. PDF has ${totalPages} pages.`,
              },
            ],
          };
        }
      }

      // Determine which pages to process
      const pagesToProcess = pageFilter
        ? pageFilter
        : Array.from({ length: totalPages }, (_, i) => i + 1);

      // Build pages
      const resultPages: PdfPage[] = [];

      if (scanned && ocr_fallback) {
        // OCR path
        const ocrResults = await ocrPdfPages(resolvedPath, pagesToProcess);

        for (const ocrResult of ocrResults) {
          const content = ocrResult.text || "";
          resultPages.push({
            index: ocrResult.pageIndex,
            content,
            word_count: content
              .trim()
              .split(/\s+/)
              .filter((w) => w.length > 0).length,
            ocr_used: true,
          });
        }
      } else if (scanned && !ocr_fallback) {
        // No OCR but no text either
        return {
          isError: false,
          content: [
            {
              type: "text",
              text:
                "Warning: This appears to be a scanned PDF with no extractable text. " +
                "Enable `ocr_fallback: true` to use Tesseract OCR for text extraction.",
            },
          ],
        };
      } else {
        // Native text PDF
        for (const pageNum of pagesToProcess) {
          const pageIdx = pageNum - 1;
          const text = pageTexts[pageIdx] || "";
          resultPages.push({
            index: pageNum,
            content: text,
            word_count: text
              .trim()
              .split(/\s+/)
              .filter((w) => w.length > 0).length,
            ocr_used: false,
          });
        }
      }

      // Extract tables if requested
      let tables: PdfTable[] | undefined;
      if (extract_tables) {
        try {
          const allTables = await extractTablesWithPdfplumber(resolvedPath);
          // Filter tables to only requested pages
          if (pageFilter) {
            tables = allTables.filter((t) => pageFilter.includes(t.page));
          } else {
            tables = allTables;
          }
        } catch (error) {
          // Return tables extraction error as warning in output
          tables = [];
        }
      }

      const metadata: PdfMetadata | undefined = include_metadata
        ? {
            title: String(pdfInfo?.["Title"] || "") || undefined,
            author: String(pdfInfo?.["Author"] || "") || undefined,
            page_count: totalPages,
            file_size_bytes: fileSize,
            pdf_version: pdfData.version || undefined,
            is_scanned: scanned,
            encrypted: isEncrypted,
          }
        : undefined;

      const result: PdfResult = {
        metadata,
        pages: resultPages,
        tables,
      };

      if (output_format === "json") {
        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      }

      // Format as readable text
      let output = "";

      if (metadata) {
        output += "## PDF Metadata\n\n";
        if (metadata.title) output += `**Title:** ${metadata.title}\n`;
        if (metadata.author) output += `**Author:** ${metadata.author}\n`;
        output += `**Page Count:** ${metadata.page_count}\n`;
        output += `**File Size:** ${(metadata.file_size_bytes / 1024).toFixed(1)} KB\n`;
        if (metadata.pdf_version) output += `**PDF Version:** ${metadata.pdf_version}\n`;
        output += `**Scanned (Image-based):** ${metadata.is_scanned}\n`;
        output += `**Encrypted:** ${metadata.encrypted}\n`;
        output += "\n---\n\n";
      }

      for (const page of resultPages) {
        output += `## Page ${page.index}`;
        if (page.ocr_used) output += " _(OCR)_";
        output += "\n\n";
        output += page.content || "_No text content_";
        output += `\n\n_Word count: ${page.word_count}_\n\n---\n\n`;
      }

      if (tables && tables.length > 0) {
        output += "## Tables\n\n";
        for (const table of tables) {
          output += `### Table ${table.index + 1} (Page ${table.page})\n\n`;
          for (const row of table.rows) {
            const cells = row.map((c) => (c !== null ? String(c) : ""));
            output += `| ${cells.join(" | ")} |\n`;
          }
          output += "\n";
        }
      }

      return {
        content: [{ type: "text", text: output }],
      };
    }
  );
}
