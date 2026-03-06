import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as XLSX from "xlsx";
import { validateFilePath, getFileSize } from "../utils/fileUtils.js";
import type { ExcelResult, SheetData } from "../types.js";

const inputSchema = z.object({
  file_path: z
    .string()
    .describe("Absolute or relative path to the .xlsx file"),
  sheets: z
    .array(z.string())
    .optional()
    .describe(
      "Optional: specific sheet names to read. Omit for all sheets."
    ),
  max_rows: z
    .number()
    .int()
    .positive()
    .default(1000)
    .describe("Maximum rows to read per sheet (default 1000)"),
  include_formulas: z
    .boolean()
    .default(false)
    .describe(
      "If true, return raw formulas instead of computed values where available"
    ),
  header_row: z
    .boolean()
    .default(true)
    .describe("Whether to treat the first row as column headers"),
});

type CellValue = string | number | boolean | null;

function getCellValue(
  ws: XLSX.WorkSheet,
  cellRef: string,
  includeFormulas: boolean
): CellValue {
  const cell = ws[cellRef] as XLSX.CellObject | undefined;
  if (!cell) return null;

  if (includeFormulas && cell.f) {
    return `=${cell.f}`;
  }

  if (cell.v === undefined || cell.v === null) return null;

  // Format date cells
  if (cell.t === "d" && cell.v instanceof Date) {
    return cell.v.toISOString();
  }

  return cell.v as CellValue;
}

function parseSheet(
  ws: XLSX.WorkSheet,
  maxRows: number,
  includeFormulas: boolean,
  hasHeaderRow: boolean
): SheetData {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  const numRows = Math.min(range.e.r - range.s.r + 1, maxRows + (hasHeaderRow ? 1 : 0));
  const numCols = range.e.c - range.s.c + 1;

  const allRows: CellValue[][] = [];

  for (let r = range.s.r; r < range.s.r + numRows; r++) {
    const row: CellValue[] = [];
    for (let c = range.s.c; c < range.s.c + numCols; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      row.push(getCellValue(ws, cellRef, includeFormulas));
    }
    allRows.push(row);
  }

  if (hasHeaderRow && allRows.length > 0) {
    const headers = allRows[0].map((h) =>
      h !== null && h !== undefined ? String(h) : ""
    );
    const rows = allRows.slice(1);
    return {
      headers,
      rows,
      row_count: rows.length,
      col_count: numCols,
    };
  }

  return {
    rows: allRows,
    row_count: allRows.length,
    col_count: numCols,
  };
}

export function registerReadExcelTool(server: McpServer): void {
  server.tool(
    "read_excel",
    "Extracts cell data, sheet names, and optionally named ranges from a .xlsx file.",
    inputSchema.shape,
    async (args) => {
      const {
        file_path,
        sheets: sheetFilter,
        max_rows,
        include_formulas,
        header_row,
      } = inputSchema.parse(args);

      const validation = validateFilePath(file_path, ".xlsx");
      if (!validation.valid) {
        return {
          isError: true,
          content: [{ type: "text", text: `Error: ${validation.error}` }],
        };
      }

      const resolvedPath = validation.resolvedPath!;

      let workbook: XLSX.WorkBook;
      try {
        workbook = XLSX.readFile(resolvedPath, {
          cellFormula: include_formulas,
          cellDates: true,
          cellNF: false,
          cellStyles: false,
        });
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        if (
          msg.toLowerCase().includes("password") ||
          msg.toLowerCase().includes("encrypt")
        ) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: "Error: Workbook is password-protected. Please remove the password before reading.",
              },
            ],
          };
        }
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Error: Could not read Excel file (corrupt or invalid .xlsx): ${msg}`,
            },
          ],
        };
      }

      const allSheetNames = workbook.SheetNames;

      // Validate requested sheets
      if (sheetFilter) {
        const notFound = sheetFilter.filter(
          (s) => !allSheetNames.includes(s)
        );
        if (notFound.length > 0) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: `Error: Sheet(s) not found: ${notFound.join(", ")}. Available sheets: ${allSheetNames.join(", ")}`,
              },
            ],
          };
        }
      }

      const targetSheets = sheetFilter || allSheetNames;
      const sheetsData: Record<string, SheetData> = {};

      for (const sheetName of targetSheets) {
        const ws = workbook.Sheets[sheetName];
        if (!ws) continue;

        if (!ws["!ref"]) {
          sheetsData[sheetName] = {
            rows: [],
            row_count: 0,
            col_count: 0,
          };
          continue;
        }

        sheetsData[sheetName] = parseSheet(ws, max_rows, include_formulas, header_row);
      }

      const fileSize = getFileSize(resolvedPath);
      const result: ExcelResult = {
        metadata: {
          sheet_names: allSheetNames,
          file_size_bytes: fileSize,
        },
        sheets: sheetsData,
      };

      // Format output
      let output = "## Workbook Metadata\n\n";
      output += `**Sheets:** ${result.metadata.sheet_names.join(", ")}\n`;
      output += `**File Size:** ${(result.metadata.file_size_bytes / 1024).toFixed(1)} KB\n\n---\n\n`;

      for (const [sheetName, sheetData] of Object.entries(result.sheets)) {
        output += `## Sheet: ${sheetName}\n\n`;
        output += `**Rows:** ${sheetData.row_count} | **Columns:** ${sheetData.col_count}\n\n`;

        if (sheetData.row_count === 0) {
          output += "_Empty sheet_\n\n";
        } else {
          // Render as markdown table
          const headers = sheetData.headers;
          const rows = sheetData.rows;

          if (headers && headers.length > 0) {
            output += `| ${headers.join(" | ")} |\n`;
            output += `| ${headers.map(() => "---").join(" | ")} |\n`;
          }

          for (const row of rows) {
            const cells = row.map((c) => (c !== null && c !== undefined ? String(c) : ""));
            output += `| ${cells.join(" | ")} |\n`;
          }
          output += "\n";
        }

        output += "---\n\n";
      }

      return {
        content: [{ type: "text", text: output }],
      };
    }
  );
}
