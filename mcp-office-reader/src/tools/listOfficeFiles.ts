import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { listOfficeFilesInDirectory } from "../utils/fileUtils.js";
import type { ListOfficeFilesResult } from "../types.js";

const inputSchema = z.object({
  directory: z.string().describe("Directory path to scan"),
  recursive: z.boolean().default(false),
  file_types: z
    .array(z.enum(["docx", "pptx", "xlsx", "pdf"]))
    .default(["docx", "pptx", "xlsx", "pdf"]),
});

export function registerListOfficeFilesTool(server: McpServer): void {
  server.tool(
    "list_office_files",
    "Lists all .docx, .pptx, .xlsx, and .pdf files in a given directory.",
    inputSchema.shape,
    async (args) => {
      const { directory, recursive, file_types } = inputSchema.parse(args);

      let files;
      try {
        files = listOfficeFilesInDirectory(directory, recursive, file_types);
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return {
          isError: true,
          content: [{ type: "text", text: `Error: ${msg}` }],
        };
      }

      const result: ListOfficeFilesResult = {
        directory,
        files,
        total_count: files.length,
      };

      if (files.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No ${file_types.join(", ")} files found in: ${directory}${recursive ? " (recursive)" : ""}`,
            },
          ],
        };
      }

      let output = `## Office Files in ${directory}\n\n`;
      output += `**Total:** ${files.length} file(s) | **Types:** ${file_types.join(", ")} | **Recursive:** ${recursive}\n\n`;

      // Group by extension
      const byExt: Record<string, typeof files> = {};
      for (const f of files) {
        byExt[f.extension] = byExt[f.extension] || [];
        byExt[f.extension].push(f);
      }

      for (const [ext, extFiles] of Object.entries(byExt)) {
        output += `### .${ext.toUpperCase()} Files (${extFiles.length})\n\n`;
        for (const f of extFiles) {
          const sizeKb = (f.size_bytes / 1024).toFixed(1);
          const modified = f.modified
            ? new Date(f.modified).toLocaleDateString()
            : "unknown";
          output += `- **${f.name}** (${sizeKb} KB, modified: ${modified})\n`;
          output += `  Path: \`${f.path}\`\n`;
        }
        output += "\n";
      }

      return {
        content: [{ type: "text", text: output }],
      };
    }
  );
}
