import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerReadWordTool } from "./tools/readWord.js";
import { registerReadPowerPointTool } from "./tools/readPowerPoint.js";
import { registerReadExcelTool } from "./tools/readExcel.js";
import { registerReadPdfTool } from "./tools/readPdf.js";
import { registerListOfficeFilesTool } from "./tools/listOfficeFiles.js";

const server = new McpServer({
  name: "office-reader",
  version: "1.0.0",
});

registerReadWordTool(server);
registerReadPowerPointTool(server);
registerReadExcelTool(server);
registerReadPdfTool(server);
registerListOfficeFilesTool(server);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("Office Reader MCP server running on stdio");
