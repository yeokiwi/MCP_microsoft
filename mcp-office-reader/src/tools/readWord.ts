import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { readFileSync } from "fs";
import AdmZip from "adm-zip";
import mammoth from "mammoth";
import TurndownService from "turndown";
import { XMLParser } from "fast-xml-parser";
import { validateFilePath, getFileSize } from "../utils/fileUtils.js";
import { ensureArray } from "../utils/xmlUtils.js";
import type { DocumentMetadata, TableData, WordDocumentResult } from "../types.js";

const inputSchema = z.object({
  file_path: z.string().describe("Absolute or relative path to the .docx file"),
  output_format: z
    .enum(["markdown", "plain_text", "json"])
    .default("markdown")
    .describe("Format of the extracted content"),
  include_metadata: z
    .boolean()
    .default(true)
    .describe(
      "Whether to include document metadata (author, dates, word count)"
    ),
  outline_only: z
    .boolean()
    .default(false)
    .describe(
      "If true, return only the document outline (list of section headings with 1-indexed numbers) without full content. Use this first on long documents to discover structure before fetching specific sections."
    ),
  sections: z
    .array(z.number().int().positive())
    .optional()
    .describe(
      "Optional: specific 1-indexed section numbers to extract (sections are delimited by headings). Call with outline_only: true first to see available sections."
    ),
});

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  isArray: (name) =>
    ["w:p", "w:r", "w:t", "w:tbl", "w:tr", "w:tc"].includes(name),
});

function extractMetadata(zip: AdmZip): DocumentMetadata {
  try {
    const coreXmlEntry = zip.getEntry("docProps/core.xml");
    if (!coreXmlEntry) return {};

    const coreXml = coreXmlEntry.getData().toString("utf-8");
    const parsed = xmlParser.parse(coreXml) as Record<string, unknown>;
    const props = (parsed["cp:coreProperties"] ||
      parsed["Properties"] ||
      {}) as Record<string, unknown>;

    return {
      title: String(props["dc:title"] || props["title"] || ""),
      author:
        String(
          props["dc:creator"] || props["lastModifiedBy"] || props["author"] || ""
        ) || undefined,
      created: props["dcterms:created"]
        ? String(
            (props["dcterms:created"] as Record<string, unknown>)["#text"] ||
              props["dcterms:created"]
          )
        : undefined,
      modified: props["dcterms:modified"]
        ? String(
            (props["dcterms:modified"] as Record<string, unknown>)["#text"] ||
              props["dcterms:modified"]
          )
        : undefined,
    };
  } catch {
    return {};
  }
}

function extractTablesFromXml(zip: AdmZip): TableData[] {
  const tables: TableData[] = [];
  try {
    const docXmlEntry = zip.getEntry("word/document.xml");
    if (!docXmlEntry) return tables;

    const docXml = docXmlEntry.getData().toString("utf-8");
    const parsed = xmlParser.parse(docXml) as Record<string, unknown>;

    const body = (
      (parsed["w:document"] as Record<string, unknown>)?.["w:body"] as Record<
        string,
        unknown
      >
    );
    if (!body) return tables;

    const tblElements = ensureArray(body["w:tbl"]);
    tblElements.forEach((tbl, tableIndex) => {
      const rows: string[][] = [];
      const tblObj = tbl as Record<string, unknown>;
      const trElements = ensureArray(tblObj["w:tr"]);

      for (const tr of trElements) {
        const row: string[] = [];
        const trObj = tr as Record<string, unknown>;
        const tcElements = ensureArray(trObj["w:tc"]);

        for (const tc of tcElements) {
          const tcObj = tc as Record<string, unknown>;
          const paragraphs = ensureArray(tcObj["w:p"]);
          const cellText = paragraphs
            .map((p) => {
              const pObj = p as Record<string, unknown>;
              const runs = ensureArray(pObj["w:r"]);
              return runs
                .map((r) => {
                  const rObj = r as Record<string, unknown>;
                  const tNodes = ensureArray(rObj["w:t"]);
                  return tNodes
                    .map((t) => {
                      if (typeof t === "string") return t;
                      if (typeof t === "object" && t !== null) {
                        return String(
                          (t as Record<string, unknown>)["#text"] || ""
                        );
                      }
                      return "";
                    })
                    .join("");
                })
                .join("");
            })
            .join("\n");
          row.push(cellText);
        }
        rows.push(row);
      }

      tables.push({ index: tableIndex, rows });
    });
  } catch {
    // Return any tables collected so far
  }
  return tables;
}

function countWords(text: string): number {
  return text
    .trim()
    .split(/\s+/)
    .filter((w) => w.length > 0).length;
}

interface DocSection { index: number; heading: string; content: string; }

function splitIntoSections(content: string): DocSection[] {
  const lines = content.split("\n");
  const sections: DocSection[] = [];
  let sectionIdx = 0;
  let currentHeading = "";
  let currentLines: string[] = [];

  for (const line of lines) {
    if (/^#{1,6} /.test(line)) {
      if (currentLines.length > 0 || currentHeading) {
        sections.push({ index: sectionIdx++, heading: currentHeading || "(preamble)", content: currentLines.join("\n").trim() });
      }
      currentHeading = line.replace(/^#+\s*/, "").trim();
      currentLines = [line];
    } else {
      currentLines.push(line);
    }
  }
  if (currentLines.length > 0 || currentHeading) {
    sections.push({ index: sectionIdx, heading: currentHeading || "(document)", content: currentLines.join("\n").trim() });
  }
  return sections;
}

export function registerReadWordTool(server: McpServer): void {
  server.tool(
    "read_word_document",
    "Extracts text content, headings, tables, and metadata from a .docx file.",
    inputSchema.shape,
    async (args) => {
      const { file_path, output_format, include_metadata, outline_only, sections: sectionFilter } =
        inputSchema.parse(args);

      const validation = validateFilePath(file_path, ".docx");
      if (!validation.valid) {
        return {
          isError: true,
          content: [{ type: "text", text: `Error: ${validation.error}` }],
        };
      }

      const resolvedPath = validation.resolvedPath!;

      let zip: AdmZip;
      try {
        zip = new AdmZip(resolvedPath);
      } catch (error) {
        const msg =
          error instanceof Error ? error.message : String(error);
        if (
          msg.toLowerCase().includes("password") ||
          msg.toLowerCase().includes("encrypt")
        ) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: "Error: Document is password-protected. Please remove the password before reading.",
              },
            ],
          };
        }
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Error: Could not open file as a ZIP archive (corrupt or invalid .docx): ${msg}`,
            },
          ],
        };
      }

      // Extract content using mammoth
      let content = "";
      try {
        const buffer = readFileSync(resolvedPath);
        if (output_format === "plain_text") {
          const result = await mammoth.extractRawText({ buffer });
          content = result.value;
        } else {
          const result = await mammoth.convertToHtml({ buffer });
          const td = new TurndownService({
            headingStyle: "atx",
            bulletListMarker: "-",
          });
          content = output_format === "markdown" ? td.turndown(result.value) : result.value;
        }
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Error: Failed to extract content from Word document: ${msg}`,
            },
          ],
        };
      }

      // Section-based handling (only for markdown/html formats, not plain_text)
      const allSections = output_format !== "plain_text" ? splitIntoSections(content) : [];

      // outline_only: return just the document structure
      if (outline_only) {
        const meta = include_metadata ? extractMetadata(zip) : {};
        let outlineText = "## Document Outline\n\n";
        if (meta.title) outlineText += `**Title:** ${meta.title}\n`;
        if (meta.author) outlineText += `**Author:** ${meta.author}\n`;
        outlineText += `**Word Count:** ${countWords(content)}\n`;
        outlineText += `**Sections:** ${allSections.length}\n\n`;
        outlineText += "### Sections\n\n";
        for (const sec of allSections) {
          outlineText += `${sec.index + 1}. ${sec.heading}\n`;
        }
        outlineText += "\n\n_Call again with `sections: [n, m, ...]` to fetch specific sections._";
        return { content: [{ type: "text", text: outlineText }] };
      }

      // sections filter: extract only selected sections
      if (sectionFilter && sectionFilter.length > 0 && allSections.length > 0) {
        const bad = sectionFilter.filter((s) => s < 1 || s > allSections.length);
        if (bad.length) {
          return {
            isError: true,
            content: [{ type: "text", text: `Error: Section(s) out of range: ${bad.join(", ")} (total: ${allSections.length}). Use outline_only: true to see available sections.` }],
          };
        }
        content = sectionFilter.map((s) => allSections[s - 1].content).join("\n\n");
      }

      // Extract tables from XML
      const tables = extractTablesFromXml(zip);

      // Build result
      const result: WordDocumentResult = {
        content,
        tables,
      };

      if (include_metadata) {
        const meta = extractMetadata(zip);
        result.metadata = {
          ...meta,
          word_count: countWords(content),
          title: meta.title || undefined,
          author: meta.author || undefined,
        };
      }

      if (output_format === "json") {
        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      }

      // Format as readable text
      let output = "";
      if (result.metadata) {
        output += "## Document Metadata\n\n";
        if (result.metadata.title) output += `**Title:** ${result.metadata.title}\n`;
        if (result.metadata.author) output += `**Author:** ${result.metadata.author}\n`;
        if (result.metadata.created) output += `**Created:** ${result.metadata.created}\n`;
        if (result.metadata.modified) output += `**Modified:** ${result.metadata.modified}\n`;
        if (result.metadata.word_count !== undefined)
          output += `**Word Count:** ${result.metadata.word_count}\n`;
        if (allSections.length > 0)
          output += `**Total Sections:** ${allSections.length}\n`;
        output += "\n---\n\n";
      }

      output += result.content;

      if (tables.length > 0) {
        output += "\n\n## Tables\n\n";
        for (const table of tables) {
          output += `### Table ${table.index + 1}\n\n`;
          for (const row of table.rows) {
            output += `| ${row.join(" | ")} |\n`;
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
