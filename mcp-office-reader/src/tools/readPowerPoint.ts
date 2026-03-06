import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import AdmZip from "adm-zip";
import { XMLParser } from "fast-xml-parser";
import { validateFilePath } from "../utils/fileUtils.js";
import { ensureArray } from "../utils/xmlUtils.js";
import type {
  PowerPointResult,
  PresentationMetadata,
  Slide,
  SlideShape,
} from "../types.js";

const inputSchema = z.object({
  file_path: z
    .string()
    .describe("Absolute or relative path to the .pptx file"),
  slides: z
    .array(z.number().int().positive())
    .optional()
    .describe(
      "Optional: specific 1-indexed slide numbers to extract. Omit for all slides."
    ),
  include_notes: z
    .boolean()
    .default(true)
    .describe("Whether to include speaker notes for each slide"),
  include_shapes: z
    .boolean()
    .default(false)
    .describe("Whether to include shape/element metadata per slide"),
});

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  isArray: (name) =>
    [
      "p:sp",
      "p:pic",
      "p:graphicFrame",
      "a:p",
      "a:r",
      "a:t",
      "Relationship",
    ].includes(name),
});

interface RelationshipItem {
  "@_Id"?: string;
  "@_Type"?: string;
  "@_Target"?: string;
}

function getSlideOrder(zip: AdmZip): string[] {
  try {
    const relsEntry = zip.getEntry("ppt/_rels/presentation.xml.rels");
    if (!relsEntry) {
      // Fall back to directory listing
      return zip
        .getEntries()
        .filter(
          (e) =>
            e.entryName.startsWith("ppt/slides/slide") &&
            e.entryName.endsWith(".xml") &&
            !e.entryName.includes("/_rels/")
        )
        .map((e) => e.entryName)
        .sort();
    }

    const relsXml = relsEntry.getData().toString("utf-8");
    const parsed = xmlParser.parse(relsXml) as Record<string, unknown>;
    const relationships = parsed["Relationships"] as Record<string, unknown>;
    const rels = ensureArray(
      relationships?.["Relationship"] as RelationshipItem | RelationshipItem[]
    );

    const slideRels = rels
      .filter(
        (r) =>
          r["@_Type"]?.includes("slide") &&
          !r["@_Type"]?.includes("slideLayout") &&
          !r["@_Type"]?.includes("slideMaster")
      )
      .sort((a, b) => {
        // Sort by Id (rId1, rId2, etc.)
        const aId = parseInt(a["@_Id"]?.replace("rId", "") || "0");
        const bId = parseInt(b["@_Id"]?.replace("rId", "") || "0");
        return aId - bId;
      });

    return slideRels
      .map((r) => {
        const target = r["@_Target"] || "";
        return target.startsWith("slides/")
          ? `ppt/${target}`
          : `ppt/slides/${target}`;
      })
      .filter(Boolean);
  } catch {
    return zip
      .getEntries()
      .filter(
        (e) =>
          e.entryName.startsWith("ppt/slides/slide") &&
          e.entryName.endsWith(".xml") &&
          !e.entryName.includes("/_rels/")
      )
      .map((e) => e.entryName)
      .sort();
  }
}

function getNotesPath(
  zip: AdmZip,
  slideEntryName: string
): string | null {
  try {
    // Get the relationships for this slide
    const slideFile = slideEntryName.split("/").pop()!;
    const relsPath = `ppt/slides/_rels/${slideFile}.rels`;
    const relsEntry = zip.getEntry(relsPath);
    if (!relsEntry) return null;

    const relsXml = relsEntry.getData().toString("utf-8");
    const parsed = xmlParser.parse(relsXml) as Record<string, unknown>;
    const relationships = parsed["Relationships"] as Record<string, unknown>;
    const rels = ensureArray(
      relationships?.["Relationship"] as RelationshipItem | RelationshipItem[]
    );

    const notesRel = rels.find((r) =>
      r["@_Type"]?.includes("notesSlide")
    );
    if (!notesRel || !notesRel["@_Target"]) return null;

    const target = notesRel["@_Target"];
    return target.startsWith("../")
      ? `ppt/${target.replace("../", "")}`
      : `ppt/slides/${target}`;
  } catch {
    return null;
  }
}

function extractTextRuns(node: unknown): string {
  if (!node || typeof node !== "object") return "";
  const obj = node as Record<string, unknown>;

  // a:t nodes contain actual text
  if ("#text" in obj) return String(obj["#text"]);

  let text = "";
  const runs = ensureArray(obj["a:r"] as unknown);
  for (const run of runs) {
    const runObj = run as Record<string, unknown>;
    const tNodes = ensureArray(runObj["a:t"] as unknown);
    for (const t of tNodes) {
      if (typeof t === "string") {
        text += t;
      } else if (typeof t === "object" && t !== null) {
        text += String((t as Record<string, unknown>)["#text"] || "");
      }
    }
  }
  return text;
}

function extractSlideContent(
  slideXml: string
): { title: string | undefined; content: string; shapes: SlideShape[] } {
  let title: string | undefined;
  const contentLines: string[] = [];
  const shapes: SlideShape[] = [];

  try {
    const parsed = xmlParser.parse(slideXml) as Record<string, unknown>;
    const spTree = (
      (parsed["p:sld"] as Record<string, unknown>)?.["p:cSld"] as Record<
        string,
        unknown
      >
    )?.["p:spTree"] as Record<string, unknown>;

    if (!spTree) return { title, content: "", shapes };

    const spElements = ensureArray(spTree["p:sp"] as unknown);

    for (const sp of spElements) {
      const spObj = sp as Record<string, unknown>;
      const nvSpPr = spObj["p:nvSpPr"] as Record<string, unknown>;
      const spPr = nvSpPr?.["p:nvPr"] as Record<string, unknown>;
      const ph = spPr?.["p:ph"] as Record<string, unknown>;
      const phType = ph?.["@_type"] as string | undefined;
      const isTitle =
        phType === "title" || phType === "ctrTitle" || phType === "subTitle";

      const txBody = spObj["p:txBody"] as Record<string, unknown>;
      if (!txBody) continue;

      const paragraphs = ensureArray(txBody["a:p"] as unknown);
      const shapeText = paragraphs
        .map((p) => extractTextRuns(p))
        .filter((t) => t.trim().length > 0)
        .join("\n");

      if (!shapeText) continue;

      if (isTitle && !title) {
        title = shapeText;
      } else {
        contentLines.push(shapeText);
      }

      shapes.push({
        type: phType || "body",
        text: shapeText,
      });
    }
  } catch {
    // Return what we have
  }

  return { title, content: contentLines.join("\n"), shapes };
}

function extractNotesContent(notesXml: string): string {
  try {
    const parsed = xmlParser.parse(notesXml) as Record<string, unknown>;
    const spTree = (
      (parsed["p:notes"] as Record<string, unknown>)?.["p:cSld"] as Record<
        string,
        unknown
      >
    )?.["p:spTree"] as Record<string, unknown>;

    if (!spTree) return "";

    const spElements = ensureArray(spTree["p:sp"] as unknown);
    const notesTexts: string[] = [];

    for (const sp of spElements) {
      const spObj = sp as Record<string, unknown>;
      const nvSpPr = spObj["p:nvSpPr"] as Record<string, unknown>;
      const spPr = nvSpPr?.["p:nvPr"] as Record<string, unknown>;
      const ph = spPr?.["p:ph"] as Record<string, unknown>;
      const phType = ph?.["@_type"] as string | undefined;

      // Skip the slide image placeholder (type="sldImg") and only get body text
      if (phType === "sldImg") continue;

      const txBody = spObj["p:txBody"] as Record<string, unknown>;
      if (!txBody) continue;

      const paragraphs = ensureArray(txBody["a:p"] as unknown);
      const text = paragraphs
        .map((p) => extractTextRuns(p))
        .filter((t) => t.trim().length > 0)
        .join("\n");

      if (text) notesTexts.push(text);
    }

    return notesTexts.join("\n");
  } catch {
    return "";
  }
}

function extractPresentationMetadata(zip: AdmZip): Partial<PresentationMetadata> {
  try {
    const coreEntry = zip.getEntry("docProps/core.xml");
    if (!coreEntry) return {};

    const coreXml = coreEntry.getData().toString("utf-8");
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      textNodeName: "#text",
    });
    const parsed = parser.parse(coreXml) as Record<string, unknown>;
    const props = (parsed["cp:coreProperties"] || {}) as Record<string, unknown>;

    return {
      title: String(props["dc:title"] || "") || undefined,
      author: String(props["dc:creator"] || "") || undefined,
    };
  } catch {
    return {};
  }
}

export function registerReadPowerPointTool(server: McpServer): void {
  server.tool(
    "read_powerpoint",
    "Extracts slide content, speaker notes, and structure from a .pptx file.",
    inputSchema.shape,
    async (args) => {
      const { file_path, slides: slideFilter, include_notes, include_shapes } =
        inputSchema.parse(args);

      const validation = validateFilePath(file_path, ".pptx");
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
        const msg = error instanceof Error ? error.message : String(error);
        if (msg.toLowerCase().includes("password") || msg.toLowerCase().includes("encrypt")) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: "Error: Presentation is password-protected. Please remove the password before reading.",
              },
            ],
          };
        }
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Error: Could not open file as a ZIP archive (corrupt or invalid .pptx): ${msg}`,
            },
          ],
        };
      }

      const slideEntries = getSlideOrder(zip);

      if (slideEntries.length === 0) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "Error: No slides found in the presentation.",
            },
          ],
        };
      }

      // Validate slide filter
      if (slideFilter) {
        const maxSlide = slideEntries.length;
        const invalidSlides = slideFilter.filter((s) => s > maxSlide || s < 1);
        if (invalidSlides.length > 0) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: `Error: Slide indices out of range: ${invalidSlides.join(", ")}. Presentation has ${maxSlide} slides.`,
              },
            ],
          };
        }
      }

      const meta = extractPresentationMetadata(zip);
      const parsedSlides: Slide[] = [];

      for (let i = 0; i < slideEntries.length; i++) {
        const slideNum = i + 1;
        if (slideFilter && !slideFilter.includes(slideNum)) continue;

        const entryName = slideEntries[i];
        const slideEntry = zip.getEntry(entryName);
        if (!slideEntry) continue;

        const slideXml = slideEntry.getData().toString("utf-8");
        const { title, content, shapes } = extractSlideContent(slideXml);

        const slide: Slide = {
          index: slideNum,
          title,
          content,
        };

        if (include_notes) {
          const notesPath = getNotesPath(zip, entryName);
          if (notesPath) {
            const notesEntry = zip.getEntry(notesPath);
            if (notesEntry) {
              const notesXml = notesEntry.getData().toString("utf-8");
              slide.notes = extractNotesContent(notesXml);
            }
          }
          if (slide.notes === undefined) slide.notes = "";
        }

        if (include_shapes) {
          slide.shapes = shapes;
        }

        parsedSlides.push(slide);
      }

      const result: PowerPointResult = {
        metadata: {
          title: meta.title,
          author: meta.author,
          slide_count: slideEntries.length,
        },
        slides: parsedSlides,
      };

      // Format output
      let output = "## Presentation Metadata\n\n";
      if (result.metadata.title) output += `**Title:** ${result.metadata.title}\n`;
      if (result.metadata.author) output += `**Author:** ${result.metadata.author}\n`;
      output += `**Slide Count:** ${result.metadata.slide_count}\n\n---\n\n`;

      for (const slide of result.slides) {
        output += `## Slide ${slide.index}`;
        if (slide.title) output += `: ${slide.title}`;
        output += "\n\n";

        if (slide.content) {
          output += `${slide.content}\n\n`;
        }

        if (include_notes && slide.notes) {
          output += `**Speaker Notes:** ${slide.notes}\n\n`;
        }

        output += "---\n\n";
      }

      return {
        content: [{ type: "text", text: output }],
      };
    }
  );
}
