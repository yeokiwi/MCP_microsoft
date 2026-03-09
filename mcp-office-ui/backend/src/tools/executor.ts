/**
 * Tool executor: implements Office/PDF reading tools as plain async functions.
 * Re-uses the same libraries as mcp-office-reader but without MCP wrappers.
 */

import { existsSync, readFileSync, statSync, readdirSync } from "fs";
import { resolve, extname, basename, join } from "path";
import AdmZip from "adm-zip";
import mammoth from "mammoth";
import TurndownService from "turndown";
import * as XLSX from "xlsx";
import pdfParse from "pdf-parse";
import { XMLParser } from "fast-xml-parser";
import { execFile } from "child_process";
import { promisify } from "util";
import { fileURLToPath } from "url";
import { dirname } from "path";

const execFileAsync = promisify(execFile);

// ── Shared helpers ────────────────────────────────────────────────────────────

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  isArray: (name) =>
    ["w:p","w:r","w:t","w:tbl","w:tr","w:tc","a:p","a:r","a:t","p:sp","Relationship"].includes(name),
});

function ensureArray<T>(v: T | T[] | null | undefined): T[] {
  if (v == null) return [];
  return Array.isArray(v) ? v : [v];
}

function validateFile(filePath: string, ext?: string): { ok: boolean; path: string; error?: string } {
  const resolved = resolve(filePath);
  if (!existsSync(resolved)) return { ok: false, path: resolved, error: `File not found: ${resolved}` };
  const stat = statSync(resolved);
  if (!stat.isFile()) return { ok: false, path: resolved, error: `Not a file: ${resolved}` };
  if (stat.size === 0) return { ok: false, path: resolved, error: `File is empty: ${resolved}` };
  if (ext && extname(resolved).toLowerCase() !== ext)
    return { ok: false, path: resolved, error: `Expected ${ext} file, got ${extname(resolved)}` };
  return { ok: true, path: resolved };
}

// ── Word ──────────────────────────────────────────────────────────────────────

export interface WordSection {
  index: number;
  heading: string;
  content: string;
}

export interface WordResult {
  metadata?: { title?: string; author?: string; created?: string; modified?: string; word_count?: number; section_count?: number };
  content: string;
  tables: { index: number; rows: string[][] }[];
  sections?: { index: number; heading: string }[];
  error?: string;
}

/** Split markdown content into sections delimited by any heading level. */
function splitIntoSections(content: string): WordSection[] {
  const lines = content.split("\n");
  const sections: WordSection[] = [];
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

export async function readWordDocument(args: {
  file_path: string;
  output_format?: "markdown" | "plain_text" | "json";
  include_metadata?: boolean;
  outline_only?: boolean;
  sections?: number[];
}): Promise<WordResult> {
  const fmt = args.output_format ?? "markdown";
  const { ok, path: p, error } = validateFile(args.file_path, ".docx");
  if (!ok) return { content: "", tables: [], error };

  let zip: AdmZip;
  try { zip = new AdmZip(p); }
  catch (e) {
    const msg = (e as Error).message ?? "";
    if (msg.toLowerCase().includes("password") || msg.toLowerCase().includes("encrypt"))
      return { content: "", tables: [], error: "Document is password-protected." };
    return { content: "", tables: [], error: `Corrupt or invalid .docx: ${msg}` };
  }

  // Extract content
  let content = "";
  try {
    const buf = readFileSync(p);
    if (fmt === "plain_text") {
      content = (await mammoth.extractRawText({ buffer: buf })).value;
    } else {
      const html = (await mammoth.convertToHtml({ buffer: buf })).value;
      const td = new TurndownService({ headingStyle: "atx", bulletListMarker: "-" });
      content = td.turndown(html);
    }
  } catch (e) {
    return { content: "", tables: [], error: `Failed to extract content: ${(e as Error).message}` };
  }

  // Extract tables from XML
  const tables: { index: number; rows: string[][] }[] = [];
  try {
    const docEntry = zip.getEntry("word/document.xml");
    if (docEntry) {
      const parsed = xmlParser.parse(docEntry.getData().toString("utf-8")) as Record<string, unknown>;
      const body = ((parsed["w:document"] as Record<string, unknown>)?.["w:body"]) as Record<string, unknown>;
      if (body) {
        ensureArray(body["w:tbl"]).forEach((tbl, i) => {
          const rows: string[][] = [];
          const tblObj = tbl as Record<string, unknown>;
          for (const tr of ensureArray(tblObj["w:tr"])) {
            const row: string[] = [];
            for (const tc of ensureArray((tr as Record<string, unknown>)["w:tc"])) {
              const tcObj = tc as Record<string, unknown>;
              const cellText = ensureArray(tcObj["w:p"])
                .map((p) => ensureArray((p as Record<string, unknown>)["w:r"])
                  .map((r) => ensureArray((r as Record<string, unknown>)["w:t"])
                    .map((t) => typeof t === "string" ? t : String((t as Record<string, unknown>)["#text"] ?? ""))
                    .join(""))
                  .join(""))
                .join("\n");
              row.push(cellText);
            }
            rows.push(row);
          }
          tables.push({ index: i, rows });
        });
      }
    }
  } catch { /* best-effort */ }

  // Section-based filtering (only applicable to markdown format)
  const allSections = fmt !== "plain_text" ? splitIntoSections(content) : [];
  const sectionFilter = args.sections;

  if (args.outline_only) {
    // Return just the document outline (headings) without full content
    const outline = allSections.map(({ index, heading }) => ({ index, heading }));
    const wordCount = content.trim().split(/\s+/).filter(Boolean).length;
    return {
      metadata: args.include_metadata !== false ? {
        word_count: wordCount,
        section_count: allSections.length,
      } : undefined,
      content: "",
      tables: [],
      sections: outline,
    };
  }

  if (sectionFilter && sectionFilter.length > 0 && allSections.length > 0) {
    const bad = sectionFilter.filter((s) => s < 1 || s > allSections.length);
    if (bad.length) return { content: "", tables: [], error: `Section(s) out of range: ${bad.join(", ")} (total sections: ${allSections.length}). Call with outline_only: true to see available sections.` };
    content = sectionFilter.map((s) => allSections[s - 1].content).join("\n\n");
  }

  // Metadata
  let metadata: WordResult["metadata"] | undefined;
  if (args.include_metadata !== false) {
    try {
      const coreEntry = zip.getEntry("docProps/core.xml");
      if (coreEntry) {
        const c = xmlParser.parse(coreEntry.getData().toString("utf-8")) as Record<string, unknown>;
        const props = (c["cp:coreProperties"] ?? {}) as Record<string, unknown>;
        metadata = {
          title: String(props["dc:title"] ?? "") || undefined,
          author: String(props["dc:creator"] ?? "") || undefined,
          created: ((props["dcterms:created"] as Record<string, unknown>)?.["#text"] as string) ?? undefined,
          modified: ((props["dcterms:modified"] as Record<string, unknown>)?.["#text"] as string) ?? undefined,
          word_count: content.trim().split(/\s+/).filter(Boolean).length,
          section_count: allSections.length,
        };
      }
    } catch { /* ignore */ }
  }

  return { metadata, content, tables };
}

// ── PowerPoint ────────────────────────────────────────────────────────────────

export interface PptResult {
  metadata: { title?: string; author?: string; slide_count: number };
  slides: { index: number; title?: string; content: string; notes?: string }[];
  error?: string;
}

export async function readPowerPoint(args: {
  file_path: string;
  slides?: number[];
  include_notes?: boolean;
}): Promise<PptResult> {
  const { ok, path: p, error } = validateFile(args.file_path, ".pptx");
  if (!ok) return { metadata: { slide_count: 0 }, slides: [], error };

  let zip: AdmZip;
  try { zip = new AdmZip(p); }
  catch (e) {
    const msg = (e as Error).message ?? "";
    if (msg.toLowerCase().includes("password") || msg.toLowerCase().includes("encrypt"))
      return { metadata: { slide_count: 0 }, slides: [], error: "Presentation is password-protected." };
    return { metadata: { slide_count: 0 }, slides: [], error: `Corrupt or invalid .pptx: ${msg}` };
  }

  // Slide order from presentation.xml.rels
  interface RelItem { "@_Type"?: string; "@_Target"?: string; "@_Id"?: string; }
  const slideEntries = (() => {
    try {
      const relsEntry = zip.getEntry("ppt/_rels/presentation.xml.rels");
      if (!relsEntry) throw new Error("no rels");
      const parsed = xmlParser.parse(relsEntry.getData().toString("utf-8")) as Record<string, unknown>;
      const rels = ensureArray((parsed["Relationships"] as Record<string, unknown>)?.["Relationship"] as RelItem | RelItem[]);
      return rels
        .filter((r) => r["@_Type"]?.includes("slide") && !r["@_Type"]?.includes("Layout") && !r["@_Type"]?.includes("Master"))
        .sort((a, b) => parseInt(a["@_Id"]?.replace("rId","") ?? "0") - parseInt(b["@_Id"]?.replace("rId","") ?? "0"))
        .map((r) => { const t = r["@_Target"] ?? ""; return t.startsWith("slides/") ? `ppt/${t}` : `ppt/slides/${t}`; });
    } catch {
      return zip.getEntries()
        .filter((e) => e.entryName.startsWith("ppt/slides/slide") && e.entryName.endsWith(".xml") && !e.entryName.includes("/_rels/"))
        .map((e) => e.entryName).sort();
    }
  })();

  if (slideEntries.length === 0) return { metadata: { slide_count: 0 }, slides: [], error: "No slides found." };

  const slideFilter = args.slides;
  if (slideFilter) {
    const bad = slideFilter.filter((s) => s < 1 || s > slideEntries.length);
    if (bad.length) return { metadata: { slide_count: slideEntries.length }, slides: [], error: `Slide(s) out of range: ${bad.join(", ")} (total: ${slideEntries.length})` };
  }

  function extractText(node: unknown): string {
    if (!node || typeof node !== "object") return "";
    const obj = node as Record<string, unknown>;
    if ("#text" in obj) return String(obj["#text"]);
    return ensureArray(obj["a:r"]).map((r) => {
      return ensureArray((r as Record<string, unknown>)["a:t"]).map((t) =>
        typeof t === "string" ? t : String((t as Record<string, unknown>)["#text"] ?? "")).join("");
    }).join("");
  }

  const slides: PptResult["slides"] = [];
  for (let i = 0; i < slideEntries.length; i++) {
    const num = i + 1;
    if (slideFilter && !slideFilter.includes(num)) continue;
    const entry = zip.getEntry(slideEntries[i]);
    if (!entry) continue;
    const parsed = xmlParser.parse(entry.getData().toString("utf-8")) as Record<string, unknown>;
    const spTree = ((parsed["p:sld"] as Record<string, unknown>)?.["p:cSld"] as Record<string, unknown>)?.["p:spTree"] as Record<string, unknown>;
    let title: string | undefined;
    const bodyLines: string[] = [];

    for (const sp of ensureArray(spTree?.["p:sp"])) {
      const spObj = sp as Record<string, unknown>;
      const ph = ((spObj["p:nvSpPr"] as Record<string, unknown>)?.["p:nvPr"] as Record<string, unknown>)?.["p:ph"] as Record<string, unknown>;
      const phType = ph?.["@_type"] as string | undefined;
      const isTitle = phType === "title" || phType === "ctrTitle" || phType === "subTitle";
      const txBody = spObj["p:txBody"] as Record<string, unknown>;
      if (!txBody) continue;
      const text = ensureArray(txBody["a:p"]).map(extractText).filter(Boolean).join("\n");
      if (!text) continue;
      if (isTitle && !title) title = text;
      else bodyLines.push(text);
    }

    let notes: string | undefined;
    if (args.include_notes !== false) {
      try {
        const slideFile = slideEntries[i].split("/").pop()!;
        const relsEntry = zip.getEntry(`ppt/slides/_rels/${slideFile}.rels`);
        if (relsEntry) {
          interface RelItem2 { "@_Type"?: string; "@_Target"?: string; }
          const r = xmlParser.parse(relsEntry.getData().toString("utf-8")) as Record<string, unknown>;
          const notesRel = ensureArray((r["Relationships"] as Record<string, unknown>)?.["Relationship"] as RelItem2 | RelItem2[])
            .find((rel) => rel["@_Type"]?.includes("notesSlide"));
          if (notesRel?.["@_Target"]) {
            const t = notesRel["@_Target"];
            const notesPath = t.startsWith("../") ? `ppt/${t.replace("../","")}` : `ppt/slides/${t}`;
            const notesEntry = zip.getEntry(notesPath);
            if (notesEntry) {
              const np = xmlParser.parse(notesEntry.getData().toString("utf-8")) as Record<string, unknown>;
              const nSpTree = ((np["p:notes"] as Record<string, unknown>)?.["p:cSld"] as Record<string, unknown>)?.["p:spTree"] as Record<string, unknown>;
              notes = ensureArray(nSpTree?.["p:sp"]).flatMap((sp) => {
                const txBody = (sp as Record<string, unknown>)["p:txBody"] as Record<string, unknown>;
                const ph = (((sp as Record<string, unknown>)["p:nvSpPr"] as Record<string, unknown>)?.["p:nvPr"] as Record<string, unknown>)?.["p:ph"] as Record<string, unknown>;
                if (ph?.["@_type"] === "sldImg" || !txBody) return [];
                return ensureArray(txBody["a:p"]).map(extractText).filter(Boolean);
              }).join("\n");
            }
          }
        }
      } catch { /* ignore */ }
      notes = notes ?? "";
    }

    slides.push({ index: num, title, content: bodyLines.join("\n"), notes });
  }

  // Metadata
  let meta: PptResult["metadata"] = { slide_count: slideEntries.length };
  try {
    const core = zip.getEntry("docProps/core.xml");
    if (core) {
      const c = xmlParser.parse(core.getData().toString("utf-8")) as Record<string, unknown>;
      const props = (c["cp:coreProperties"] ?? {}) as Record<string, unknown>;
      meta = { ...meta, title: String(props["dc:title"] ?? "") || undefined, author: String(props["dc:creator"] ?? "") || undefined };
    }
  } catch { /* ignore */ }

  return { metadata: meta, slides };
}

// ── Excel ─────────────────────────────────────────────────────────────────────

export interface ExcelResult {
  metadata: { sheet_names: string[]; file_size_bytes: number };
  sheets: Record<string, { headers?: string[]; rows: (string | number | boolean | null)[][]; row_count: number; col_count: number }>;
  error?: string;
}

export async function readExcel(args: {
  file_path: string;
  sheets?: string[];
  max_rows?: number;
  start_row?: number;
  end_row?: number;
  columns?: string[];
  include_formulas?: boolean;
  header_row?: boolean;
}): Promise<ExcelResult> {
  const { ok, path: p, error } = validateFile(args.file_path, ".xlsx");
  if (!ok) return { metadata: { sheet_names: [], file_size_bytes: 0 }, sheets: {}, error };

  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.readFile(p, { cellFormula: args.include_formulas, cellDates: true, cellStyles: false });
  } catch (e) {
    const msg = (e as Error).message ?? "";
    if (msg.toLowerCase().includes("password") || msg.toLowerCase().includes("encrypt"))
      return { metadata: { sheet_names: [], file_size_bytes: 0 }, sheets: {}, error: "Workbook is password-protected." };
    return { metadata: { sheet_names: [], file_size_bytes: 0 }, sheets: {}, error: `Corrupt or invalid .xlsx: ${msg}` };
  }

  const allNames = wb.SheetNames;
  const target = args.sheets ?? allNames;
  const notFound = target.filter((s) => !allNames.includes(s));
  if (notFound.length) return { metadata: { sheet_names: allNames, file_size_bytes: statSync(p).size }, sheets: {}, error: `Sheets not found: ${notFound.join(", ")}. Available: ${allNames.join(", ")}` };

  const maxRows = args.max_rows ?? 1000;
  const hasHeader = args.header_row !== false;
  const sheets: ExcelResult["sheets"] = {};

  for (const name of target) {
    const ws = wb.Sheets[name];
    if (!ws?.["!ref"]) { sheets[name] = { rows: [], row_count: 0, col_count: 0 }; continue; }
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const totalDataRows = range.e.r - range.s.r + 1 - (hasHeader ? 1 : 0);

    // Determine row window (1-indexed data rows, excluding header)
    const startDataRow = Math.max(1, args.start_row ?? 1);
    const endDataRow = Math.min(totalDataRows, args.end_row ?? Math.min(totalDataRows, startDataRow + maxRows - 1));
    const headerOffset = hasHeader ? 1 : 0;

    const numCols = range.e.c - range.s.c + 1;
    const allRows: (string | number | boolean | null)[][] = [];

    // Always read header row first if present
    const headerStart = range.s.r;
    const dataStart = range.s.r + headerOffset + (startDataRow - 1);
    const dataEnd = range.s.r + headerOffset + endDataRow - 1;

    const rowsToRead = hasHeader
      ? [headerStart, ...Array.from({ length: dataEnd - dataStart + 1 }, (_, i) => dataStart + i)]
      : Array.from({ length: dataEnd - dataStart + 1 }, (_, i) => dataStart + i);

    for (const r of rowsToRead) {
      const row: (string | number | boolean | null)[] = [];
      for (let c = range.s.c; c < range.s.c + numCols; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
        if (!cell || cell.v == null) { row.push(null); continue; }
        if (args.include_formulas && cell.f) { row.push(`=${cell.f}`); continue; }
        if (cell.t === "d" && cell.v instanceof Date) { row.push(cell.v.toISOString()); continue; }
        row.push(cell.v as string | number | boolean);
      }
      allRows.push(row);
    }

    if (hasHeader && allRows.length > 0) {
      const allHeaders = allRows[0].map((h) => (h != null ? String(h) : ""));
      let dataRows = allRows.slice(1);

      // Apply column filter if specified
      if (args.columns && args.columns.length > 0) {
        const colIndices = args.columns.map((col) => allHeaders.indexOf(col)).filter((i) => i >= 0);
        const filteredHeaders = colIndices.map((i) => allHeaders[i]);
        dataRows = dataRows.map((row) => colIndices.map((i) => row[i] ?? null));
        sheets[name] = { headers: filteredHeaders, rows: dataRows, row_count: dataRows.length, col_count: filteredHeaders.length };
      } else {
        sheets[name] = { headers: allHeaders, rows: dataRows, row_count: dataRows.length, col_count: numCols };
      }
    } else {
      sheets[name] = { rows: allRows, row_count: allRows.length, col_count: numCols };
    }
  }

  return { metadata: { sheet_names: allNames, file_size_bytes: statSync(p).size }, sheets };
}

// ── PDF ───────────────────────────────────────────────────────────────────────

export interface PdfResult {
  metadata?: { title?: string; author?: string; page_count: number; file_size_bytes: number; pdf_version?: string; is_scanned: boolean; encrypted: boolean };
  pages: { index: number; content: string; word_count: number; ocr_used: boolean }[];
  tables?: { page: number; index: number; rows: (string | null)[][] }[];
  error?: string;
  warning?: string;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

export async function readPdf(args: {
  file_path: string;
  pages?: number[];
  output_format?: "markdown" | "plain_text" | "json";
  extract_tables?: boolean;
  ocr_fallback?: boolean;
  include_metadata?: boolean;
}): Promise<PdfResult> {
  const { ok, path: p, error } = validateFile(args.file_path, ".pdf");
  if (!ok) return { pages: [], error };

  // Check pdfplumber availability
  if (args.extract_tables) {
    try { await execFileAsync("python3", ["-c", "import pdfplumber"]); }
    catch {
      try { await execFileAsync("python", ["-c", "import pdfplumber"]); }
      catch { return { pages: [], error: "pdfplumber not installed. Run: pip install pdfplumber" }; }
    }
  }

  const buf = readFileSync(p);
  const pageTexts: string[] = [];
  let pdfData: Awaited<ReturnType<typeof pdfParse>>;

  try {
    pdfData = await pdfParse(buf, {
      pagerender: (pd: { getTextContent: () => Promise<{ items: { str: string }[] }> }) =>
        pd.getTextContent().then((c) => { pageTexts.push(c.items.map((i) => i.str).join(" ")); return ""; }),
    });
  } catch (e) {
    const msg = (e as Error).message ?? "";
    if (msg.toLowerCase().includes("password") || msg.toLowerCase().includes("encrypt"))
      return { pages: [], error: "PDF is password-protected. Provide the password to decrypt." };
    return { pages: [], error: `Failed to parse PDF: ${msg}` };
  }

  const totalPages = pdfData.numpages;
  const isScanned = pdfData.text.length / Math.max(totalPages, 1) < 50;
  const pageFilter = args.pages;

  if (pageFilter) {
    const bad = pageFilter.filter((p) => p < 1 || p > totalPages);
    if (bad.length) return { pages: [], error: `Pages out of range: ${bad.join(", ")} (total: ${totalPages})` };
  }

  const toProcess = pageFilter ?? Array.from({ length: totalPages }, (_, i) => i + 1);
  const resultPages: PdfResult["pages"] = [];

  if (isScanned) {
    if (!args.ocr_fallback && args.ocr_fallback !== undefined)
      return { pages: [], warning: "Scanned PDF detected with no extractable text. Enable ocr_fallback: true to use Tesseract OCR." };
    // OCR via tesseract.js + pdf2pic
    try {
      const { fromPath } = await import("pdf2pic");
      const { createWorker } = await import("tesseract.js");
      const BATCH = 4;
      for (let i = 0; i < toProcess.length; i += BATCH) {
        const batch = toProcess.slice(i, i + BATCH);
        const results = await Promise.all(batch.map(async (pageNum) => {
          try {
            const convert = fromPath(p, { density: 300, format: "png", width: 2480, height: 3508, saveFilename: `ocr_${Date.now()}_${pageNum}`, savePath: "/tmp" });
            const img = await convert(pageNum);
            if (!img?.path) return { index: pageNum, content: "", word_count: 0, ocr_used: true };
            const worker = await createWorker("eng");
            const { data: { text } } = await worker.recognize(img.path);
            await worker.terminate();
            try { (await import("fs")).unlinkSync(img.path); } catch { /* ignore */ }
            return { index: pageNum, content: text, word_count: text.trim().split(/\s+/).filter(Boolean).length, ocr_used: true };
          } catch { return { index: pageNum, content: "", word_count: 0, ocr_used: true }; }
        }));
        resultPages.push(...results);
      }
    } catch (e) {
      return { pages: [], error: `OCR failed: ${(e as Error).message}. Ensure tesseract.js and pdf2pic are installed.` };
    }
  } else {
    for (const pageNum of toProcess) {
      const text = pageTexts[pageNum - 1] ?? "";
      resultPages.push({ index: pageNum, content: text, word_count: text.trim().split(/\s+/).filter(Boolean).length, ocr_used: false });
    }
  }

  // Tables via pdfplumber
  let tables: PdfResult["tables"] | undefined;
  if (args.extract_tables) {
    try {
      const scriptPath = resolve(__dirname, "../../../../mcp-office-reader/scripts/extract_pdf_tables.py");
      let stdout: string;
      try { ({ stdout } = await execFileAsync("python3", [scriptPath, p])); }
      catch { ({ stdout } = await execFileAsync("python", [scriptPath, p])); }
      const raw = JSON.parse(stdout) as { page: number; rows: (string | null)[][] }[];
      tables = raw.map((t, idx) => ({ ...t, index: idx }));
      if (pageFilter) tables = tables.filter((t) => pageFilter.includes(t.page));
    } catch { tables = []; }
  }

  const pdfInfo = pdfData.info as Record<string, unknown>;
  const metadata = args.include_metadata !== false ? {
    title: String(pdfInfo?.["Title"] ?? "") || undefined,
    author: String(pdfInfo?.["Author"] ?? "") || undefined,
    page_count: totalPages,
    file_size_bytes: statSync(p).size,
    pdf_version: pdfData.version || undefined,
    is_scanned: isScanned,
    encrypted: false,
  } : undefined;

  return { metadata, pages: resultPages, tables };
}

// ── List files ────────────────────────────────────────────────────────────────

export interface ListFilesResult {
  directory: string;
  files: { path: string; name: string; extension: string; size_bytes: number; modified?: string }[];
  total_count: number;
  error?: string;
}

export async function listOfficeFiles(args: {
  directory: string;
  recursive?: boolean;
  file_types?: string[];
}): Promise<ListFilesResult> {
  const dir = resolve(args.directory);
  if (!existsSync(dir)) return { directory: dir, files: [], total_count: 0, error: `Directory not found: ${dir}` };
  if (!statSync(dir).isDirectory()) return { directory: dir, files: [], total_count: 0, error: `Not a directory: ${dir}` };

  const exts = (args.file_types ?? ["docx", "pptx", "xlsx", "pdf"]).map((t) => `.${t.toLowerCase()}`);
  const files: ListFilesResult["files"] = [];

  function scan(d: string) {
    for (const entry of readdirSync(d, { withFileTypes: true })) {
      const full = join(d, entry.name);
      if (entry.isDirectory() && args.recursive) { scan(full); continue; }
      if (!entry.isFile()) continue;
      const ext = extname(entry.name).toLowerCase();
      if (!exts.includes(ext)) continue;
      const s = statSync(full);
      files.push({ path: full, name: basename(entry.name), extension: ext.slice(1), size_bytes: s.size, modified: s.mtime.toISOString() });
    }
  }

  try { scan(dir); } catch (e) { return { directory: dir, files: [], total_count: 0, error: (e as Error).message }; }
  return { directory: dir, files, total_count: files.length };
}
