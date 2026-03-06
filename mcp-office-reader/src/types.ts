// Shared TypeScript types for the MCP Office Reader

export interface DocumentMetadata {
  title?: string;
  author?: string;
  created?: string;
  modified?: string;
  word_count?: number;
}

export interface TableData {
  index: number;
  rows: string[][];
}

export interface WordDocumentResult {
  metadata?: DocumentMetadata;
  content: string;
  tables: TableData[];
}

export interface SlideShape {
  type: string;
  text?: string;
}

export interface Slide {
  index: number;
  title?: string;
  content: string;
  notes?: string;
  shapes?: SlideShape[];
}

export interface PresentationMetadata {
  title?: string;
  author?: string;
  slide_count: number;
}

export interface PowerPointResult {
  metadata: PresentationMetadata;
  slides: Slide[];
}

export interface SheetData {
  headers?: string[];
  rows: (string | number | boolean | null)[][];
  row_count: number;
  col_count: number;
}

export interface ExcelMetadata {
  sheet_names: string[];
  file_size_bytes: number;
}

export interface ExcelResult {
  metadata: ExcelMetadata;
  sheets: Record<string, SheetData>;
}

export interface PdfPage {
  index: number;
  content: string;
  word_count: number;
  ocr_used: boolean;
}

export interface PdfTable {
  page: number;
  index: number;
  rows: (string | null)[][];
}

export interface PdfMetadata {
  title?: string;
  author?: string;
  page_count: number;
  file_size_bytes: number;
  pdf_version?: string;
  is_scanned: boolean;
  encrypted: boolean;
}

export interface PdfResult {
  metadata?: PdfMetadata;
  pages: PdfPage[];
  tables?: PdfTable[];
  error?: string;
}

export interface OfficeFileInfo {
  path: string;
  name: string;
  extension: string;
  size_bytes: number;
  modified?: string;
}

export interface ListOfficeFilesResult {
  directory: string;
  files: OfficeFileInfo[];
  total_count: number;
}
