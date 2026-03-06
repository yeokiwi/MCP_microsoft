import { createWorker } from "tesseract.js";
import { fromPath } from "pdf2pic";
import { existsSync, unlinkSync } from "fs";
import { join, dirname } from "path";
import { tmpdir } from "os";

const OCR_CONCURRENCY = 4;

export interface OcrPageResult {
  pageIndex: number;
  text: string;
  success: boolean;
  error?: string;
}

export async function convertPdfPageToImage(
  filePath: string,
  pageNumber: number,
  outputDir?: string
): Promise<string | null> {
  const dir = outputDir || tmpdir();
  try {
    const convert = fromPath(filePath, {
      density: 300,
      format: "png",
      width: 2480,
      height: 3508,
      saveFilename: `ocr_page_${Date.now()}`,
      savePath: dir,
    });

    const result = await convert(pageNumber);
    if (result && result.path && existsSync(result.path)) {
      return result.path;
    }
    return null;
  } catch (error) {
    return null;
  }
}

export async function runOcrOnImage(imagePath: string): Promise<string> {
  const worker = await createWorker("eng");
  try {
    const {
      data: { text },
    } = await worker.recognize(imagePath);
    return text;
  } finally {
    await worker.terminate();
    // Clean up temp image
    try {
      if (existsSync(imagePath)) {
        unlinkSync(imagePath);
      }
    } catch {
      // ignore cleanup errors
    }
  }
}

export async function ocrPdfPages(
  filePath: string,
  pageNumbers: number[]
): Promise<OcrPageResult[]> {
  const results: OcrPageResult[] = [];

  // Process in batches to respect concurrency limit
  for (let i = 0; i < pageNumbers.length; i += OCR_CONCURRENCY) {
    const batch = pageNumbers.slice(i, i + OCR_CONCURRENCY);

    const batchResults = await Promise.all(
      batch.map(async (pageNum): Promise<OcrPageResult> => {
        try {
          const imagePath = await convertPdfPageToImage(filePath, pageNum);
          if (!imagePath) {
            return {
              pageIndex: pageNum,
              text: "",
              success: false,
              error: "Failed to convert PDF page to image",
            };
          }

          const text = await runOcrOnImage(imagePath);
          return { pageIndex: pageNum, text, success: true };
        } catch (error) {
          return {
            pageIndex: pageNum,
            text: "",
            success: false,
            error: error instanceof Error ? error.message : String(error),
          };
        }
      })
    );

    results.push(...batchResults);
  }

  return results;
}

export function isScannedPdf(
  extractedText: string,
  pageCount: number
): boolean {
  if (pageCount === 0) return true;
  const avgCharsPerPage = extractedText.length / pageCount;
  return avgCharsPerPage < 50;
}
