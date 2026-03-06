import { existsSync, statSync, readdirSync } from "fs";
import { resolve, extname, basename, join } from "path";
import type { OfficeFileInfo } from "../types.js";

export const SUPPORTED_EXTENSIONS = [".docx", ".pptx", ".xlsx", ".pdf"] as const;
export type SupportedExtension = (typeof SUPPORTED_EXTENSIONS)[number];

export function resolvePath(filePath: string): string {
  return resolve(filePath);
}

export function validateFilePath(
  filePath: string,
  expectedExtension?: string
): { valid: boolean; error?: string; resolvedPath?: string } {
  const resolvedPath = resolvePath(filePath);

  if (!existsSync(resolvedPath)) {
    return { valid: false, error: `File not found: ${resolvedPath}` };
  }

  const stat = statSync(resolvedPath);
  if (!stat.isFile()) {
    return { valid: false, error: `Path is not a file: ${resolvedPath}` };
  }

  if (stat.size === 0) {
    return { valid: false, error: `File is empty: ${resolvedPath}` };
  }

  const ext = extname(resolvedPath).toLowerCase();

  if (expectedExtension && ext !== expectedExtension.toLowerCase()) {
    return {
      valid: false,
      error: `File extension mismatch: expected ${expectedExtension}, got ${ext}`,
    };
  }

  return { valid: true, resolvedPath };
}

export function getFileSize(filePath: string): number {
  try {
    return statSync(filePath).size;
  } catch {
    return 0;
  }
}

export function getFileModified(filePath: string): string | undefined {
  try {
    return statSync(filePath).mtime.toISOString();
  } catch {
    return undefined;
  }
}

export function listOfficeFilesInDirectory(
  directory: string,
  recursive: boolean,
  fileTypes: string[]
): OfficeFileInfo[] {
  const resolvedDir = resolvePath(directory);

  if (!existsSync(resolvedDir)) {
    throw new Error(`Directory not found: ${resolvedDir}`);
  }

  const stat = statSync(resolvedDir);
  if (!stat.isDirectory()) {
    throw new Error(`Path is not a directory: ${resolvedDir}`);
  }

  const results: OfficeFileInfo[] = [];
  const allowedExts = fileTypes.map((t) => `.${t.toLowerCase()}`);

  function scanDir(dirPath: string): void {
    const entries = readdirSync(dirPath, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = join(dirPath, entry.name);
      if (entry.isDirectory() && recursive) {
        scanDir(fullPath);
      } else if (entry.isFile()) {
        const ext = extname(entry.name).toLowerCase();
        if (allowedExts.includes(ext)) {
          const fileStat = statSync(fullPath);
          results.push({
            path: fullPath,
            name: basename(entry.name),
            extension: ext.slice(1),
            size_bytes: fileStat.size,
            modified: fileStat.mtime.toISOString(),
          });
        }
      }
    }
  }

  scanDir(resolvedDir);
  return results;
}
