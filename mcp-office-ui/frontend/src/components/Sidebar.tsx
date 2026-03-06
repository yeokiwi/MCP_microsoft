import { useState, useCallback } from "react";
import type { OfficeFile } from "../types.js";
import { FILE_ICONS } from "../types.js";

interface Props {
  onFileSelect: (path: string) => void;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function Sidebar({ onFileSelect }: Props) {
  const [directory, setDirectory] = useState("");
  const [recursive, setRecursive] = useState(false);
  const [files, setFiles] = useState<OfficeFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const browseDirectory = useCallback(async () => {
    if (!directory.trim()) return;
    setLoading(true);
    setError("");
    try {
      const params = new URLSearchParams({ directory, recursive: String(recursive) });
      const res = await fetch(`/api/files?${params}`);
      const data = await res.json() as { files?: OfficeFile[]; error?: string };
      if (data.error) { setError(data.error); setFiles([]); }
      else setFiles(data.files ?? []);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }, [directory, recursive]);

  const grouped = files.reduce<Record<string, OfficeFile[]>>((acc, f) => {
    acc[f.extension] = acc[f.extension] ?? [];
    acc[f.extension].push(f);
    return acc;
  }, {});

  return (
    <aside className="sidebar">
      <div className="sidebar-header">
        <h2 className="sidebar-title">📁 File Browser</h2>
      </div>

      <div className="sidebar-search">
        <input
          className="sidebar-input"
          type="text"
          placeholder="Directory path…"
          value={directory}
          onChange={(e) => setDirectory(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && browseDirectory()}
        />
        <label className="sidebar-checkbox">
          <input type="checkbox" checked={recursive} onChange={(e) => setRecursive(e.target.checked)} />
          <span>Recursive</span>
        </label>
        <button className="btn btn-primary sidebar-btn" onClick={browseDirectory} disabled={loading}>
          {loading ? "Scanning…" : "Scan"}
        </button>
      </div>

      {error && <p className="sidebar-error">{error}</p>}

      <div className="sidebar-files">
        {files.length === 0 && !loading && (
          <p className="sidebar-empty">No files found. Enter a directory path and click Scan.</p>
        )}

        {Object.entries(grouped).map(([ext, extFiles]) => (
          <div key={ext} className="file-group">
            <div className="file-group-header">
              {FILE_ICONS[ext] ?? "📄"} .{ext.toUpperCase()} ({extFiles.length})
            </div>
            {extFiles.map((f) => (
              <button
                key={f.path}
                className="file-item"
                onClick={() => onFileSelect(f.path)}
                title={f.path}
              >
                <span className="file-name">{f.name}</span>
                <span className="file-meta">{formatSize(f.size_bytes)}</span>
              </button>
            ))}
          </div>
        ))}
      </div>
    </aside>
  );
}
