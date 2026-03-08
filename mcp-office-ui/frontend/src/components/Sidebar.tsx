import { useState, useCallback, useMemo } from "react";
import type { OfficeFile } from "../types.js";
import { FILE_ICONS } from "../types.js";

type SortField = "date" | "size" | "type";
type SortDir = "asc" | "desc";

interface Props {
  selectedPaths: Set<string>;
  onToggleFile: (path: string) => void;
  onSelectAll: (paths: string[]) => void;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(iso?: string): string {
  if (!iso) return "—";
  return new Date(iso).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

export function Sidebar({ selectedPaths, onToggleFile, onSelectAll }: Props) {
  const [directory, setDirectory] = useState("");
  const [recursive, setRecursive] = useState(false);
  const [files, setFiles] = useState<OfficeFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [sortField, setSortField] = useState<SortField>("type");
  const [sortDir, setSortDir] = useState<SortDir>("asc");

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

  const sortedFiles = useMemo(() => {
    const sorted = [...files].sort((a, b) => {
      let cmp = 0;
      if (sortField === "date") {
        cmp = (a.modified ?? "").localeCompare(b.modified ?? "");
      } else if (sortField === "size") {
        cmp = a.size_bytes - b.size_bytes;
      } else {
        // type: sort by extension then name
        cmp = a.extension.localeCompare(b.extension) || a.name.localeCompare(b.name);
      }
      return sortDir === "asc" ? cmp : -cmp;
    });
    return sorted;
  }, [files, sortField, sortDir]);

  const handleSortField = (field: SortField) => {
    if (field === sortField) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortField(field);
      setSortDir("asc");
    }
  };

  const allPaths = sortedFiles.map((f) => f.path);
  const allSelected = allPaths.length > 0 && allPaths.every((p) => selectedPaths.has(p));

  const handleToggleAll = () => {
    onSelectAll(allSelected ? [] : allPaths);
  };

  const grouped = useMemo(() => {
    if (sortField === "type") {
      // Group by extension when sorting by type
      return sortedFiles.reduce<Record<string, OfficeFile[]>>((acc, f) => {
        acc[f.extension] = acc[f.extension] ?? [];
        acc[f.extension].push(f);
        return acc;
      }, {});
    }
    // For date/size sorts, show a flat "All files" group to preserve order
    return sortedFiles.length > 0 ? { "": sortedFiles } : {};
  }, [sortedFiles, sortField]);

  const sortLabel = (field: SortField) => {
    const labels: Record<SortField, string> = { date: "Date", size: "Size", type: "Type" };
    const active = field === sortField;
    return `${labels[field]}${active ? (sortDir === "asc" ? " ↑" : " ↓") : ""}`;
  };

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

        {files.length > 0 && (
          <>
            {/* Sort bar */}
            <div className="sort-bar">
              <span className="sort-bar-label">Sort:</span>
              {(["type", "date", "size"] as SortField[]).map((f) => (
                <button
                  key={f}
                  className={`sort-btn${sortField === f ? " sort-btn--active" : ""}`}
                  onClick={() => handleSortField(f)}
                >
                  {sortLabel(f)}
                </button>
              ))}
            </div>

            {/* Select all bar */}
            <div className="sidebar-select-bar">
              <button className="btn-link" onClick={handleToggleAll}>
                {allSelected ? "Deselect all" : "Select all"} ({files.length})
              </button>
            </div>
          </>
        )}

        {Object.entries(grouped).map(([ext, extFiles]) => (
          <div key={ext} className="file-group">
            {ext && (
              <div className="file-group-header">
                {FILE_ICONS[ext] ?? "📄"} .{ext.toUpperCase()} ({extFiles.length})
              </div>
            )}
            {extFiles.map((f) => {
              const isSelected = selectedPaths.has(f.path);
              return (
                <button
                  key={f.path}
                  className={`file-item${isSelected ? " file-item--selected" : ""}`}
                  onClick={() => onToggleFile(f.path)}
                  title={f.path}
                >
                  <span className="file-item-check">{isSelected ? "☑" : "☐"}</span>
                  <span className="file-name">{f.name}</span>
                  <span className="file-meta">
                    {sortField === "date" ? formatDate(f.modified) : formatSize(f.size_bytes)}
                  </span>
                </button>
              );
            })}
          </div>
        ))}
      </div>
    </aside>
  );
}
