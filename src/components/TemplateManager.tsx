"use client";

import { useState, useEffect, useRef } from "react";
import { Upload, CheckCircle, AlertCircle, Trash2 } from "lucide-react";

interface TemplateStatus {
  exists:        boolean;
  size?:         number | null;
  updatedAt?:    string | null;
  originalName?: string | null;
  sheetCount?:   number | null;
}

function fmtDate(iso: string) {
  return new Date(iso).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

export default function TemplateManager() {
  const [status,    setStatus]    = useState<TemplateStatus | null>(null);
  const [loading,   setLoading]   = useState(true);
  const [uploading, setUploading] = useState(false);
  const [deleting,  setDeleting]  = useState(false);
  const [msg,       setMsg]       = useState<{ ok: boolean; text: string } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  async function fetchStatus() {
    setLoading(true);
    try {
      const res = await fetch("/api/template");
      if (res.ok) setStatus(await res.json());
    } catch { /* ignore */ }
    finally { setLoading(false); }
  }
  useEffect(() => { fetchStatus(); }, []);

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true); setMsg(null);
    const fd = new FormData(); fd.append("file", file);
    try {
      const res = await fetch("/api/template", { method: "POST", body: fd });
      const data = await res.json();
      setMsg(res.ok
        ? { ok: true,  text: "Template saved." }
        : { ok: false, text: data.error ?? "Upload failed" });
      if (res.ok) fetchStatus();
    } catch {
      setMsg({ ok: false, text: "Upload failed" });
    } finally {
      setUploading(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  async function handleDelete() {
    if (!confirm("Remove the master template?")) return;
    setDeleting(true); setMsg(null);
    try {
      const res = await fetch("/api/template", { method: "DELETE" });
      if (res.ok) { setMsg({ ok: true, text: "Template removed." }); fetchStatus(); }
      else { const d = await res.json(); setMsg({ ok: false, text: d.error ?? "Remove failed" }); }
    } catch {
      setMsg({ ok: false, text: "Remove failed" });
    } finally { setDeleting(false); }
  }

  if (loading) return null;

  // Template exists — show a compact confirmation strip
  if (status?.exists) {
    return (
      <div className="flex items-center gap-2.5 px-4 py-2.5 rounded-lg bg-emerald-50 border border-emerald-200 text-[11px]">
        <CheckCircle className="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" />
        <span className="text-emerald-800 font-medium">Master template active</span>
        {status.originalName && (
          <span className="text-emerald-600 truncate max-w-[200px]">{status.originalName}</span>
        )}
        {status.updatedAt && (
          <span className="text-emerald-500 ml-auto flex-shrink-0">{fmtDate(status.updatedAt)}</span>
        )}

        {msg && (
          <span className={`ml-2 ${msg.ok ? "text-emerald-700" : "text-red-600"}`}>{msg.text}</span>
        )}

        <div className="flex items-center gap-1 ml-2 flex-shrink-0">
          <label className={`flex items-center gap-1 px-2 py-1 rounded text-[10px] font-medium cursor-pointer transition border
            ${uploading ? "opacity-50 cursor-not-allowed border-transparent" : "border-emerald-300 text-emerald-700 hover:bg-emerald-100"}`}
          >
            <Upload className="w-2.5 h-2.5" />
            {uploading ? "…" : "Replace"}
            <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" disabled={uploading} onChange={handleUpload} />
          </label>
          <button
            onClick={handleDelete}
            disabled={deleting}
            className="flex items-center gap-1 px-2 py-1 rounded text-[10px] font-medium border border-transparent text-emerald-500 hover:text-red-500 hover:border-red-200 transition disabled:opacity-50"
          >
            <Trash2 className="w-2.5 h-2.5" />
            {deleting ? "…" : "Remove"}
          </button>
        </div>
      </div>
    );
  }

  // No template — show an amber prompt to upload
  return (
    <div className="flex items-center gap-2.5 px-4 py-2.5 rounded-lg bg-amber-50 border border-amber-200 text-[11px]">
      <AlertCircle className="w-3.5 h-3.5 text-amber-500 flex-shrink-0" />
      <span className="text-amber-800 font-medium">No master template</span>
      <span className="text-amber-600">Upload your production dashboard workbook once — it persists for all future months.</span>

      {msg && (
        <span className={`ml-2 ${msg.ok ? "text-emerald-700" : "text-red-600"}`}>{msg.text}</span>
      )}

      <label className={`ml-auto flex items-center gap-1.5 px-3 py-1.5 rounded text-[11px] font-medium cursor-pointer transition flex-shrink-0
        ${uploading
          ? "bg-slate-100 text-slate-400 cursor-not-allowed"
          : "bg-[#0F2044] text-white hover:bg-[#162d5e]"}`}
      >
        <Upload className="w-3 h-3" />
        {uploading ? "Uploading…" : "Upload template"}
        <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" disabled={uploading} onChange={handleUpload} />
      </label>
    </div>
  );
}
