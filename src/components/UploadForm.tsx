"use client";

import { useState, useRef, useCallback } from "react";
import { Upload, FileSpreadsheet, Loader2, CheckCircle, AlertCircle } from "lucide-react";

interface Props {
  onSuccess: () => void;
}

export default function UploadForm({ onSuccess }: Props) {
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState<"idle" | "uploading" | "success" | "error">("idle");
  const [errorMsg, setErrorMsg] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = (f: File) => {
    if (!f.name.match(/\.xlsx?$/i)) {
      setErrorMsg("Please upload an Excel file (.xlsx or .xls)");
      setStatus("error");
      return;
    }
    setFile(f);
    setStatus("idle");
    setErrorMsg("");
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  }, []);

  const handleGenerate = async () => {
    if (!file) return;
    setStatus("uploading");
    setErrorMsg("");

    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch("/api/generate-report", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const json = await res.json().catch(() => ({}));
        throw new Error(json.error ?? "Failed to generate report");
      }

      // Trigger download
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const disposition = res.headers.get("Content-Disposition") ?? "";
      const match = disposition.match(/filename="(.+?)"/);
      a.download = match?.[1] ?? "25N-Report.pdf";
      a.click();
      URL.revokeObjectURL(url);

      setStatus("success");
      onSuccess();
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : "Something went wrong");
      setStatus("error");
    }
  };

  const isLoading = status === "uploading";

  return (
    <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
      <h2 className="text-lg font-semibold text-[#0F2044] mb-4">Generate New Report</h2>

      {/* Drop zone */}
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => !isLoading && inputRef.current?.click()}
        className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
          dragging ? "border-[#0F2044] bg-blue-50" : "border-gray-200 hover:border-[#C9A84C] hover:bg-amber-50"
        } ${isLoading ? "pointer-events-none opacity-60" : ""}`}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }}
        />

        {file ? (
          <div className="flex flex-col items-center gap-2">
            <FileSpreadsheet className="w-10 h-10 text-[#0F2044]" />
            <p className="font-medium text-[#0F2044]">{file.name}</p>
            <p className="text-sm text-gray-400">{(file.size / 1024).toFixed(0)} KB — click to change</p>
          </div>
        ) : (
          <div className="flex flex-col items-center gap-2">
            <Upload className="w-10 h-10 text-gray-300" />
            <p className="font-medium text-gray-600">Drop your Excel file here</p>
            <p className="text-sm text-gray-400">or click to browse (.xlsx)</p>
          </div>
        )}
      </div>

      {/* Status messages */}
      {status === "error" && (
        <div className="mt-3 flex items-center gap-2 text-red-600 text-sm">
          <AlertCircle className="w-4 h-4 flex-shrink-0" />
          <span>{errorMsg}</span>
        </div>
      )}
      {status === "success" && (
        <div className="mt-3 flex items-center gap-2 text-green-600 text-sm">
          <CheckCircle className="w-4 h-4 flex-shrink-0" />
          <span>Report generated and downloaded successfully.</span>
        </div>
      )}

      {/* Generate button */}
      <button
        onClick={handleGenerate}
        disabled={!file || isLoading}
        className="mt-4 w-full bg-[#0F2044] hover:bg-[#1A3461] disabled:opacity-40 text-white font-semibold py-3 px-6 rounded-lg transition flex items-center justify-center gap-2"
      >
        {isLoading ? (
          <>
            <Loader2 className="w-4 h-4 animate-spin" />
            Generating report…
          </>
        ) : (
          "Generate PDF Report"
        )}
      </button>
    </div>
  );
}
