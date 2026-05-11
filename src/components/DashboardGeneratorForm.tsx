"use client";

import { useState, useRef, useCallback } from "react";
import { Upload, FileSpreadsheet, Loader2, CheckCircle, AlertCircle, TableProperties } from "lucide-react";

interface Props {
  onSuccess?: () => void;
}

export default function DashboardGeneratorForm({ onSuccess }: Props) {
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

      const res = await fetch("/api/generate-dashboard", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const json = await res.json().catch(() => ({}));
        throw new Error(json.error ?? "Failed to generate dashboard");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const disposition = res.headers.get("Content-Disposition") ?? "";
      const match = disposition.match(/filename="(.+?)"/);
      a.download = match?.[1] ?? "Dashboard - Generated.xlsx";
      a.click();
      URL.revokeObjectURL(url);

      setStatus("success");
      onSuccess?.();
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : "Something went wrong");
      setStatus("error");
    }
  };

  const isLoading = status === "uploading";

  return (
    <div className="bg-white rounded-xl border border-[#C9A84C] shadow-sm p-6">
      <div className="flex items-center gap-2 mb-1">
        <TableProperties className="w-5 h-5 text-[#C9A84C]" />
        <h2 className="text-lg font-semibold text-[#0F2044]">Generate Analysis Dashboard</h2>
      </div>
      <p className="text-sm text-gray-500 mb-4">
        Upload your raw data workbook. The tool will auto-populate the{" "}
        <span className="font-medium text-[#0F2044]">P.1</span>,{" "}
        <span className="font-medium text-[#0F2044]">P.2</span>, and{" "}
        <span className="font-medium text-[#0F2044]">P.4</span> dashboard tabs — then download
        the completed Excel for review before generating the PDF.
      </p>

      {/* Drop zone */}
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => !isLoading && inputRef.current?.click()}
        className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
          dragging
            ? "border-[#C9A84C] bg-amber-50"
            : "border-gray-200 hover:border-[#C9A84C] hover:bg-amber-50"
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
            <p className="font-medium text-gray-600">Drop your raw data workbook here</p>
            <p className="text-sm text-gray-400">or click to browse (.xlsx)</p>
          </div>
        )}
      </div>

      {status === "error" && (
        <div className="mt-3 flex items-center gap-2 text-red-600 text-sm">
          <AlertCircle className="w-4 h-4 flex-shrink-0" />
          <span>{errorMsg}</span>
        </div>
      )}

      {status === "success" && (
        <div className="mt-3 rounded-lg bg-green-50 border border-green-200 p-3 text-sm text-green-800">
          <div className="flex items-center gap-2 font-medium mb-1">
            <CheckCircle className="w-4 h-4" />
            Dashboard generated and downloaded.
          </div>
          <p className="text-green-700 pl-6">
            Open the Excel file and review <strong>P.1</strong>, <strong>P.2</strong>, and{" "}
            <strong>P.4</strong>. Once approved, upload the file below to generate the PDF report.
          </p>
        </div>
      )}

      <button
        onClick={handleGenerate}
        disabled={!file || isLoading}
        className="mt-4 w-full bg-[#C9A84C] hover:bg-[#b8973d] disabled:opacity-40 text-white font-semibold py-3 px-6 rounded-lg transition flex items-center justify-center gap-2"
      >
        {isLoading ? (
          <>
            <Loader2 className="w-4 h-4 animate-spin" />
            Generating dashboard…
          </>
        ) : (
          "Generate Dashboard Excel"
        )}
      </button>
    </div>
  );
}
