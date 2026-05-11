"use client";

import { useEffect, useState, useCallback } from "react";
import { Download, FileText, Clock } from "lucide-react";

interface Report {
  id: string;
  filename: string;
  report_date: string;
  created_at: string;
  downloadUrl: string | null;
}

interface Props {
  refreshTrigger: number;
}

export default function ReportHistory({ refreshTrigger }: Props) {
  const [reports, setReports] = useState<Report[]>([]);
  const [loading, setLoading] = useState(true);

  const fetchReports = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/reports");
      const json = await res.json();
      setReports(json.reports ?? []);
    } catch {
      // silent
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchReports();
  }, [fetchReports, refreshTrigger]);

  if (loading) {
    return (
      <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
        <h2 className="text-lg font-semibold text-[#0F2044] mb-4">Report History</h2>
        <div className="flex items-center justify-center py-8">
          <div className="w-6 h-6 border-2 border-[#0F2044] border-t-transparent rounded-full animate-spin" />
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
      <h2 className="text-lg font-semibold text-[#0F2044] mb-4">Report History</h2>

      {reports.length === 0 ? (
        <div className="text-center py-12 text-gray-400">
          <FileText className="w-10 h-10 mx-auto mb-3 opacity-30" />
          <p className="text-sm">No reports yet — upload your first Excel file above.</p>
        </div>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100">
                <th className="text-left py-2 pr-4 text-xs font-semibold text-gray-400 uppercase tracking-wide">File</th>
                <th className="text-left py-2 pr-4 text-xs font-semibold text-gray-400 uppercase tracking-wide">Report Date</th>
                <th className="text-left py-2 pr-4 text-xs font-semibold text-gray-400 uppercase tracking-wide">Generated</th>
                <th className="text-right py-2 text-xs font-semibold text-gray-400 uppercase tracking-wide">Download</th>
              </tr>
            </thead>
            <tbody>
              {reports.map((report) => (
                <tr key={report.id} className="border-b border-gray-50 hover:bg-gray-50 transition">
                  <td className="py-3 pr-4">
                    <div className="flex items-center gap-2">
                      <FileText className="w-4 h-4 text-[#0F2044] flex-shrink-0" />
                      <span className="text-gray-700 truncate max-w-[200px]">{report.filename}</span>
                    </div>
                  </td>
                  <td className="py-3 pr-4 text-gray-600">
                    {new Date(report.report_date).toLocaleDateString("en-US", {
                      month: "short",
                      day: "numeric",
                      year: "numeric",
                    })}
                  </td>
                  <td className="py-3 pr-4 text-gray-400">
                    <div className="flex items-center gap-1">
                      <Clock className="w-3 h-3" />
                      {new Date(report.created_at).toLocaleString("en-US", {
                        month: "short",
                        day: "numeric",
                        hour: "numeric",
                        minute: "2-digit",
                      })}
                    </div>
                  </td>
                  <td className="py-3 text-right">
                    {report.downloadUrl ? (
                      <a
                        href={report.downloadUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-1 text-[#0F2044] hover:text-[#C9A84C] font-medium transition"
                      >
                        <Download className="w-4 h-4" />
                        PDF
                      </a>
                    ) : (
                      <span className="text-gray-300 text-xs">Expired</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
