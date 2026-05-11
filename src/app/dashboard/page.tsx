"use client";

import { useState } from "react";
import UploadForm from "@/components/UploadForm";
import DashboardGeneratorForm from "@/components/DashboardGeneratorForm";
import ReportHistory from "@/components/ReportHistory";
import ExcelHistory from "@/components/ExcelHistory";
import { LogOut } from "lucide-react";
import { useRouter } from "next/navigation";

function StepLabel({ step, label }: { step: number; label: string }) {
  return (
    <div className="flex items-center gap-3">
      <div className="flex-shrink-0 w-7 h-7 rounded-full bg-[#0F2044] text-white text-sm font-bold flex items-center justify-center">
        {step}
      </div>
      <span className="text-base font-semibold text-[#0F2044]">{label}</span>
    </div>
  );
}

export default function DashboardPage() {
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [excelRefreshTrigger, setExcelRefreshTrigger] = useState(0);
  const router = useRouter();

  const handleSignOut = async () => {
    await fetch("/api/auth/logout", { method: "POST" });
    router.push("/login");
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-[#0F2044] text-white px-6 py-4 flex items-center justify-between shadow-lg">
        <div>
          <span className="font-bold text-lg">25N Coworking</span>
          <span className="ml-2 text-[#C9A84C] text-sm font-medium">Report Generator</span>
        </div>
        <button
          onClick={handleSignOut}
          className="flex items-center gap-2 text-gray-300 hover:text-white text-sm transition"
        >
          <LogOut className="w-4 h-4" />
          Sign out
        </button>
      </nav>

      <main className="max-w-3xl mx-auto px-4 py-8 space-y-6">
        <div>
          <h1 className="text-2xl font-bold text-[#0F2044]">Occupancy Reports</h1>
          <p className="text-gray-500 text-sm mt-1">
            Two-step workflow: auto-generate the analysis dashboard from raw data, then produce the PDF for management.
          </p>
        </div>

        <div className="space-y-2">
          <StepLabel step={1} label="Generate Analysis Dashboard" />
          <DashboardGeneratorForm onSuccess={() => setExcelRefreshTrigger((n) => n + 1)} />
        </div>

        <ExcelHistory refreshTrigger={excelRefreshTrigger} />

        <div className="flex items-center gap-3 py-1">
          <div className="flex-1 border-t border-dashed border-gray-300" />
          <span className="text-xs text-gray-400 font-medium uppercase tracking-wide">then</span>
          <div className="flex-1 border-t border-dashed border-gray-300" />
        </div>

        <div className="space-y-2">
          <StepLabel step={2} label="Generate PDF Report" />
          <UploadForm onSuccess={() => setRefreshTrigger((n) => n + 1)} />
        </div>

        <ReportHistory refreshTrigger={refreshTrigger} />
      </main>
    </div>
  );
}
