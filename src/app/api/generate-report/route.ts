import { NextRequest, NextResponse } from "next/server";
import { renderToBuffer } from "@react-pdf/renderer";
import React from "react";
import { parseExcelWorkbook } from "@/lib/excel-parser";
import { ReportDocument } from "@/lib/pdf-generator";
import { createServiceClient } from "@/lib/supabase-server";
import { verifySessionToken } from "@/lib/auth";

export const dynamic = "force-dynamic";

const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20 MB

export async function POST(request: NextRequest) {
  if (!(await verifySessionToken(request.cookies.get("auth")?.value ?? ""))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }
    if (!file.name.match(/\.xlsx?$/i)) {
      return NextResponse.json({ error: "File must be an Excel file (.xlsx or .xls)" }, { status: 400 });
    }
    // WX-10: Reject oversized files before loading into memory
    if (file.size > MAX_FILE_SIZE) {
      return NextResponse.json({ error: "File too large (max 20 MB)" }, { status: 400 });
    }

    // Parse Excel
    const arrayBuffer = await file.arrayBuffer();
    let reportData;
    try {
      reportData = parseExcelWorkbook(arrayBuffer);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Failed to parse Excel file";
      return NextResponse.json({ error: msg }, { status: 400 });
    }

    // Generate PDF
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const pdfBuffer = await renderToBuffer(
      React.createElement(ReportDocument, { data: reportData }) as any
    );

    // Store in Supabase
    const supabase = createServiceClient();
    const timestamp = Date.now();
    const pdfPath = `reports/${timestamp}.pdf`;

    const { error: uploadError } = await supabase.storage
      .from("reports")
      .upload(pdfPath, pdfBuffer, { contentType: "application/pdf" });

    if (!uploadError) {
      await supabase.from("reports").insert({
        filename: file.name,
        report_date: new Date().toISOString().split("T")[0],
        pdf_path: pdfPath,
      });
    }

    // Filename uses full month name e.g. "25N-Report-August-2024.pdf"
    const safeDateStr = reportData.reportDate.replace(/\s+/g, "-").replace(/[^A-Za-z0-9\-]/g, "");
    const filename = `25N-Report-${safeDateStr}.pdf`;
    return new NextResponse(Buffer.from(pdfBuffer), {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    console.error("Report generation error:", err);
    return NextResponse.json(
      { error: "Failed to generate report. Please check the file format." },
      { status: 500 }
    );
  }
}
