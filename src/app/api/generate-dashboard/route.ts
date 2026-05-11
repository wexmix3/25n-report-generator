import { NextRequest, NextResponse } from "next/server";
import { generateDashboardExcel } from "@/lib/excel-dashboard-generator";
import { verifySessionToken } from "@/lib/auth";
import { createServiceClient } from "@/lib/supabase-server";

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
      return NextResponse.json(
        { error: "File must be an Excel file (.xlsx or .xls)" },
        { status: 400 }
      );
    }
    if (file.size > MAX_FILE_SIZE) {
      return NextResponse.json({ error: "File too large (max 20 MB)" }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();

    let outputBuffer: Buffer;
    try {
      outputBuffer = generateDashboardExcel(arrayBuffer);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Failed to process Excel file";
      return NextResponse.json({ error: msg }, { status: 400 });
    }

    const baseName = file.name.replace(/\.xlsx?$/i, "");
    const filename = `${baseName} - Generated Dashboard.xlsx`;

    // Save to Supabase Storage + record in excel_downloads (best-effort, don't block download)
    try {
      const supabase = createServiceClient();
      const timestamp = Date.now();
      const excelPath = `excel-dashboards/${timestamp}.xlsx`;

      const { error: uploadError } = await supabase.storage
        .from("reports")
        .upload(excelPath, outputBuffer, {
          contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

      if (!uploadError) {
        await supabase.from("excel_downloads").insert({
          filename: file.name,
          report_date: new Date().toISOString().split("T")[0],
          excel_path: excelPath,
        });
      }
    } catch {
      // Non-fatal — still return the file
    }

    return new NextResponse(Buffer.from(outputBuffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    console.error("Dashboard generation error:", err);
    return NextResponse.json(
      { error: "Failed to generate dashboard. Please check the file format." },
      { status: 500 }
    );
  }
}
