import { NextRequest, NextResponse } from "next/server";
import { createServiceClient } from "@/lib/supabase-server";
import { verifySessionToken } from "@/lib/auth";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!(await verifySessionToken(request.cookies.get("auth")?.value ?? ""))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createServiceClient();
  const { data: reports, error } = await supabase
    .from("reports")
    .select("*")
    .order("created_at", { ascending: false })
    .limit(50);

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  const reportsWithUrls = await Promise.all(
    (reports ?? []).map(async (report) => {
      const { data } = await supabase.storage
        .from("reports")
        .createSignedUrl(report.pdf_path, 604800); // 7-day TTL
      return { ...report, downloadUrl: data?.signedUrl ?? null };
    })
  );

  return NextResponse.json({ reports: reportsWithUrls });
}
