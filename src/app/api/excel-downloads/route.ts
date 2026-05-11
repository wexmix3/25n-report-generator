import { NextRequest, NextResponse } from "next/server";
import { createServiceClient } from "@/lib/supabase-server";
import { verifySessionToken } from "@/lib/auth";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!(await verifySessionToken(request.cookies.get("auth")?.value ?? ""))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createServiceClient();
  const { data: downloads, error } = await supabase
    .from("excel_downloads")
    .select("*")
    .order("created_at", { ascending: false })
    .limit(50);

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  const downloadsWithUrls = await Promise.all(
    (downloads ?? []).map(async (record) => {
      const { data } = await supabase.storage
        .from("reports")
        .createSignedUrl(record.excel_path, 604800); // 7-day TTL
      return { ...record, downloadUrl: data?.signedUrl ?? null };
    })
  );

  return NextResponse.json({ downloads: downloadsWithUrls });
}
