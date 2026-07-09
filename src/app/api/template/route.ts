import { NextRequest, NextResponse } from "next/server";
import { verifySessionToken } from "@/lib/auth";
import { createServiceClient } from "@/lib/supabase-server";
import * as XLSX from "xlsx-js-style";

export const dynamic = "force-dynamic";

const TEMPLATE_BUCKET = "reports";
const TEMPLATE_PATH   = "templates/master.xlsx";
const META_PATH       = "templates/master.meta.json";
const MAX_FILE_SIZE   = 20 * 1024 * 1024; // 20 MB

interface TemplateMeta {
  originalName:     string;
  originalBaseName: string;
  sheetNames:       string[];
  sheetCount:       number;
  namedRangeCount:  number;
  uploadedAt:       string;
}

async function auth(request: NextRequest) {
  return verifySessionToken(request.cookies.get("auth")?.value ?? "");
}

// GET: return template metadata (exists, size, last updated, sheet inventory)
export async function GET(request: NextRequest) {
  if (!(await auth(request))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createServiceClient();
  const { data, error } = await supabase.storage
    .from(TEMPLATE_BUCKET)
    .list("templates", { search: "master.xlsx" });

  if (error) {
    return NextResponse.json({ error: "Storage error" }, { status: 500 });
  }

  const file = data?.find((f) => f.name === "master.xlsx");
  if (!file) {
    return NextResponse.json({ exists: false });
  }

  // Try to load companion metadata
  let meta: TemplateMeta | null = null;
  try {
    const { data: metaData, error: metaErr } = await supabase.storage
      .from(TEMPLATE_BUCKET)
      .download(META_PATH);
    if (!metaErr && metaData) {
      meta = JSON.parse(await metaData.text()) as TemplateMeta;
    }
  } catch {
    // Non-fatal — meta may not exist for templates uploaded before this feature
  }

  return NextResponse.json({
    exists:          true,
    size:            file.metadata?.size ?? null,
    updatedAt:       file.updated_at ?? file.created_at ?? null,
    originalName:    meta?.originalName    ?? null,
    sheetNames:      meta?.sheetNames      ?? null,
    sheetCount:      meta?.sheetCount      ?? null,
    namedRangeCount: meta?.namedRangeCount ?? null,
  });
}

// POST: upload or replace template
export async function POST(request: NextRequest) {
  if (!(await auth(request))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const formData = await request.formData();
  const file = formData.get("file") as File | null;

  if (!file) {
    return NextResponse.json({ error: "No file provided" }, { status: 400 });
  }
  if (!file.name.match(/\.xlsx?$/i)) {
    return NextResponse.json({ error: "File must be .xlsx or .xls" }, { status: 400 });
  }
  if (file.size > MAX_FILE_SIZE) {
    return NextResponse.json({ error: "File too large (max 20 MB)" }, { status: 400 });
  }

  const buffer = Buffer.from(await file.arrayBuffer());
  const supabase = createServiceClient();

  // Parse workbook to extract metadata
  let sheetNames: string[] = [];
  let namedRangeCount = 0;
  try {
    const wb = XLSX.read(buffer, { type: "buffer", cellDates: false });
    sheetNames = wb.SheetNames;
    const names = (wb.Workbook as Record<string, unknown> | undefined)?.Names;
    namedRangeCount = Array.isArray(names) ? names.length : 0;
  } catch {
    // Non-fatal — upload still proceeds without metadata
  }

  const originalBaseName = file.name.replace(/\.xlsx?$/i, "");
  const meta: TemplateMeta = {
    originalName:     file.name,
    originalBaseName,
    sheetNames,
    sheetCount:       sheetNames.length,
    namedRangeCount,
    uploadedAt:       new Date().toISOString(),
  };

  // Remove existing template + meta first (upsert via remove + upload)
  await supabase.storage.from(TEMPLATE_BUCKET).remove([TEMPLATE_PATH, META_PATH]);

  const { error } = await supabase.storage
    .from(TEMPLATE_BUCKET)
    .upload(TEMPLATE_PATH, buffer, {
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

  if (error) {
    return NextResponse.json({ error: `Upload failed: ${error.message}` }, { status: 500 });
  }

  // Upload companion metadata (best-effort — don't block on failure)
  try {
    const metaBytes = Buffer.from(JSON.stringify(meta, null, 2));
    await supabase.storage
      .from(TEMPLATE_BUCKET)
      .upload(META_PATH, metaBytes, { contentType: "application/json" });
  } catch {
    // Non-fatal
  }

  return NextResponse.json({ success: true });
}

// DELETE: remove template (revert to regeneration mode)
export async function DELETE(request: NextRequest) {
  if (!(await auth(request))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createServiceClient();
  const { error } = await supabase.storage
    .from(TEMPLATE_BUCKET)
    .remove([TEMPLATE_PATH]);

  if (error) {
    return NextResponse.json({ error: `Delete failed: ${error.message}` }, { status: 500 });
  }

  return NextResponse.json({ success: true });
}
