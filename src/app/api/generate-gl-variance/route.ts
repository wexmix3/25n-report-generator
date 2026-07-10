import { NextRequest, NextResponse } from "next/server";
import { renderToBuffer } from "@react-pdf/renderer";
import React from "react";
import { GLVarianceDocument } from "@/lib/gl-variance-pdf";
import type { GLVarianceData } from "@/lib/gl-variance-pdf";

export const dynamic = "force-dynamic";

function isAuthorized(request: NextRequest): boolean {
  const key = process.env.MONTH_CLOSE_API_KEY;
  if (!key) return false;
  return request.headers.get("x-internal-key") === key;
}

export async function POST(request: NextRequest) {
  if (!isAuthorized(request)) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  let data: GLVarianceData;
  try {
    data = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }

  if (!data.location || !data.month) {
    return NextResponse.json({ error: "Missing required fields: location, month" }, { status: 400 });
  }

  let pdfBuffer: Buffer;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    pdfBuffer = await renderToBuffer(
      React.createElement(GLVarianceDocument, { data }) as any
    );
  } catch (err) {
    console.error("GL variance PDF render error:", err);
    return NextResponse.json({ error: "Failed to render PDF" }, { status: 500 });
  }

  const slug = `${data.location.toLowerCase()}-${data.month.replace(/\s+/g, "-").toLowerCase()}`;
  const filename = `25N-${slug}-gl-variance.pdf`;

  return new NextResponse(Buffer.from(pdfBuffer), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${filename}"`,
    },
  });
}
