import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

// OAuth callback no longer used — password auth only
export async function GET() {
  return NextResponse.redirect("/login");
}
