import { NextResponse, type NextRequest } from "next/server";
import { verifySessionToken } from "@/lib/auth";

export async function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;
  const token = request.cookies.get("auth")?.value ?? "";
  const authed = await verifySessionToken(token);

  // /api/generate-month-close uses its own x-internal-key auth
  if (!authed && pathname !== "/login" && !pathname.startsWith("/api/auth") && pathname !== "/api/generate-month-close" && pathname !== "/api/generate-financial-packet") {
    return NextResponse.redirect(new URL("/login", request.url));
  }

  if (authed && pathname === "/login") {
    return NextResponse.redirect(new URL("/dashboard", request.url));
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
