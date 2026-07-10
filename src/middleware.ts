import { NextResponse, type NextRequest } from "next/server";
import { verifySessionToken } from "@/lib/auth";

export async function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;
  const token = request.cookies.get("auth")?.value ?? "";
  const authed = await verifySessionToken(token);

  // These routes use their own x-internal-key auth (called from Python scripts, not the browser)
  const internalKeyRoutes = ["/api/generate-month-close", "/api/generate-financial-packet", "/api/generate-gl-variance"];
  if (!authed && pathname !== "/login" && !pathname.startsWith("/api/auth") && !internalKeyRoutes.includes(pathname)) {
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
