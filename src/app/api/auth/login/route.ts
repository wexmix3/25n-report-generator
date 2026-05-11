import { NextRequest, NextResponse } from "next/server";
import { createSessionToken } from "@/lib/auth";

export const dynamic = "force-dynamic";

// Simple in-memory rate limiter: max 5 attempts per IP per 15 minutes
const attempts = new Map<string, { count: number; resetAt: number }>();

export async function POST(request: NextRequest) {
  const ip =
    request.headers.get("x-forwarded-for")?.split(",")[0].trim() ??
    request.headers.get("x-real-ip") ??
    "unknown";
  const now = Date.now();
  const WINDOW = 15 * 60 * 1000;
  const LIMIT = 5;

  const entry = attempts.get(ip);
  if (entry) {
    if (now > entry.resetAt) {
      attempts.set(ip, { count: 1, resetAt: now + WINDOW });
    } else if (entry.count >= LIMIT) {
      return NextResponse.json({ error: "Too many attempts. Try again later." }, { status: 429 });
    } else {
      entry.count++;
    }
  } else {
    attempts.set(ip, { count: 1, resetAt: now + WINDOW });
  }

  const { password } = await request.json();

  if (password !== process.env.APP_PASSWORD) {
    return NextResponse.json({ error: "Incorrect password" }, { status: 401 });
  }

  const token = await createSessionToken();
  const response = NextResponse.json({ ok: true });
  response.cookies.set("auth", token, {
    httpOnly: true,
    secure: true,
    sameSite: "lax",
    maxAge: 60 * 60 * 24 * 30, // 30 days
    path: "/",
  });
  return response;
}
