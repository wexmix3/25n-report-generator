// Uses Web Crypto API (globalThis.crypto.subtle) — works in Edge Runtime and Node.js
const SECRET = process.env.SESSION_SECRET ?? "dev-secret-change-in-prod";
const TTL = 30 * 24 * 60 * 60 * 1000; // 30 days

async function getKey(secret: string): Promise<CryptoKey> {
  return globalThis.crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign", "verify"]
  );
}

function toBase64Url(buf: ArrayBuffer): string {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

function fromBase64Url(s: string): string {
  return atob(s.replace(/-/g, "+").replace(/_/g, "/"));
}

export async function createSessionToken(): Promise<string> {
  const expiry = Date.now() + TTL;
  const payload = `auth:${expiry}`;
  const key = await getKey(SECRET);
  const sig = await globalThis.crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload));
  return `${toBase64Url(new TextEncoder().encode(payload).buffer)}.${toBase64Url(sig)}`;
}

export async function verifySessionToken(token: string): Promise<boolean> {
  try {
    const dotIdx = token.lastIndexOf(".");
    if (dotIdx < 0) return false;
    const payloadB64 = token.slice(0, dotIdx);
    const sigB64 = token.slice(dotIdx + 1);
    const payload = fromBase64Url(payloadB64);
    const sigBytes = Uint8Array.from(atob(sigB64.replace(/-/g, "+").replace(/_/g, "/")), (c) => c.charCodeAt(0));
    const key = await getKey(SECRET);
    const valid = await globalThis.crypto.subtle.verify("HMAC", key, sigBytes, new TextEncoder().encode(payload));
    if (!valid) return false;
    const expiry = parseInt(payload.split(":")[1]);
    return !isNaN(expiry) && Date.now() < expiry;
  } catch {
    return false;
  }
}
