"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";

export default function LoginPage() {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const router = useRouter();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError("");

    const res = await fetch("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password }),
    });

    if (res.ok) {
      router.push("/dashboard");
    } else {
      setError("Incorrect password.");
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#0F2044] flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-2xl p-10 w-full max-w-md text-center">
        <div className="mb-8">
          <h1 className="text-2xl font-bold text-[#0F2044]">25N Coworking</h1>
          <p className="text-gray-500 mt-1 text-sm">Report Generator</p>
          <div className="mt-4 w-12 h-1 bg-[#C9A84C] mx-auto rounded-full" />
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <input
            type="password"
            placeholder="Enter password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoFocus
            className="w-full border border-gray-200 rounded-lg px-4 py-3 text-gray-800 focus:outline-none focus:ring-2 focus:ring-[#0F2044] text-center text-lg tracking-widest"
          />
          {error && <p className="text-red-500 text-sm">{error}</p>}
          <button
            type="submit"
            disabled={loading || !password}
            className="w-full bg-[#0F2044] hover:bg-[#1A3461] disabled:opacity-40 text-white font-semibold py-3 px-6 rounded-lg transition"
          >
            {loading ? "Checking…" : "Enter"}
          </button>
        </form>
      </div>
    </div>
  );
}
