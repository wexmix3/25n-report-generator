import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  serverExternalPackages: ["@react-pdf/renderer", "xlsx"],
  typescript: { ignoreBuildErrors: true },
};

export default nextConfig;
