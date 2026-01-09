import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  output: "standalone", // Required for Docker deployment
  serverExternalPackages: ["docx", "mathjax-full"],
};

export default nextConfig;
