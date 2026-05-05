/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  // standalone: gera .next/standalone com server.js auto-contido (Docker/Vercel)
  output: "standalone",
  // serverActions são habilitadas por padrão desde Next 14.2; o body limit
  // foi movido para o nível raiz da config:
  experimental: {
    serverActions: {
      bodySizeLimit: "20mb",
    },
  },
};

export default nextConfig;
