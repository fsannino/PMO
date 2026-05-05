/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  // standalone: gera .next/standalone com server.js auto-contido (Docker).
  // Na Vercel não é necessário — ela usa o output padrão (.next).
  // Ativar só quando rodando build para Docker (DOCKER_BUILD=1).
  output: process.env.DOCKER_BUILD === "1" ? "standalone" : undefined,
  experimental: {
    serverActions: {
      bodySizeLimit: "20mb",
    },
  },
};

export default nextConfig;
