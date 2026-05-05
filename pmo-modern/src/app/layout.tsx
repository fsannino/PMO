import type { Metadata } from "next";
import { Providers } from "@/components/Providers";
import "./globals.css";

// Força toda a app como dinâmica em runtime — todas as páginas usam
// getServerSession() e Prisma; tentar coletar page data em build (Vercel)
// falha porque envs/db não estão acessíveis nessa fase.
export const dynamic = "force-dynamic";
export const revalidate = 0;

export const metadata: Metadata = {
  title: "PMO CollabZ",
  description: "Sistema integrado de gestão de projetos",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pt-BR">
      <body>
        <Providers>{children}</Providers>
      </body>
    </html>
  );
}
