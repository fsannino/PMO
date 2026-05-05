import NextAuth from "next-auth";
import { authOptions } from "@/lib/auth";

// Força o route como dinâmico em runtime — sem isso o build da Vercel tenta
// pré-renderizar e falha porque NextAuth precisa de env vars (NEXTAUTH_SECRET)
// que só estão disponíveis em runtime.
export const dynamic = "force-dynamic";
export const runtime = "nodejs";

const handler = NextAuth(authOptions);

export { handler as GET, handler as POST };
