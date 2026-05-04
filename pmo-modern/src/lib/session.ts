import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";

/**
 * Garante que o request atual é de um ADMIN. Use em server actions / RSC.
 * Lança se não for — caller deve devolver { error } ou redirect.
 */
export async function requireAdminSession() {
  const session = await getServerSession(authOptions);
  if (!session?.user) throw new Error("Não autenticado.");
  if (session.user.role !== "ADMIN") throw new Error("Acesso negado: requer ADMIN.");
  return session;
}

/** Retorna sessão ou null sem lançar — útil para diferenciação de UI. */
export async function getSessionOrNull() {
  return getServerSession(authOptions);
}
