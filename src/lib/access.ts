import { prisma } from "./db";
import type { Module, Role } from "./enums";

/**
 * Verifica se o usuário tem acesso ao projeto + módulo na ação requerida.
 * - ADMIN bypass total.
 * - Para os demais, consulta a tabela Access (matriz user×project×module).
 */
export type AccessAction = "read" | "write" | "admin";

export async function hasAccess(
  userId: string,
  projectId: string,
  module: Module,
  action: AccessAction = "read",
): Promise<boolean> {
  const user = await prisma.user.findUnique({
    where: { id: userId },
    select: { role: true, active: true },
  });
  if (!user || !user.active) return false;
  if (user.role === "ADMIN") return true;

  const access = await prisma.access.findUnique({
    where: { userId_projectId_module: { userId, projectId, module } },
    select: { canRead: true, canWrite: true, canAdmin: true },
  });
  if (!access) return false;
  if (action === "read") return access.canRead;
  if (action === "write") return access.canRead && access.canWrite;
  if (action === "admin") return access.canAdmin;
  return false;
}

/** Lança 403 (via thrown Response) se não tiver acesso. Use em route handlers / RSC. */
export class ForbiddenError extends Error {
  constructor(message = "Acesso negado") {
    super(message);
    this.name = "ForbiddenError";
  }
}

export async function assertAccess(
  userId: string,
  projectId: string,
  module: Module,
  action: AccessAction = "read",
) {
  const ok = await hasAccess(userId, projectId, module, action);
  if (!ok) throw new ForbiddenError();
}

/** Helpers de role (sem precisar consultar o banco). */
export function roleAtLeast(role: Role, min: Role): boolean {
  const order: Role[] = ["VIEWER", "MEMBER", "MANAGER", "ADMIN"];
  return order.indexOf(role) >= order.indexOf(min);
}

export function isAdmin(role: Role): boolean {
  return role === "ADMIN";
}
