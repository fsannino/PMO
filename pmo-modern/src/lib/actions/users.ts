"use server";

import { z } from "zod";
import bcrypt from "bcryptjs";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { requireAdminSession } from "@/lib/session";
import { RoleSchema } from "@/lib/enums";

const UserUpsertSchema = z.object({
  id: z.string().optional(),
  email: z.string().email("E-mail inválido").max(120),
  name: z.string().min(2, "Nome muito curto").max(120),
  role: RoleSchema,
  unidadeId: z.string().optional().nullable(),
  areaId: z.string().optional().nullable(),
  active: z.coerce.boolean().default(true),
  password: z.string().min(6, "Senha mínimo 6 caracteres").optional().or(z.literal("")),
});

export type ActionResult = { ok: true } | { ok: false; error: string };

export async function upsertUser(formData: FormData): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const parsed = UserUpsertSchema.safeParse({
      id: formData.get("id") || undefined,
      email: formData.get("email"),
      name: formData.get("name"),
      role: formData.get("role"),
      unidadeId: formData.get("unidadeId") || null,
      areaId: formData.get("areaId") || null,
      active: formData.get("active") === "on",
      password: formData.get("password") || "",
    });
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    }
    const data = parsed.data;
    const baseData = {
      email: data.email.toLowerCase(),
      name: data.name,
      role: data.role,
      unidadeId: data.unidadeId || null,
      areaId: data.areaId || null,
      active: data.active,
    };

    if (data.id) {
      await prisma.user.update({
        where: { id: data.id },
        data: {
          ...baseData,
          ...(data.password ? { passwordHash: await bcrypt.hash(data.password, 10) } : {}),
        },
      });
    } else {
      if (!data.password) return { ok: false, error: "Senha é obrigatória ao criar usuário." };
      await prisma.user.create({
        data: { ...baseData, passwordHash: await bcrypt.hash(data.password, 10) },
      });
    }
    revalidatePath("/admin/users");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

export async function toggleUserActive(id: string): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const u = await prisma.user.findUnique({ where: { id }, select: { active: true } });
    if (!u) return { ok: false, error: "Usuário não encontrado." };
    await prisma.user.update({ where: { id }, data: { active: !u.active } });
    revalidatePath("/admin/users");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

export async function resetPassword(id: string, newPassword: string): Promise<ActionResult> {
  try {
    await requireAdminSession();
    if (newPassword.length < 6) return { ok: false, error: "Senha mínimo 6 caracteres." };
    await prisma.user.update({
      where: { id },
      data: { passwordHash: await bcrypt.hash(newPassword, 10) },
    });
    revalidatePath("/admin/users");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

function errorMessage(e: unknown): string {
  if (e instanceof Error) {
    if (e.message.includes("Unique constraint")) return "E-mail já cadastrado.";
    return e.message;
  }
  return "Erro desconhecido";
}
