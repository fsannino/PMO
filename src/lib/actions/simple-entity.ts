"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { requireAdminSession } from "@/lib/session";

/**
 * CRUDs genéricos para entidades-mestre simples (Area, Frente, Governanca,
 * Equipe, Comite, Unidade) — todas têm a mesma forma {code, name, active}.
 */

export type SimpleEntityKind =
  | "area"
  | "frente"
  | "governanca"
  | "equipe"
  | "comite"
  | "unidade";

const FormSchema = z.object({
  id: z.string().optional(),
  code: z.string().min(1, "Código obrigatório").max(20).regex(/^[A-Za-z0-9_-]+$/, "Apenas letras, números, _ e -"),
  name: z.string().min(2, "Nome muito curto").max(120),
  active: z.coerce.boolean().default(true),
});

function delegate(kind: SimpleEntityKind) {
  switch (kind) {
    case "area":       return prisma.area;
    case "frente":     return prisma.frente;
    case "governanca": return prisma.governanca;
    case "equipe":     return prisma.equipe;
    case "comite":     return prisma.comite;
    case "unidade":    return prisma.unidade;
  }
}

export type ActionResult = { ok: true } | { ok: false; error: string };

export async function upsertSimpleEntity(
  kind: SimpleEntityKind,
  formData: FormData,
): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const parsed = FormSchema.safeParse({
      id: formData.get("id") || undefined,
      code: formData.get("code"),
      name: formData.get("name"),
      active: formData.get("active") === "on" || formData.get("active") === "true",
    });
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    }
    const data = parsed.data;
    const d = delegate(kind) as { create: Function; update: Function };
    if (data.id) {
      await d.update({ where: { id: data.id }, data: { code: data.code.toUpperCase(), name: data.name, active: data.active } });
    } else {
      await d.create({ data: { code: data.code.toUpperCase(), name: data.name, active: data.active } });
    }
    revalidatePath(`/admin/${pluralPath(kind)}`);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

export async function deleteSimpleEntity(
  kind: SimpleEntityKind,
  id: string,
): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const d = delegate(kind) as { delete: Function };
    await d.delete({ where: { id } });
    revalidatePath(`/admin/${pluralPath(kind)}`);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

function pluralPath(kind: SimpleEntityKind) {
  return {
    area: "areas",
    frente: "frentes",
    governanca: "governancas",
    equipe: "equipes",
    comite: "comites",
    unidade: "unidades",
  }[kind];
}

function errorMessage(e: unknown): string {
  if (e instanceof Error) {
    if (e.message.includes("Unique constraint")) return "Código já existe.";
    if (e.message.includes("Foreign key")) return "Existe(m) registro(s) vinculado(s); inative em vez de excluir.";
    return e.message;
  }
  return "Erro desconhecido";
}
