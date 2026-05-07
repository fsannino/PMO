"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { requireAdminSession } from "@/lib/session";
import { ModuleSchema, ProjectStatusSchema } from "@/lib/enums";

const ProjectSchema = z.object({
  id: z.string().optional(),
  code: z.string().min(2, "Código mínimo 2 caracteres").max(30).regex(/^[A-Za-z0-9_-]+$/, "Apenas letras, números, _ e -"),
  name: z.string().min(2).max(200),
  description: z.string().max(2000).optional().or(z.literal("")),
  module: ModuleSchema,
  status: ProjectStatusSchema,
  priority: z.coerce.number().int().min(0).max(99).default(0),
  ownerId: z.string().min(1, "Owner obrigatório"),
  unidadeId: z.string().optional().or(z.literal("")),
  governancaId: z.string().optional().or(z.literal("")),
  startDate: z.string().min(1, "Data de início obrigatória"),
  endDate: z.string().min(1, "Data de fim obrigatória"),
  baselineDate: z.string().optional().or(z.literal("")),
});

export type ActionResult = { ok: true } | { ok: false; error: string };

export async function upsertProject(formData: FormData): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const parsed = ProjectSchema.safeParse({
      id: formData.get("id") || undefined,
      code: formData.get("code"),
      name: formData.get("name"),
      description: formData.get("description") ?? "",
      module: formData.get("module"),
      status: formData.get("status"),
      priority: formData.get("priority") ?? 0,
      ownerId: formData.get("ownerId"),
      unidadeId: formData.get("unidadeId") ?? "",
      governancaId: formData.get("governancaId") ?? "",
      startDate: formData.get("startDate"),
      endDate: formData.get("endDate"),
      baselineDate: formData.get("baselineDate") ?? "",
    });
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues.map((i) => `${i.path.join(".")}: ${i.message}`).join(", ") };
    }
    const d = parsed.data;
    const start = new Date(d.startDate);
    const end = new Date(d.endDate);
    if (end < start) return { ok: false, error: "Data de fim anterior ao início." };

    const data = {
      code: d.code.toUpperCase(),
      name: d.name,
      description: d.description || null,
      module: d.module,
      status: d.status,
      priority: d.priority,
      ownerId: d.ownerId,
      unidadeId: d.unidadeId || null,
      governancaId: d.governancaId || null,
      startDate: start,
      endDate: end,
      baselineDate: d.baselineDate ? new Date(d.baselineDate) : null,
    };

    if (d.id) {
      await prisma.project.update({ where: { id: d.id }, data });
    } else {
      await prisma.project.create({ data });
    }
    revalidatePath("/admin/projects");
    revalidatePath("/projects");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

export async function deleteProject(id: string): Promise<ActionResult> {
  try {
    await requireAdminSession();
    await prisma.project.delete({ where: { id } });
    revalidatePath("/admin/projects");
    revalidatePath("/projects");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: errorMessage(e) };
  }
}

function errorMessage(e: unknown): string {
  if (e instanceof Error) {
    if (e.message.includes("Unique constraint")) return "Código de projeto já existe.";
    if (e.message.includes("Foreign key")) return "Existe(m) registro(s) vinculado(s); arquive em vez de excluir.";
    return e.message;
  }
  return "Erro desconhecido";
}
