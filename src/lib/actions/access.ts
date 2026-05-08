"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { requireAdminSession } from "@/lib/session";
import { ModuleSchema } from "@/lib/enums";

const RowSchema = z.object({
  userId: z.string(),
  canRead: z.boolean(),
  canWrite: z.boolean(),
  canAdmin: z.boolean(),
});
const PayloadSchema = z.object({
  projectId: z.string(),
  module: ModuleSchema,
  rows: z.array(RowSchema),
});

export type ActionResult = { ok: true } | { ok: false; error: string };

export async function saveAccessMatrix(payload: unknown): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const parsed = PayloadSchema.safeParse(payload);
    if (!parsed.success) return { ok: false, error: parsed.error.message };
    const { projectId, module, rows } = parsed.data;

    await prisma.$transaction(async (tx) => {
      for (const r of rows) {
        const noAccess = !r.canRead && !r.canWrite && !r.canAdmin;
        if (noAccess) {
          await tx.access.deleteMany({
            where: { userId: r.userId, projectId, module },
          });
          continue;
        }
        await tx.access.upsert({
          where: { userId_projectId_module: { userId: r.userId, projectId, module } },
          update: { canRead: r.canRead, canWrite: r.canWrite, canAdmin: r.canAdmin },
          create: {
            userId: r.userId,
            projectId,
            module,
            canRead: r.canRead,
            canWrite: r.canWrite,
            canAdmin: r.canAdmin,
          },
        });
      }
    });

    revalidatePath("/admin/access");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}
