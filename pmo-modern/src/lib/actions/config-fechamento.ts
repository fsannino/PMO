"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { requireAdminSession } from "@/lib/session";

const TimeRe = /^([01]\d|2[0-3]):[0-5]\d$/;

const DaySchema = z.object({
  dayOfWeek: z.number().int().min(0).max(6),
  enabled: z.boolean(),
  startTime: z.string().regex(TimeRe, "Hora inicial inválida (HH:mm)"),
  endTime: z.string().regex(TimeRe, "Hora final inválida (HH:mm)"),
});

const PayloadSchema = z.object({
  projectId: z.string().nullable(), // null = global
  days: z.array(DaySchema).length(7),
});

export type ActionResult = { ok: true } | { ok: false; error: string };

export async function saveConfigFechamento(payload: unknown): Promise<ActionResult> {
  try {
    await requireAdminSession();
    const parsed = PayloadSchema.safeParse(payload);
    if (!parsed.success) return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    const { projectId, days } = parsed.data;

    for (const d of days) {
      if (d.enabled && d.startTime >= d.endTime) {
        return { ok: false, error: `Dia ${d.dayOfWeek}: hora inicial deve ser anterior à final.` };
      }
    }

    await prisma.$transaction(async (tx) => {
      for (const d of days) {
        const existing = await tx.configFechamento.findFirst({
          where: { projectId, dayOfWeek: d.dayOfWeek },
        });
        if (existing) {
          await tx.configFechamento.update({
            where: { id: existing.id },
            data: { enabled: d.enabled, startTime: d.startTime, endTime: d.endTime },
          });
        } else {
          await tx.configFechamento.create({
            data: {
              projectId,
              dayOfWeek: d.dayOfWeek,
              enabled: d.enabled,
              startTime: d.startTime,
              endTime: d.endTime,
            },
          });
        }
      }
    });

    revalidatePath("/admin/config-fechamento");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}
