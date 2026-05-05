"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { getMeasurementWindow, periodOf } from "@/lib/closing-window";

export type ActionResult = { ok: true } | { ok: false; error: string };

const InputSchema = z.object({
  taskId: z.string(),
  percentDone: z.coerce.number().min(0).max(100),
  hoursWorked: z.coerce.number().min(0).optional().nullable(),
  comment: z.string().max(1000).optional().nullable(),
  confirm: z.coerce.boolean().default(false),
});

async function loadAndAuth(taskId: string) {
  const session = await getServerSession(authOptions);
  if (!session?.user) throw new Error("Não autenticado.");
  const task = await prisma.task.findUnique({
    where: { id: taskId },
    include: { project: { select: { id: true, module: true } } },
  });
  if (!task) throw new Error("Tarefa não encontrada.");
  if (task.deletedAt) throw new Error("Tarefa excluída.");
  if (task.isSummary) throw new Error("Tarefas-resumo não recebem medição.");
  const moduleParsed = ModuleSchema.parse(task.project.module);
  const ok = await hasAccess(session.user.id, task.project.id, moduleParsed, "write");
  if (!ok) throw new Error("Sem permissão para medir esta tarefa.");
  return { session, task };
}

/**
 * Cria/atualiza medição do período corrente. Se confirm=true, valida janela
 * HB e cria MeasurementLock impedindo edições futuras no período.
 */
export async function upsertMeasurement(input: unknown): Promise<ActionResult> {
  try {
    const parsed = InputSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    const { taskId, percentDone, hoursWorked, comment, confirm } = parsed.data;

    const { session, task } = await loadAndAuth(taskId);
    const period = periodOf();

    // bloqueio: período já confirmado
    const lock = await prisma.measurementLock.findUnique({
      where: { taskId_period: { taskId, period } },
    });
    if (lock) {
      return { ok: false, error: "Medição deste período já está confirmada — bloqueada." };
    }

    if (confirm) {
      const win = await getMeasurementWindow(task.project.id);
      if (!win.open) {
        return { ok: false, error: `Janela de medição fechada: ${win.reason}` };
      }
    }

    // upsert da medição (uma por taskId × period × user)
    const existing = await prisma.measurement.findFirst({
      where: { taskId, period, userId: session.user.id, confirmed: false },
    });

    let measurementId: string;
    if (existing) {
      const m = await prisma.measurement.update({
        where: { id: existing.id },
        data: {
          percentDone,
          hoursWorked: hoursWorked ?? null,
          comment: comment ?? null,
          confirmed: confirm,
          confirmedAt: confirm ? new Date() : null,
        },
      });
      measurementId = m.id;
    } else {
      const m = await prisma.measurement.create({
        data: {
          taskId,
          userId: session.user.id,
          percentDone,
          hoursWorked: hoursWorked ?? null,
          comment: comment ?? null,
          period,
          confirmed: confirm,
          confirmedAt: confirm ? new Date() : null,
        },
      });
      measurementId = m.id;
    }

    if (confirm) {
      await prisma.measurementLock.create({
        data: { taskId, period, measurementId },
      });
      // propaga percentDone para a tarefa + status + history
      const newStatus = percentDone >= 100 ? "DONE" : percentDone > 0 ? "IN_PROGRESS" : "NOT_STARTED";
      await prisma.task.update({
        where: { id: taskId },
        data: { percentDone, status: newStatus },
      });
      await prisma.taskHistory.create({
        data: {
          taskId,
          userId: session.user.id,
          action: "MEASURED",
          field: "percentDone",
          newValue: String(percentDone),
          note: `Medição confirmada (${period})${comment ? `: ${comment}` : ""}`,
        },
      });
    }

    revalidatePath(`/projects/${task.project.id}/measurement`);
    revalidatePath(`/projects/${task.project.id}/tasks`);
    revalidatePath(`/projects/${task.project.id}`);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}
