"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { hasAccess } from "@/lib/access";
import { ModuleSchema, TaskStatusSchema, DependencyTypeSchema } from "@/lib/enums";
import { wouldCreateCycle } from "@/lib/cycle-check";

export type ActionResult<T = void> = { ok: true; data?: T } | { ok: false; error: string };

async function requireWriteOnProject(projectId: string) {
  const session = await getServerSession(authOptions);
  if (!session?.user) throw new Error("Não autenticado.");
  const project = await prisma.project.findUnique({
    where: { id: projectId },
    select: { module: true },
  });
  if (!project) throw new Error("Projeto não encontrado.");
  const module = ModuleSchema.parse(project.module);
  const ok = await hasAccess(session.user.id, projectId, module, "write");
  if (!ok) throw new Error("Sem permissão de escrita neste projeto.");
  return { userId: session.user.id, module };
}

const TaskUpsertSchema = z.object({
  id: z.string().optional(),
  projectId: z.string(),
  parentId: z.string().optional().nullable(),
  externalId: z.string().optional().nullable(),
  wbs: z.string().max(50).optional().nullable(),
  name: z.string().min(2, "Nome muito curto").max(200),
  description: z.string().max(2000).optional().nullable(),
  startDate: z.string().min(1, "Data início obrigatória"),
  endDate: z.string().min(1, "Data fim obrigatória"),
  baselineStart: z.string().optional().nullable(),
  baselineEnd: z.string().optional().nullable(),
  percentDone: z.coerce.number().min(0).max(100).default(0),
  status: TaskStatusSchema.optional(),
  isMilestone: z.coerce.boolean().default(false),
  isSummary: z.coerce.boolean().default(false),
  assigneeId: z.string().optional().nullable(),
  areaId: z.string().optional().nullable(),
  frenteId: z.string().optional().nullable(),
  equipeId: z.string().optional().nullable(),
});

export async function upsertTask(input: unknown): Promise<ActionResult<{ id: string }>> {
  try {
    const parsed = TaskUpsertSchema.safeParse(input);
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    }
    const d = parsed.data;
    const { userId } = await requireWriteOnProject(d.projectId);

    const start = new Date(d.startDate);
    const end = new Date(d.endDate);
    if (end < start) return { ok: false, error: "Data fim anterior ao início." };

    // status auto: se isMilestone e percent=100 -> DONE; se >0 -> IN_PROGRESS; senão NOT_STARTED.
    const status =
      d.status ??
      (d.percentDone >= 100 ? "DONE" : d.percentDone > 0 ? "IN_PROGRESS" : "NOT_STARTED");

    const durationDays = Math.max(
      1,
      Math.round((end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000)),
    );

    const data = {
      projectId: d.projectId,
      parentId: d.parentId || null,
      externalId: d.externalId || null,
      wbs: d.wbs || null,
      name: d.name,
      description: d.description || null,
      startDate: start,
      endDate: end,
      baselineStart: d.baselineStart ? new Date(d.baselineStart) : null,
      baselineEnd: d.baselineEnd ? new Date(d.baselineEnd) : null,
      durationDays,
      percentDone: d.percentDone,
      status,
      isMilestone: d.isMilestone,
      isSummary: d.isSummary,
      assigneeId: d.assigneeId || null,
      areaId: d.areaId || null,
      frenteId: d.frenteId || null,
      equipeId: d.equipeId || null,
    };

    let task: { id: string };
    if (d.id) {
      const before = await prisma.task.findUnique({ where: { id: d.id } });
      if (!before) return { ok: false, error: "Tarefa não encontrada." };
      if (d.parentId && d.parentId === d.id) return { ok: false, error: "Tarefa não pode ser pai de si mesma." };

      task = await prisma.task.update({ where: { id: d.id }, data, select: { id: true } });

      // diff -> history
      const changes: Array<{ field: string; oldValue: string; newValue: string }> = [];
      const compareKeys: (keyof typeof data)[] = [
        "name", "startDate", "endDate", "percentDone", "status", "isMilestone",
        "parentId", "assigneeId", "equipeId", "areaId", "frenteId",
      ];
      for (const k of compareKeys) {
        const o = (before as Record<string, unknown>)[k];
        const n = (data as Record<string, unknown>)[k];
        const ov = o instanceof Date ? o.toISOString() : String(o ?? "");
        const nv = n instanceof Date ? n.toISOString() : String(n ?? "");
        if (ov !== nv) changes.push({ field: k as string, oldValue: ov, newValue: nv });
      }
      if (changes.length) {
        await prisma.taskHistory.createMany({
          data: changes.map((c) => ({
            taskId: task.id,
            userId,
            action: "UPDATED",
            field: c.field,
            oldValue: c.oldValue,
            newValue: c.newValue,
          })),
        });
      }
    } else {
      task = await prisma.task.create({ data, select: { id: true } });
      await prisma.taskHistory.create({
        data: { taskId: task.id, userId, action: "CREATED", note: `Tarefa "${d.name}" criada.` },
      });
    }

    revalidatePath(`/projects/${d.projectId}/tasks`);
    revalidatePath(`/projects/${d.projectId}/gantt`);
    revalidatePath(`/projects/${d.projectId}`);
    return { ok: true, data: task };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}

export async function softDeleteTask(taskId: string): Promise<ActionResult> {
  try {
    const task = await prisma.task.findUnique({ where: { id: taskId } });
    if (!task) return { ok: false, error: "Tarefa não encontrada." };
    const { userId } = await requireWriteOnProject(task.projectId);
    if (task.deletedAt) return { ok: false, error: "Já está excluída." };

    await prisma.$transaction([
      prisma.task.update({ where: { id: taskId }, data: { deletedAt: new Date() } }),
      prisma.taskHistory.create({
        data: { taskId, userId, action: "DELETED", note: `Tarefa "${task.name}" excluída.` },
      }),
    ]);

    revalidatePath(`/projects/${task.projectId}/tasks`);
    revalidatePath(`/projects/${task.projectId}/gantt`);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}

export async function restoreTask(taskId: string): Promise<ActionResult> {
  try {
    const task = await prisma.task.findUnique({ where: { id: taskId } });
    if (!task) return { ok: false, error: "Tarefa não encontrada." };
    const { userId } = await requireWriteOnProject(task.projectId);
    if (!task.deletedAt) return { ok: false, error: "Tarefa não está excluída." };

    await prisma.$transaction([
      prisma.task.update({ where: { id: taskId }, data: { deletedAt: null } }),
      prisma.taskHistory.create({
        data: { taskId, userId, action: "RESTORED", note: `Tarefa "${task.name}" restaurada.` },
      }),
    ]);

    revalidatePath(`/projects/${task.projectId}/tasks`);
    revalidatePath(`/projects/${task.projectId}/gantt`);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}

const DependencySchema = z.object({
  predecessorId: z.string(),
  successorId: z.string(),
  type: DependencyTypeSchema.default("FS"),
  lagDays: z.coerce.number().int().default(0),
});

export async function createDependency(input: unknown): Promise<ActionResult> {
  try {
    const parsed = DependencySchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.message };
    const { predecessorId, successorId, type, lagDays } = parsed.data;
    if (predecessorId === successorId) return { ok: false, error: "Predecessor não pode ser igual ao sucessor." };

    const [pred, succ] = await Promise.all([
      prisma.task.findUnique({ where: { id: predecessorId }, select: { projectId: true } }),
      prisma.task.findUnique({ where: { id: successorId }, select: { projectId: true } }),
    ]);
    if (!pred || !succ) return { ok: false, error: "Tarefa(s) não encontrada(s)." };
    if (pred.projectId !== succ.projectId) return { ok: false, error: "Dependência só permitida dentro do mesmo projeto." };

    const { userId } = await requireWriteOnProject(succ.projectId);

    if (await wouldCreateCycle(predecessorId, successorId)) {
      return { ok: false, error: "Esta dependência criaria um ciclo." };
    }

    await prisma.taskDependency.create({
      data: { predecessorId, successorId, type, lagDays },
    });
    await prisma.taskHistory.create({
      data: {
        taskId: successorId,
        userId,
        action: "UPDATED",
        field: "dependency",
        newValue: `+ pred=${predecessorId} type=${type} lag=${lagDays}`,
      },
    });

    revalidatePath(`/projects/${succ.projectId}/tasks`);
    revalidatePath(`/projects/${succ.projectId}/gantt`);
    return { ok: true };
  } catch (e) {
    if (e instanceof Error && e.message.includes("Unique constraint"))
      return { ok: false, error: "Esta dependência já existe." };
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}

export async function deleteDependency(id: string): Promise<ActionResult> {
  try {
    const dep = await prisma.taskDependency.findUnique({
      where: { id },
      include: { successor: { select: { projectId: true } } },
    });
    if (!dep) return { ok: false, error: "Dependência não encontrada." };
    const { userId } = await requireWriteOnProject(dep.successor.projectId);

    await prisma.taskDependency.delete({ where: { id } });
    await prisma.taskHistory.create({
      data: {
        taskId: dep.successorId,
        userId,
        action: "UPDATED",
        field: "dependency",
        oldValue: `- pred=${dep.predecessorId}`,
      },
    });

    revalidatePath(`/projects/${dep.successor.projectId}/tasks`);
    revalidatePath(`/projects/${dep.successor.projectId}/gantt`);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}
