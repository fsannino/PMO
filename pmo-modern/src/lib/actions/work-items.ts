"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { hasAccess } from "@/lib/access";
import {
  ModuleSchema,
  IssueStatusSchema,
  RiskStatusSchema,
  ActionStatusSchema,
  ChangeRequestStatusSchema,
  SeveritySchema,
  PrioritySchema,
} from "@/lib/enums";

export type ActionResult<T = void> = { ok: true; data?: T } | { ok: false; error: string };

async function requireWriteOnProject(projectId: string) {
  const session = await getServerSession(authOptions);
  if (!session?.user) throw new Error("Não autenticado.");
  const project = await prisma.project.findUnique({ where: { id: projectId }, select: { module: true } });
  if (!project) throw new Error("Projeto não encontrado.");
  const module = ModuleSchema.parse(project.module);
  const ok = await hasAccess(session.user.id, projectId, module, "write");
  if (!ok) throw new Error("Sem permissão de escrita neste projeto.");
  return { userId: session.user.id, module };
}

function revalidateProject(projectId: string) {
  revalidatePath(`/projects/${projectId}`);
  revalidatePath(`/projects/${projectId}/issues`);
  revalidatePath(`/projects/${projectId}/risks`);
  revalidatePath(`/projects/${projectId}/actions`);
  revalidatePath(`/projects/${projectId}/change-requests`);
  revalidatePath("/issues");
  revalidatePath("/risks");
  revalidatePath("/change-requests");
}

// ════════════════════════════════════════════════════════════════════════
// ISSUES
// ════════════════════════════════════════════════════════════════════════

const IssueSchema = z.object({
  id: z.string().optional(),
  projectId: z.string(),
  title: z.string().min(2).max(200),
  description: z.string().max(4000).optional().nullable(),
  ownerId: z.string().optional().nullable(),
  taskId: z.string().optional().nullable(),
  areaId: z.string().optional().nullable(),
  severity: SeveritySchema.optional(),
  priority: PrioritySchema.optional(),
  status: IssueStatusSchema.optional(),
});

export async function upsertIssue(input: unknown): Promise<ActionResult<{ id: string }>> {
  try {
    const parsed = IssueSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    const d = parsed.data;
    await requireWriteOnProject(d.projectId);

    const data = {
      projectId: d.projectId,
      title: d.title,
      description: d.description || null,
      ownerId: d.ownerId || null,
      taskId: d.taskId || null,
      areaId: d.areaId || null,
      severity: d.severity ?? "MEDIUM",
      priority: d.priority ?? "MEDIUM",
      status: d.status ?? "OPEN",
      closedAt: d.status === "CLOSED" ? new Date() : null,
    };
    let id: string;
    if (d.id) {
      const r = await prisma.issue.update({ where: { id: d.id }, data, select: { id: true } });
      id = r.id;
    } else {
      const r = await prisma.issue.create({ data, select: { id: true } });
      id = r.id;
    }
    revalidateProject(d.projectId);
    return { ok: true, data: { id } };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

export async function deleteIssue(id: string): Promise<ActionResult> {
  try {
    const it = await prisma.issue.findUnique({ where: { id } });
    if (!it) return { ok: false, error: "Não encontrada." };
    await requireWriteOnProject(it.projectId);
    await prisma.issue.update({ where: { id }, data: { deletedAt: new Date() } });
    revalidateProject(it.projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

// ════════════════════════════════════════════════════════════════════════
// RISKS
// ════════════════════════════════════════════════════════════════════════

const RiskSchema = z.object({
  id: z.string().optional(),
  projectId: z.string(),
  title: z.string().min(2).max(200),
  description: z.string().max(4000).optional().nullable(),
  ownerId: z.string().optional().nullable(),
  probability: z.coerce.number().min(0).max(1).default(0.5),
  impact: z.coerce.number().min(0).max(1).default(0.5),
  mitigation: z.string().max(2000).optional().nullable(),
  status: RiskStatusSchema.optional(),
});

export async function upsertRisk(input: unknown): Promise<ActionResult<{ id: string }>> {
  try {
    const parsed = RiskSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    const d = parsed.data;
    await requireWriteOnProject(d.projectId);

    const data = {
      projectId: d.projectId,
      title: d.title,
      description: d.description || null,
      ownerId: d.ownerId || null,
      probability: d.probability,
      impact: d.impact,
      exposure: d.probability * d.impact,
      mitigation: d.mitigation || null,
      status: d.status ?? "IDENTIFIED",
    };
    let id: string;
    if (d.id) {
      const r = await prisma.risk.update({ where: { id: d.id }, data, select: { id: true } });
      id = r.id;
    } else {
      const r = await prisma.risk.create({ data, select: { id: true } });
      id = r.id;
    }
    revalidateProject(d.projectId);
    return { ok: true, data: { id } };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

export async function deleteRisk(id: string): Promise<ActionResult> {
  try {
    const it = await prisma.risk.findUnique({ where: { id } });
    if (!it) return { ok: false, error: "Não encontrado." };
    await requireWriteOnProject(it.projectId);
    await prisma.risk.update({ where: { id }, data: { deletedAt: new Date() } });
    revalidateProject(it.projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

// ════════════════════════════════════════════════════════════════════════
// ACTIONS
// ════════════════════════════════════════════════════════════════════════

const ActionItemSchema = z.object({
  id: z.string().optional(),
  projectId: z.string(),
  title: z.string().min(2).max(200),
  description: z.string().max(4000).optional().nullable(),
  ownerId: z.string().optional().nullable(),
  dueDate: z.string().optional().nullable(),
  status: ActionStatusSchema.optional(),
  priority: PrioritySchema.optional(),
});

export async function upsertActionItem(input: unknown): Promise<ActionResult<{ id: string }>> {
  try {
    const parsed = ActionItemSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    const d = parsed.data;
    await requireWriteOnProject(d.projectId);

    const data = {
      projectId: d.projectId,
      title: d.title,
      description: d.description || null,
      ownerId: d.ownerId || null,
      dueDate: d.dueDate ? new Date(d.dueDate) : null,
      status: d.status ?? "OPEN",
      priority: d.priority ?? "MEDIUM",
    };
    let id: string;
    if (d.id) {
      const r = await prisma.action.update({ where: { id: d.id }, data, select: { id: true } });
      id = r.id;
    } else {
      const r = await prisma.action.create({ data, select: { id: true } });
      id = r.id;
    }
    revalidateProject(d.projectId);
    return { ok: true, data: { id } };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

export async function deleteActionItem(id: string): Promise<ActionResult> {
  try {
    const it = await prisma.action.findUnique({ where: { id } });
    if (!it) return { ok: false, error: "Não encontrada." };
    await requireWriteOnProject(it.projectId);
    await prisma.action.update({ where: { id }, data: { deletedAt: new Date() } });
    revalidateProject(it.projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

// ════════════════════════════════════════════════════════════════════════
// CHANGE REQUESTS
// ════════════════════════════════════════════════════════════════════════

const CRSchema = z.object({
  id: z.string().optional(),
  projectId: z.string(),
  name: z.string().min(2).max(200),
  description: z.string().max(4000).optional().nullable(),
  ownerId: z.string().optional().nullable(),
  comiteId: z.string().optional().nullable(),
  priority: PrioritySchema.optional(),
  status: ChangeRequestStatusSchema.optional(),
  dueDate: z.string().optional().nullable(),
});

export async function upsertChangeRequest(input: unknown): Promise<ActionResult<{ id: string }>> {
  try {
    const parsed = CRSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.issues.map((i) => i.message).join(", ") };
    const d = parsed.data;
    await requireWriteOnProject(d.projectId);

    const data = {
      projectId: d.projectId,
      name: d.name,
      description: d.description || null,
      ownerId: d.ownerId || null,
      comiteId: d.comiteId || null,
      priority: d.priority ?? "MEDIUM",
      status: d.status ?? "OPEN",
      dueDate: d.dueDate ? new Date(d.dueDate) : null,
    };
    let id: string;
    if (d.id) {
      const r = await prisma.changeRequest.update({ where: { id: d.id }, data, select: { id: true } });
      id = r.id;
    } else {
      const r = await prisma.changeRequest.create({ data, select: { id: true } });
      id = r.id;
    }
    revalidateProject(d.projectId);
    return { ok: true, data: { id } };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

export async function deleteChangeRequest(id: string): Promise<ActionResult> {
  try {
    const it = await prisma.changeRequest.findUnique({ where: { id } });
    if (!it) return { ok: false, error: "Não encontrado." };
    await requireWriteOnProject(it.projectId);
    await prisma.changeRequest.update({ where: { id }, data: { deletedAt: new Date() } });
    revalidateProject(it.projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

const CRIncSchema = z.object({
  changeRequestId: z.string(),
  description: z.string().min(2).max(2000),
  appliedAt: z.string().optional().nullable(),
});

export async function addCRIncrement(input: unknown): Promise<ActionResult> {
  try {
    const parsed = CRIncSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.message };
    const d = parsed.data;
    const cr = await prisma.changeRequest.findUnique({ where: { id: d.changeRequestId }, select: { projectId: true } });
    if (!cr) return { ok: false, error: "CR não encontrado." };
    await requireWriteOnProject(cr.projectId);
    const last = await prisma.changeRequestIncrement.findFirst({
      where: { changeRequestId: d.changeRequestId },
      orderBy: { sequence: "desc" },
      select: { sequence: true },
    });
    await prisma.changeRequestIncrement.create({
      data: {
        changeRequestId: d.changeRequestId,
        sequence: (last?.sequence ?? 0) + 1,
        description: d.description,
        appliedAt: d.appliedAt ? new Date(d.appliedAt) : null,
      },
    });
    revalidateProject(cr.projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

// ════════════════════════════════════════════════════════════════════════
// COMENTÁRIOS POLIMÓRFICOS
// ════════════════════════════════════════════════════════════════════════

const CommentSchema = z.object({
  body: z.string().min(1).max(4000),
  taskId: z.string().optional().nullable(),
  issueId: z.string().optional().nullable(),
  riskId: z.string().optional().nullable(),
  actionId: z.string().optional().nullable(),
  changeRequestId: z.string().optional().nullable(),
});

export async function addComment(input: unknown): Promise<ActionResult> {
  try {
    const parsed = CommentSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.message };
    const d = parsed.data;
    const session = await getServerSession(authOptions);
    if (!session?.user) return { ok: false, error: "Não autenticado." };

    // resolver projectId pelo item alvo para validar acesso + revalidar paths
    let projectId: string | null = null;
    if (d.issueId)  projectId = (await prisma.issue.findUnique({ where: { id: d.issueId }, select: { projectId: true } }))?.projectId ?? null;
    else if (d.riskId) projectId = (await prisma.risk.findUnique({ where: { id: d.riskId }, select: { projectId: true } }))?.projectId ?? null;
    else if (d.actionId) projectId = (await prisma.action.findUnique({ where: { id: d.actionId }, select: { projectId: true } }))?.projectId ?? null;
    else if (d.changeRequestId) projectId = (await prisma.changeRequest.findUnique({ where: { id: d.changeRequestId }, select: { projectId: true } }))?.projectId ?? null;
    else if (d.taskId) projectId = (await prisma.task.findUnique({ where: { id: d.taskId }, select: { projectId: true } }))?.projectId ?? null;
    if (!projectId) return { ok: false, error: "Item alvo não encontrado." };

    await requireWriteOnProject(projectId);
    await prisma.comment.create({
      data: {
        body: d.body,
        authorId: session.user.id,
        taskId: d.taskId || null,
        issueId: d.issueId || null,
        riskId: d.riskId || null,
        actionId: d.actionId || null,
        changeRequestId: d.changeRequestId || null,
      },
    });
    revalidateProject(projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}

export async function deleteComment(id: string): Promise<ActionResult> {
  try {
    const c = await prisma.comment.findUnique({
      where: { id },
      include: {
        issue: { select: { projectId: true } },
        risk: { select: { projectId: true } },
        action: { select: { projectId: true } },
        changeRequest: { select: { projectId: true } },
        task: { select: { projectId: true } },
      },
    });
    if (!c) return { ok: false, error: "Comentário não encontrado." };
    const session = await getServerSession(authOptions);
    if (!session?.user) return { ok: false, error: "Não autenticado." };
    // só autor ou ADMIN podem apagar
    if (c.authorId !== session.user.id && session.user.role !== "ADMIN") {
      return { ok: false, error: "Sem permissão para apagar este comentário." };
    }
    await prisma.comment.delete({ where: { id } });
    const projectId =
      c.issue?.projectId ?? c.risk?.projectId ?? c.action?.projectId ?? c.changeRequest?.projectId ?? c.task?.projectId;
    if (projectId) revalidateProject(projectId);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro" };
  }
}
