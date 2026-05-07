"use server";

import { z } from "zod";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { hasAccess } from "@/lib/access";
import {
  ModuleSchema,
  ImportTemplateSchema,
  ImportModeSchema,
  type ImportTemplate,
  type ImportMode,
} from "@/lib/enums";
import {
  parseExcel,
  type ParsedExcel,
  type RowTask,
  type RowIssue,
  type RowRisk,
  type RowAction,
  type RowChangeRequest,
  type RowTraceability,
} from "@/lib/import-excel";
import { parseMSProjectXml } from "@/lib/import-msproject";

export type PreviewResult =
  | {
      ok: true;
      template: ImportTemplate;
      source: "EXCEL" | "MSPROJECT_XML";
      rowCount: number;
      sample: unknown[];
      warnings: string[];
      payload: string; // JSON serializado para reusar no confirm
    }
  | { ok: false; error: string };

export type ConfirmResult = { ok: true; importLogId: string; recordCount: number } | { ok: false; error: string };

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

// ─── Preview ─────────────────────────────────────────────────────────────

const PreviewInputSchema = z.object({
  projectId: z.string(),
  filename: z.string(),
  source: z.enum(["EXCEL", "MSPROJECT_XML"]),
  templateHint: ImportTemplateSchema.optional(),
  base64: z.string(), // arquivo em base64
});

export async function previewImport(input: unknown): Promise<PreviewResult> {
  try {
    const parsed = PreviewInputSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.message };
    const { projectId, source, templateHint, base64, filename } = parsed.data;

    await requireWriteOnProject(projectId);

    const buffer = Buffer.from(base64, "base64");

    if (source === "EXCEL") {
      const parsedXls = parseExcel(buffer, templateHint);
      return {
        ok: true,
        template: parsedXls.template,
        source: "EXCEL",
        rowCount: parsedXls.rows.length,
        sample: parsedXls.rows.slice(0, 50),
        warnings: parsedXls.warnings,
        payload: JSON.stringify({ template: parsedXls.template, rows: parsedXls.rows, filename }),
      };
    }

    if (source === "MSPROJECT_XML") {
      const xml = buffer.toString("utf-8");
      const r = parseMSProjectXml(xml);
      return {
        ok: true,
        template: "TK",
        source: "MSPROJECT_XML",
        rowCount: r.rows.length,
        sample: r.rows.slice(0, 50),
        warnings: r.warnings,
        payload: JSON.stringify({ template: "TK", rows: r.rows, filename }),
      };
    }

    return { ok: false, error: "Origem não suportada." };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}

// ─── Confirm ─────────────────────────────────────────────────────────────

const ConfirmInputSchema = z.object({
  projectId: z.string(),
  source: z.enum(["EXCEL", "MSPROJECT_XML"]),
  mode: ImportModeSchema,
  payload: z.string(),
});

export async function confirmImport(input: unknown): Promise<ConfirmResult> {
  try {
    const parsed = ConfirmInputSchema.safeParse(input);
    if (!parsed.success) return { ok: false, error: parsed.error.message };
    const { projectId, source, mode, payload } = parsed.data;
    const { userId } = await requireWriteOnProject(projectId);

    const data = JSON.parse(payload) as { template: ImportTemplate; rows: unknown[]; filename: string };
    const template = data.template;
    const rows = data.rows;

    const warnings: string[] = [];
    let recordCount = 0;

    switch (template) {
      case "TK":
        recordCount = await importTasks(projectId, rows as RowTask[], mode, userId, warnings);
        break;
      case "IS":
        recordCount = await importIssues(projectId, rows as RowIssue[], mode, warnings);
        break;
      case "RK":
        recordCount = await importRisks(projectId, rows as RowRisk[], mode, warnings);
        break;
      case "AC":
        recordCount = await importActions(projectId, rows as RowAction[], mode, warnings);
        break;
      case "CR":
        recordCount = await importChangeRequests(projectId, rows as RowChangeRequest[], mode, warnings);
        break;
      case "TC":
        recordCount = await importTraceability(projectId, rows as RowTraceability[], mode, warnings);
        break;
    }

    const log = await prisma.importLog.create({
      data: {
        projectId,
        userId,
        source,
        template,
        filename: data.filename,
        recordCount,
        warnings: warnings.length ? warnings.join("\n") : null,
        mode,
      },
      select: { id: true },
    });

    revalidatePath(`/projects/${projectId}/tasks`);
    revalidatePath(`/projects/${projectId}/gantt`);
    revalidatePath(`/projects/${projectId}/import`);
    revalidatePath(`/projects/${projectId}/import/history`);
    return { ok: true, importLogId: log.id, recordCount };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : "Erro desconhecido" };
  }
}

// ─── Importadores específicos ────────────────────────────────────────────

// TK: dois passos — cria/atualiza tarefas, depois resolve parent + dependências.
async function importTasks(projectId: string, rows: RowTask[], mode: ImportMode, userId: string, warnings: string[]): Promise<number> {
  if (mode === "REPLACE") {
    await prisma.task.deleteMany({ where: { projectId } });
  }

  // resolução de lookups: assignee por email, equipe/área/frente por code
  const [users, equipes, areas, frentes] = await Promise.all([
    prisma.user.findMany({ select: { id: true, email: true } }),
    prisma.equipe.findMany({ select: { id: true, code: true } }),
    prisma.area.findMany({ select: { id: true, code: true } }),
    prisma.frente.findMany({ select: { id: true, code: true } }),
  ]);
  const userByEmail = new Map(users.map((u) => [u.email.toLowerCase(), u.id]));
  const equipeByCode = new Map(equipes.map((e) => [e.code.toUpperCase(), e.id]));
  const areaByCode = new Map(areas.map((a) => [a.code.toUpperCase(), a.id]));
  const frenteByCode = new Map(frentes.map((f) => [f.code.toUpperCase(), f.id]));

  // 1ª passada: upsert por externalId
  const idMap = new Map<string, string>(); // externalId → DB id
  let count = 0;
  for (const r of rows) {
    if (!r.externalId) {
      // sem externalId: gera artificial (timestamp + index) p/ permitir match parent
      r.externalId = `import-${Date.now()}-${count}`;
    }
    const baseData = {
      projectId,
      externalId: r.externalId,
      wbs: r.wbs ?? null,
      name: r.name,
      startDate: new Date(r.startDate),
      endDate: new Date(r.endDate),
      durationDays: r.durationDays ?? Math.max(1, Math.round((new Date(r.endDate).getTime() - new Date(r.startDate).getTime()) / 86400000)),
      percentDone: r.percentDone ?? 0,
      isMilestone: !!r.isMilestone,
      isSummary: !!r.isSummary,
      assigneeId: r.assigneeEmail ? userByEmail.get(r.assigneeEmail.toLowerCase()) ?? null : null,
      equipeId: r.equipeCode ? equipeByCode.get(r.equipeCode.toUpperCase()) ?? null : null,
      areaId: r.areaCode ? areaByCode.get(r.areaCode.toUpperCase()) ?? null : null,
      frenteId: r.frenteCode ? frenteByCode.get(r.frenteCode.toUpperCase()) ?? null : null,
      status: ((r.percentDone ?? 0) >= 100 ? "DONE" : (r.percentDone ?? 0) > 0 ? "IN_PROGRESS" : "NOT_STARTED") as
        | "DONE" | "IN_PROGRESS" | "NOT_STARTED",
    };

    if (mode === "REPLACE") {
      const created = await prisma.task.create({ data: baseData, select: { id: true } });
      idMap.set(r.externalId, created.id);
      await prisma.taskHistory.create({ data: { taskId: created.id, userId, action: "IMPORTED", note: "Importado (REPLACE)" } });
      count++;
      continue;
    }

    // MERGE / ONLY_NEW: localiza por externalId
    const existing = await prisma.task.findFirst({
      where: { projectId, externalId: r.externalId },
      select: { id: true },
    });
    if (!existing) {
      const created = await prisma.task.create({ data: baseData, select: { id: true } });
      idMap.set(r.externalId, created.id);
      await prisma.taskHistory.create({ data: { taskId: created.id, userId, action: "IMPORTED", note: `Importado (${mode})` } });
      count++;
    } else if (mode === "MERGE") {
      await prisma.task.update({ where: { id: existing.id }, data: baseData });
      idMap.set(r.externalId, existing.id);
      await prisma.taskHistory.create({ data: { taskId: existing.id, userId, action: "IMPORTED", note: "Atualizado (MERGE)" } });
      count++;
    } else {
      // ONLY_NEW: pula
      idMap.set(r.externalId, existing.id);
    }
  }

  // 2ª passada: parents + dependências
  for (const r of rows) {
    const myId = idMap.get(r.externalId!);
    if (!myId) continue;
    if (r.parentExternalId) {
      const parentId = idMap.get(r.parentExternalId);
      if (parentId) {
        await prisma.task.update({ where: { id: myId }, data: { parentId } });
      } else {
        warnings.push(`Tarefa "${r.name}": parent ${r.parentExternalId} não encontrado.`);
      }
    }
    if (r.predecessorExternalIds?.length) {
      for (const p of r.predecessorExternalIds) {
        // remove sufixos tipo "5FS+1d" do MS Project
        const cleanId = p.replace(/[FS]{2}.*$/, "").trim();
        const predId = idMap.get(cleanId);
        if (!predId) {
          warnings.push(`Tarefa "${r.name}": predecessor ${p} não encontrado.`);
          continue;
        }
        try {
          await prisma.taskDependency.create({
            data: { predecessorId: predId, successorId: myId, type: "FS", lagDays: 0 },
          });
        } catch {
          // já existe — tudo bem
        }
      }
    }
  }

  return count;
}

async function importIssues(projectId: string, rows: RowIssue[], mode: ImportMode, warnings: string[]): Promise<number> {
  if (mode === "REPLACE") await prisma.issue.deleteMany({ where: { projectId } });
  const users = await prisma.user.findMany({ select: { id: true, email: true } });
  const userByEmail = new Map(users.map((u) => [u.email.toLowerCase(), u.id]));
  const areas = await prisma.area.findMany({ select: { id: true, code: true } });
  const areaByCode = new Map(areas.map((a) => [a.code.toUpperCase(), a.id]));

  let count = 0;
  for (const r of rows) {
    const data = {
      projectId,
      externalId: r.externalId ?? null,
      title: r.title,
      description: r.description ?? null,
      ownerId: r.ownerEmail ? userByEmail.get(r.ownerEmail.toLowerCase()) ?? null : null,
      areaId: r.areaCode ? areaByCode.get(r.areaCode.toUpperCase()) ?? null : null,
      severity: normalizeEnum(r.severity, ["LOW", "MEDIUM", "HIGH", "CRITICAL"], "MEDIUM"),
      priority: normalizeEnum(r.priority, ["LOW", "MEDIUM", "HIGH", "CRITICAL"], "MEDIUM"),
      status: normalizeEnum(r.status, ["OPEN", "IN_PROGRESS", "RESOLVED", "CLOSED"], "OPEN"),
      workflow: r.workflow ?? null,
      openedAt: r.openedAt ?? new Date(),
      closedAt: r.closedAt ?? null,
    };
    const existing = r.externalId
      ? await prisma.issue.findFirst({ where: { projectId, externalId: r.externalId }, select: { id: true } })
      : null;
    if (!existing) {
      await prisma.issue.create({ data });
      count++;
    } else if (mode === "MERGE") {
      await prisma.issue.update({ where: { id: existing.id }, data });
      count++;
    }
  }
  void warnings;
  return count;
}

async function importRisks(projectId: string, rows: RowRisk[], mode: ImportMode, warnings: string[]): Promise<number> {
  if (mode === "REPLACE") await prisma.risk.deleteMany({ where: { projectId } });
  const users = await prisma.user.findMany({ select: { id: true, email: true } });
  const userByEmail = new Map(users.map((u) => [u.email.toLowerCase(), u.id]));

  let count = 0;
  for (const r of rows) {
    const probability = r.probability ?? 0.5;
    const impact = r.impact ?? 0.5;
    const data = {
      projectId,
      externalId: r.externalId ?? null,
      title: r.title,
      description: r.description ?? null,
      ownerId: r.ownerEmail ? userByEmail.get(r.ownerEmail.toLowerCase()) ?? null : null,
      probability,
      impact,
      exposure: probability * impact,
      mitigation: r.mitigation ?? null,
      status: normalizeEnum(r.status, ["IDENTIFIED", "ANALYZING", "MITIGATING", "ACCEPTED", "CLOSED"], "IDENTIFIED"),
    };
    const existing = r.externalId
      ? await prisma.risk.findFirst({ where: { projectId, externalId: r.externalId }, select: { id: true } })
      : null;
    if (!existing) {
      await prisma.risk.create({ data });
      count++;
    } else if (mode === "MERGE") {
      await prisma.risk.update({ where: { id: existing.id }, data });
      count++;
    }
  }
  void warnings;
  return count;
}

async function importActions(projectId: string, rows: RowAction[], mode: ImportMode, warnings: string[]): Promise<number> {
  if (mode === "REPLACE") await prisma.action.deleteMany({ where: { projectId } });
  const users = await prisma.user.findMany({ select: { id: true, email: true } });
  const userByEmail = new Map(users.map((u) => [u.email.toLowerCase(), u.id]));

  let count = 0;
  for (const r of rows) {
    const data = {
      projectId,
      externalId: r.externalId ?? null,
      title: r.title,
      description: r.description ?? null,
      ownerId: r.ownerEmail ? userByEmail.get(r.ownerEmail.toLowerCase()) ?? null : null,
      dueDate: r.dueDate ?? null,
      status: normalizeEnum(r.status, ["OPEN", "IN_PROGRESS", "DONE", "CANCELLED"], "OPEN"),
      priority: normalizeEnum(r.priority, ["LOW", "MEDIUM", "HIGH", "CRITICAL"], "MEDIUM"),
    };
    const existing = r.externalId
      ? await prisma.action.findFirst({ where: { projectId, externalId: r.externalId }, select: { id: true } })
      : null;
    if (!existing) {
      await prisma.action.create({ data });
      count++;
    } else if (mode === "MERGE") {
      await prisma.action.update({ where: { id: existing.id }, data });
      count++;
    }
  }
  void warnings;
  return count;
}

async function importChangeRequests(projectId: string, rows: RowChangeRequest[], mode: ImportMode, warnings: string[]): Promise<number> {
  if (mode === "REPLACE") await prisma.changeRequest.deleteMany({ where: { projectId } });
  const users = await prisma.user.findMany({ select: { id: true, email: true } });
  const userByEmail = new Map(users.map((u) => [u.email.toLowerCase(), u.id]));
  const comites = await prisma.comite.findMany({ select: { id: true, code: true } });
  const comiteByCode = new Map(comites.map((c) => [c.code.toUpperCase(), c.id]));

  let count = 0;
  for (const r of rows) {
    const data = {
      projectId,
      externalId: r.externalId ?? null,
      name: r.name,
      description: r.description ?? null,
      ownerId: r.ownerEmail ? userByEmail.get(r.ownerEmail.toLowerCase()) ?? null : null,
      comiteId: r.comiteCode ? comiteByCode.get(r.comiteCode.toUpperCase()) ?? null : null,
      priority: normalizeEnum(r.priority, ["LOW", "MEDIUM", "HIGH", "CRITICAL"], "MEDIUM"),
      status: normalizeEnum(
        r.status,
        ["OPEN", "UNDER_REVIEW", "APPROVED", "IN_IMPLEMENTATION", "IMPLEMENTED", "REJECTED", "CANCELLED"],
        "OPEN",
      ),
      dueDate: r.dueDate ?? null,
    };
    const existing = r.externalId
      ? await prisma.changeRequest.findFirst({ where: { projectId, externalId: r.externalId }, select: { id: true } })
      : null;
    if (!existing) {
      await prisma.changeRequest.create({ data });
      count++;
    } else if (mode === "MERGE") {
      await prisma.changeRequest.update({ where: { id: existing.id }, data });
      count++;
    }
  }
  void warnings;
  return count;
}

async function importTraceability(projectId: string, rows: RowTraceability[], mode: ImportMode, warnings: string[]): Promise<number> {
  if (mode === "REPLACE") await prisma.traceability.deleteMany({ where: { projectId } });
  const taskByExt = new Map(
    (await prisma.task.findMany({ where: { projectId }, select: { id: true, externalId: true } }))
      .filter((t) => t.externalId)
      .map((t) => [t.externalId!, t.id]),
  );

  let count = 0;
  for (const r of rows) {
    const taskId = r.taskExternalId ? taskByExt.get(r.taskExternalId) ?? null : null;
    if (r.taskExternalId && !taskId) {
      warnings.push(`Traceability ${r.reqId}: task ${r.taskExternalId} não encontrada.`);
    }
    await prisma.traceability.create({
      data: {
        projectId,
        reqId: r.reqId,
        testId: r.testId ?? null,
        taskId,
        mapping: r.mapping ?? null,
        notes: r.notes ?? null,
      },
    });
    count++;
  }
  return count;
}

function normalizeEnum<T extends string>(value: string | undefined, allowed: T[], fallback: T): T {
  if (!value) return fallback;
  const up = value.toUpperCase().replace(/\s+/g, "_");
  return (allowed as string[]).includes(up) ? (up as T) : fallback;
}
