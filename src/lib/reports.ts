// Builders de dados dos 6 relatórios padrão. Funções puras: dado um
// projectId (e filtros opcionais), retorna estruturas serializáveis usadas
// pelos exporters XLSX/PDF e pelo preview HTML.

import { prisma } from "./db";
import { projectScorecard } from "./scorecard";

const DAY_MS = 86400000;

function fmtDate(d: Date | null | undefined): string {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("pt-BR");
}

// ─── 1. Detalhado ────────────────────────────────────────────────────────
export type ReportDetalhadoRow = {
  wbs: string;
  name: string;
  start: string;
  end: string;
  duration: number;
  percent: number;
  status: string;
  assignee: string;
  equipe: string;
  area: string;
  daysLate: number;
};

export async function reportDetalhado(projectId: string): Promise<ReportDetalhadoRow[]> {
  const today = new Date();
  const tasks = await prisma.task.findMany({
    where: { projectId, deletedAt: null },
    include: {
      assignee: { select: { name: true } },
      equipe: { select: { code: true } },
      area: { select: { code: true } },
    },
    orderBy: [{ wbs: "asc" }, { startDate: "asc" }],
  });
  return tasks.map((t) => ({
    wbs: t.wbs ?? "—",
    name: t.name,
    start: fmtDate(t.startDate),
    end: fmtDate(t.endDate),
    duration: t.durationDays ?? 0,
    percent: Math.round(t.percentDone),
    status: t.status,
    assignee: t.assignee?.name ?? "—",
    equipe: t.equipe?.code ?? "—",
    area: t.area?.code ?? "—",
    daysLate: t.percentDone < 100 ? Math.max(0, Math.floor((today.getTime() - t.endDate.getTime()) / DAY_MS)) : 0,
  }));
}

// ─── 2. Consolidado por equipe ───────────────────────────────────────────
export type ReportConsolidadoRow = {
  equipe: string;
  total: number;
  done: number;
  inProgress: number;
  notStarted: number;
  late: number;
  workDays: number;
  avgPercent: number;
};

export async function reportConsolidado(projectId: string): Promise<ReportConsolidadoRow[]> {
  const tasks = await prisma.task.findMany({
    where: { projectId, deletedAt: null, isSummary: false },
    include: { equipe: { select: { code: true } } },
  });
  const today = new Date();
  const map = new Map<string, ReportConsolidadoRow>();
  for (const t of tasks) {
    const key = t.equipe?.code ?? "—";
    const cur = map.get(key) ?? { equipe: key, total: 0, done: 0, inProgress: 0, notStarted: 0, late: 0, workDays: 0, avgPercent: 0 };
    cur.total++;
    cur.workDays += t.durationDays ?? 0;
    cur.avgPercent += t.percentDone;
    if (t.percentDone >= 100) cur.done++;
    else if (t.percentDone > 0) cur.inProgress++;
    else cur.notStarted++;
    if (t.percentDone < 100 && t.endDate < today) cur.late++;
    map.set(key, cur);
  }
  return [...map.values()].map((r) => ({
    ...r,
    avgPercent: r.total ? Math.round(r.avgPercent / r.total) : 0,
  })).sort((a, b) => a.equipe.localeCompare(b.equipe));
}

// ─── 3. Criticidade — Issues + Risks priorizados ─────────────────────────
export type ReportCriticidadeRow = {
  type: "ISSUE" | "RISK";
  title: string;
  level: string; // severity / exposure
  status: string;
  owner: string;
  metric: number; // ordenação numérica
};

export async function reportCriticidade(projectId: string): Promise<ReportCriticidadeRow[]> {
  const SEV_RANK: Record<string, number> = { CRITICAL: 4, HIGH: 3, MEDIUM: 2, LOW: 1 };
  const [issues, risks] = await Promise.all([
    prisma.issue.findMany({
      where: { projectId, deletedAt: null, status: { in: ["OPEN", "IN_PROGRESS"] } },
      include: { owner: { select: { name: true } } },
    }),
    prisma.risk.findMany({
      where: { projectId, deletedAt: null, status: { in: ["IDENTIFIED", "ANALYZING", "MITIGATING"] } },
      include: { owner: { select: { name: true } } },
    }),
  ]);
  const rows: ReportCriticidadeRow[] = [];
  for (const i of issues) {
    rows.push({
      type: "ISSUE",
      title: i.title,
      level: i.severity,
      status: i.status,
      owner: i.owner?.name ?? "—",
      metric: SEV_RANK[i.severity] ?? 0,
    });
  }
  for (const r of risks) {
    rows.push({
      type: "RISK",
      title: r.title,
      level: r.exposure.toFixed(2),
      status: r.status,
      owner: r.owner?.name ?? "—",
      metric: r.exposure * 5, // escala equivalente a sev
    });
  }
  return rows.sort((a, b) => b.metric - a.metric);
}

// ─── 4. Comentários ──────────────────────────────────────────────────────
export type ReportCommentRow = {
  date: string;
  author: string;
  target: string;
  type: string;
  body: string;
};

export async function reportComentarios(projectId: string): Promise<ReportCommentRow[]> {
  const comments = await prisma.comment.findMany({
    where: {
      OR: [
        { task: { projectId } },
        { issue: { projectId } },
        { risk: { projectId } },
        { action: { projectId } },
        { changeRequest: { projectId } },
      ],
    },
    include: {
      author: { select: { name: true } },
      task: { select: { name: true } },
      issue: { select: { title: true } },
      risk: { select: { title: true } },
      action: { select: { title: true } },
      changeRequest: { select: { name: true } },
    },
    orderBy: { createdAt: "desc" },
  });
  return comments.map((c) => {
    const target = c.task?.name ?? c.issue?.title ?? c.risk?.title ?? c.action?.title ?? c.changeRequest?.name ?? "—";
    const type =
      c.task ? "Tarefa" : c.issue ? "Issue" : c.risk ? "Risco" : c.action ? "Action" : c.changeRequest ? "CR" : "?";
    return {
      date: new Date(c.createdAt).toLocaleString("pt-BR"),
      author: c.author.name,
      target,
      type,
      body: c.body,
    };
  });
}

// ─── 5. Issues KPI por mês ───────────────────────────────────────────────
export type ReportIssuesKPIRow = {
  month: string; // YYYY-MM
  opened: number;
  closed: number;
  open: number; // saldo até este mês
};

export async function reportIssuesKPI(projectId: string): Promise<ReportIssuesKPIRow[]> {
  const issues = await prisma.issue.findMany({
    where: { projectId },
    select: { openedAt: true, closedAt: true, status: true, deletedAt: true },
  });
  const map = new Map<string, { opened: number; closed: number }>();
  function bucket(d: Date) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }
  for (const i of issues) {
    const k = bucket(i.openedAt);
    const cur = map.get(k) ?? { opened: 0, closed: 0 };
    cur.opened++;
    map.set(k, cur);
    if (i.closedAt) {
      const k2 = bucket(i.closedAt);
      const c2 = map.get(k2) ?? { opened: 0, closed: 0 };
      c2.closed++;
      map.set(k2, c2);
    }
  }
  const months = [...map.keys()].sort();
  let runningOpen = 0;
  return months.map((m) => {
    const v = map.get(m)!;
    runningOpen += v.opened - v.closed;
    return { month: m, opened: v.opened, closed: v.closed, open: runningOpen };
  });
}

// ─── 6. Tarefas atrasadas ────────────────────────────────────────────────
export type ReportAtrasadasRow = {
  wbs: string;
  name: string;
  end: string;
  daysLate: number;
  percent: number;
  assignee: string;
  equipe: string;
};

export async function reportAtrasadas(projectId: string): Promise<ReportAtrasadasRow[]> {
  const today = new Date();
  const tasks = await prisma.task.findMany({
    where: {
      projectId,
      deletedAt: null,
      isSummary: false,
      endDate: { lt: today },
      percentDone: { lt: 100 },
    },
    include: {
      assignee: { select: { name: true } },
      equipe: { select: { code: true } },
    },
    orderBy: [{ endDate: "asc" }],
  });
  return tasks.map((t) => ({
    wbs: t.wbs ?? "—",
    name: t.name,
    end: fmtDate(t.endDate),
    daysLate: Math.floor((today.getTime() - t.endDate.getTime()) / DAY_MS),
    percent: Math.round(t.percentDone),
    assignee: t.assignee?.name ?? "—",
    equipe: t.equipe?.code ?? "—",
  }));
}

// ─── Header padrão para todos os relatórios ─────────────────────────────
export type ReportHeader = {
  projectCode: string;
  projectName: string;
  generatedAt: string;
  generatedBy: string;
  scorecard: { andamentoPct: number; totalTasks: number; late: number };
};

export async function reportHeader(projectId: string, generatedBy: string): Promise<ReportHeader | null> {
  const project = await prisma.project.findUnique({
    where: { id: projectId },
    select: { code: true, name: true },
  });
  if (!project) return null;
  const sc = await projectScorecard(projectId);
  return {
    projectCode: project.code,
    projectName: project.name,
    generatedAt: new Date().toLocaleString("pt-BR"),
    generatedBy,
    scorecard: { andamentoPct: sc.andamentoPct, totalTasks: sc.totalTasks, late: sc.delayed0to10 + sc.delayedMore10 },
  };
}

// ─── Tipo de relatório (id) ─────────────────────────────────────────────
export const REPORTS = {
  detalhado: { label: "Detalhado", desc: "Todas as tarefas com %, status, responsáveis e atraso." },
  consolidado: { label: "Consolidado por equipe", desc: "Resumo de tarefas por equipe." },
  criticidade: { label: "Criticidade", desc: "Issues + Risks ordenados por severidade/exposição." },
  comentarios: { label: "Comentários", desc: "Todas as threads de discussão do projeto." },
  "issues-kpi": { label: "Issues KPI", desc: "Aberturas, fechamentos e saldo por mês." },
  atrasadas: { label: "Tarefas atrasadas", desc: "Tarefas vencidas com menos de 100% concluído." },
} as const;
export type ReportId = keyof typeof REPORTS;
