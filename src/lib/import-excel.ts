// Importador Excel unificado — substitui clsCargaPlanilhasPMOnline.dll do legado.
// Suporta os 6 templates do legado: IS (Issues), RK (Risks), AC (Actions),
// CR (Change Requests), TK (Tasks/cronograma), TC (Traceability).
//
// Auto-detecta o template pelo conjunto de cabeçalhos. Aceita PT-BR e EN,
// case/acento-insensitive.

import * as XLSX from "xlsx";
import type { ImportTemplate } from "./enums";

// ─── Tipos por template ──────────────────────────────────────────────────

export type RowTask = {
  externalId?: string;
  wbs?: string;
  name: string;
  startDate: Date;
  endDate: Date;
  durationDays?: number;
  percentDone?: number;
  isMilestone?: boolean;
  isSummary?: boolean;
  parentExternalId?: string;
  predecessorExternalIds?: string[];
  assigneeEmail?: string;
  equipeCode?: string;
  areaCode?: string;
  frenteCode?: string;
};

export type RowIssue = {
  externalId?: string;
  title: string;
  description?: string;
  ownerEmail?: string;
  areaCode?: string;
  priority?: string;
  severity?: string;
  status?: string;
  workflow?: string;
  openedAt?: Date;
  closedAt?: Date;
};

export type RowRisk = {
  externalId?: string;
  title: string;
  description?: string;
  ownerEmail?: string;
  probability?: number;
  impact?: number;
  mitigation?: string;
  status?: string;
};

export type RowAction = {
  externalId?: string;
  title: string;
  description?: string;
  ownerEmail?: string;
  dueDate?: Date;
  status?: string;
  priority?: string;
};

export type RowChangeRequest = {
  externalId?: string;
  name: string;
  description?: string;
  ownerEmail?: string;
  comiteCode?: string;
  priority?: string;
  status?: string;
  dueDate?: Date;
};

export type RowTraceability = {
  reqId: string;
  testId?: string;
  taskExternalId?: string;
  mapping?: string;
  notes?: string;
};

export type ParsedExcel =
  | { template: "TK"; rows: RowTask[]; warnings: string[] }
  | { template: "IS"; rows: RowIssue[]; warnings: string[] }
  | { template: "RK"; rows: RowRisk[]; warnings: string[] }
  | { template: "AC"; rows: RowAction[]; warnings: string[] }
  | { template: "CR"; rows: RowChangeRequest[]; warnings: string[] }
  | { template: "TC"; rows: RowTraceability[]; warnings: string[] };

// ─── Mapeamento de cabeçalhos ────────────────────────────────────────────

type FieldMap = Record<string, string>; // header normalizado → field name

const TASK_HEADERS: FieldMap = {
  id: "externalId",
  "id externo": "externalId",
  externalid: "externalId",
  "task uid": "externalId",
  task_uid: "externalId",
  uid: "externalId",
  wbs: "wbs",
  edt: "wbs",
  nome: "name",
  tarefa: "name",
  name: "name",
  task: "name",
  "task name": "name",
  inicio: "startDate",
  início: "startDate",
  start: "startDate",
  startdate: "startDate",
  "task start date": "startDate",
  "data inicio": "startDate",
  fim: "endDate",
  termino: "endDate",
  término: "endDate",
  finish: "endDate",
  enddate: "endDate",
  "task finish date": "endDate",
  "data fim": "endDate",
  duracao: "durationDays",
  duração: "durationDays",
  duration: "durationDays",
  perc: "percentDone",
  percentual: "percentDone",
  "% concluido": "percentDone",
  "% concluído": "percentDone",
  "% complete": "percentDone",
  "task pct comp": "percentDone",
  percentdone: "percentDone",
  marco: "isMilestone",
  milestone: "isMilestone",
  resumo: "isSummary",
  summary: "isSummary",
  pai: "parentExternalId",
  "id pai": "parentExternalId",
  parent: "parentExternalId",
  predecessores: "predecessorExternalIds",
  predecessoras: "predecessorExternalIds",
  predecessors: "predecessorExternalIds",
  responsavel: "assigneeEmail",
  responsável: "assigneeEmail",
  responsable: "assigneeEmail",
  assignee: "assigneeEmail",
  "email responsavel": "assigneeEmail",
  equipe: "equipeCode",
  team: "equipeCode",
  area: "areaCode",
  área: "areaCode",
  frente: "frenteCode",
};

const ISSUE_HEADERS: FieldMap = {
  id: "externalId",
  "issue id": "externalId",
  externalid: "externalId",
  nome: "title",
  titulo: "title",
  título: "title",
  title: "title",
  name: "title",
  descricao: "description",
  descrição: "description",
  description: "description",
  responsavel: "ownerEmail",
  responsável: "ownerEmail",
  owner: "ownerEmail",
  area: "areaCode",
  área: "areaCode",
  prioridade: "priority",
  priority: "priority",
  severidade: "severity",
  severity: "severity",
  status: "status",
  "overall status": "status",
  workflow: "workflow",
  "data criacao": "openedAt",
  "data criação": "openedAt",
  "opened at": "openedAt",
  abertura: "openedAt",
  "data fechamento": "closedAt",
  fechamento: "closedAt",
  "closed at": "closedAt",
};

const RISK_HEADERS: FieldMap = {
  id: "externalId",
  "risk id": "externalId",
  externalid: "externalId",
  nome: "title",
  titulo: "title",
  título: "title",
  title: "title",
  descricao: "description",
  descrição: "description",
  description: "description",
  responsavel: "ownerEmail",
  responsável: "ownerEmail",
  owner: "ownerEmail",
  probabilidade: "probability",
  probability: "probability",
  prob: "probability",
  impacto: "impact",
  impact: "impact",
  mitigacao: "mitigation",
  mitigação: "mitigation",
  mitigation: "mitigation",
  status: "status",
};

const ACTION_HEADERS: FieldMap = {
  id: "externalId",
  "action id": "externalId",
  externalid: "externalId",
  nome: "title",
  titulo: "title",
  título: "title",
  title: "title",
  descricao: "description",
  descrição: "description",
  description: "description",
  responsavel: "ownerEmail",
  responsável: "ownerEmail",
  owner: "ownerEmail",
  prazo: "dueDate",
  "data vencimento": "dueDate",
  vencimento: "dueDate",
  duedate: "dueDate",
  "due date": "dueDate",
  status: "status",
  prioridade: "priority",
  priority: "priority",
};

const CR_HEADERS: FieldMap = {
  id: "externalId",
  "cr id": "externalId",
  externalid: "externalId",
  nome: "name",
  name: "name",
  titulo: "name",
  título: "name",
  descricao: "description",
  descrição: "description",
  description: "description",
  responsavel: "ownerEmail",
  responsável: "ownerEmail",
  owner: "ownerEmail",
  comite: "comiteCode",
  comitê: "comiteCode",
  prioridade: "priority",
  priority: "priority",
  status: "status",
  "data limite": "dueDate",
  prazo: "dueDate",
  duedate: "dueDate",
};

const TC_HEADERS: FieldMap = {
  "req id": "reqId",
  reqid: "reqId",
  requisito: "reqId",
  requirement: "reqId",
  "test id": "testId",
  testid: "testId",
  teste: "testId",
  test: "testId",
  "task id": "taskExternalId",
  taskid: "taskExternalId",
  tarefa: "taskExternalId",
  task: "taskExternalId",
  mapeamento: "mapping",
  mapping: "mapping",
  notas: "notes",
  notes: "notes",
  observacoes: "notes",
  observações: "notes",
};

// Marcadores únicos por template, para auto-detecção.
const TEMPLATE_MARKERS: Record<ImportTemplate, FieldMap> = {
  TK: TASK_HEADERS,
  IS: ISSUE_HEADERS,
  RK: RISK_HEADERS,
  AC: ACTION_HEADERS,
  CR: CR_HEADERS,
  TC: TC_HEADERS,
};

// ─── Helpers ─────────────────────────────────────────────────────────────

function normalize(s: string): string {
  return s
    .toString()
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

function parseDate(value: unknown): Date | null {
  if (value == null || value === "") return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    // Excel serial date
    const utcDays = Math.floor(value - 25569);
    const ms = utcDays * 86400 * 1000;
    return new Date(ms);
  }
  const s = value.toString().trim();
  // dd/mm/yyyy ou dd-mm-yyyy
  const br = s.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})/);
  if (br) {
    const [, d, m, y] = br;
    const year = y.length === 2 ? 2000 + Number(y) : Number(y);
    return new Date(year, Number(m) - 1, Number(d));
  }
  const iso = new Date(s);
  return isNaN(iso.getTime()) ? null : iso;
}

function parseBool(v: unknown): boolean {
  if (typeof v === "boolean") return v;
  if (v == null) return false;
  const s = v.toString().trim().toLowerCase();
  return s === "1" || s === "sim" || s === "true" || s === "yes" || s === "x";
}

function parseNumber(v: unknown): number | undefined {
  if (v == null || v === "") return undefined;
  const n = typeof v === "number" ? v : Number(v.toString().replace(",", "."));
  return Number.isFinite(n) ? n : undefined;
}

function parseList(v: unknown): string[] {
  if (v == null || v === "") return [];
  return v
    .toString()
    .split(/[,;]/)
    .map((x) => x.trim())
    .filter(Boolean);
}

// Score: quantos cabeçalhos são reconhecidos por cada template.
function detectTemplate(headers: string[]): { template: ImportTemplate; mapped: number; total: number } {
  const normHeaders = headers.map(normalize);
  let best: ImportTemplate = "TK";
  let bestScore = 0;
  for (const tpl of Object.keys(TEMPLATE_MARKERS) as ImportTemplate[]) {
    const map = TEMPLATE_MARKERS[tpl];
    let count = 0;
    for (const h of normHeaders) if (map[h]) count++;
    if (count > bestScore) {
      bestScore = count;
      best = tpl;
    }
  }
  return { template: best, mapped: bestScore, total: headers.length };
}

// Constrói mapeamento header original → field do template selecionado.
function buildHeaderMap(headers: string[], template: ImportTemplate): Record<string, string> {
  const tplMap = TEMPLATE_MARKERS[template];
  const out: Record<string, string> = {};
  for (const h of headers) {
    const field = tplMap[normalize(h)];
    if (field) out[h] = field;
  }
  return out;
}

// ─── Parsers por template ────────────────────────────────────────────────

function parseTaskRows(rows: Record<string, unknown>[], headerMap: Record<string, string>): { rows: RowTask[]; warnings: string[] } {
  const warnings: string[] = [];
  const out: RowTask[] = [];
  rows.forEach((row, idx) => {
    const t: Partial<RowTask> = {};
    for (const [src, field] of Object.entries(headerMap)) {
      const v = row[src];
      switch (field) {
        case "startDate":
        case "endDate": {
          const d = parseDate(v);
          if (d) (t as Record<string, unknown>)[field] = d;
          break;
        }
        case "durationDays":
        case "percentDone":
          (t as Record<string, unknown>)[field] = parseNumber(v);
          break;
        case "isMilestone":
        case "isSummary":
          (t as Record<string, unknown>)[field] = parseBool(v);
          break;
        case "predecessorExternalIds":
          (t as Record<string, unknown>)[field] = parseList(v);
          break;
        default:
          (t as Record<string, unknown>)[field] = v ? v.toString().trim() : undefined;
      }
    }
    if (!t.name) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem nome).`);
      return;
    }
    if (!t.startDate || !t.endDate) {
      warnings.push(`Linha ${idx + 2} ("${t.name}"): datas ausentes/inválidas.`);
      return;
    }
    if (t.endDate < t.startDate) {
      warnings.push(`Linha ${idx + 2} ("${t.name}"): fim < início (será mantido).`);
    }
    out.push(t as RowTask);
  });
  return { rows: out, warnings };
}

function parseIssueRows(rows: Record<string, unknown>[], headerMap: Record<string, string>): { rows: RowIssue[]; warnings: string[] } {
  const warnings: string[] = [];
  const out: RowIssue[] = [];
  rows.forEach((row, idx) => {
    const r: Partial<RowIssue> = {};
    for (const [src, field] of Object.entries(headerMap)) {
      const v = row[src];
      if (field === "openedAt" || field === "closedAt") {
        const d = parseDate(v);
        if (d) (r as Record<string, unknown>)[field] = d;
      } else {
        (r as Record<string, unknown>)[field] = v ? v.toString().trim() : undefined;
      }
    }
    if (!r.title) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem título).`);
      return;
    }
    out.push(r as RowIssue);
  });
  return { rows: out, warnings };
}

function parseRiskRows(rows: Record<string, unknown>[], headerMap: Record<string, string>): { rows: RowRisk[]; warnings: string[] } {
  const warnings: string[] = [];
  const out: RowRisk[] = [];
  rows.forEach((row, idx) => {
    const r: Partial<RowRisk> = {};
    for (const [src, field] of Object.entries(headerMap)) {
      const v = row[src];
      if (field === "probability" || field === "impact") {
        let n = parseNumber(v);
        if (n !== undefined && n > 1) n = n / (n > 10 ? 100 : 10); // 0..10 ou 0..100 → 0..1
        (r as Record<string, unknown>)[field] = n;
      } else {
        (r as Record<string, unknown>)[field] = v ? v.toString().trim() : undefined;
      }
    }
    if (!r.title) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem título).`);
      return;
    }
    out.push(r as RowRisk);
  });
  return { rows: out, warnings };
}

function parseActionRows(rows: Record<string, unknown>[], headerMap: Record<string, string>): { rows: RowAction[]; warnings: string[] } {
  const warnings: string[] = [];
  const out: RowAction[] = [];
  rows.forEach((row, idx) => {
    const r: Partial<RowAction> = {};
    for (const [src, field] of Object.entries(headerMap)) {
      const v = row[src];
      if (field === "dueDate") {
        const d = parseDate(v);
        if (d) (r as Record<string, unknown>)[field] = d;
      } else {
        (r as Record<string, unknown>)[field] = v ? v.toString().trim() : undefined;
      }
    }
    if (!r.title) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem título).`);
      return;
    }
    out.push(r as RowAction);
  });
  return { rows: out, warnings };
}

function parseCRRows(rows: Record<string, unknown>[], headerMap: Record<string, string>): { rows: RowChangeRequest[]; warnings: string[] } {
  const warnings: string[] = [];
  const out: RowChangeRequest[] = [];
  rows.forEach((row, idx) => {
    const r: Partial<RowChangeRequest> = {};
    for (const [src, field] of Object.entries(headerMap)) {
      const v = row[src];
      if (field === "dueDate") {
        const d = parseDate(v);
        if (d) (r as Record<string, unknown>)[field] = d;
      } else {
        (r as Record<string, unknown>)[field] = v ? v.toString().trim() : undefined;
      }
    }
    if (!r.name) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem nome).`);
      return;
    }
    out.push(r as RowChangeRequest);
  });
  return { rows: out, warnings };
}

function parseTCRows(rows: Record<string, unknown>[], headerMap: Record<string, string>): { rows: RowTraceability[]; warnings: string[] } {
  const warnings: string[] = [];
  const out: RowTraceability[] = [];
  rows.forEach((row, idx) => {
    const r: Partial<RowTraceability> = {};
    for (const [src, field] of Object.entries(headerMap)) {
      const v = row[src];
      (r as Record<string, unknown>)[field] = v ? v.toString().trim() : undefined;
    }
    if (!r.reqId) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem reqId).`);
      return;
    }
    out.push(r as RowTraceability);
  });
  return { rows: out, warnings };
}

// ─── Entry point ────────────────────────────────────────────────────────

export function parseExcel(buffer: Buffer, hint?: ImportTemplate): ParsedExcel {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) {
    return { template: hint ?? "TK", rows: [], warnings: ["Planilha vazia."] } as ParsedExcel;
  }
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { raw: true, defval: "" });
  if (rows.length === 0) {
    return { template: hint ?? "TK", rows: [], warnings: ["Sem linhas."] } as ParsedExcel;
  }

  const headers = Object.keys(rows[0]);
  const detected = hint ?? detectTemplate(headers).template;
  const headerMap = buildHeaderMap(headers, detected);

  const baseWarnings: string[] = [];
  const recognized = Object.keys(headerMap).length;
  if (recognized === 0) {
    baseWarnings.push(
      `Nenhum cabeçalho reconhecido para o template ${detected}. ` +
      `Cabeçalhos encontrados: ${headers.join(", ")}.`,
    );
  }

  switch (detected) {
    case "TK": {
      const { rows: r, warnings } = parseTaskRows(rows, headerMap);
      return { template: "TK", rows: r, warnings: [...baseWarnings, ...warnings] };
    }
    case "IS": {
      const { rows: r, warnings } = parseIssueRows(rows, headerMap);
      return { template: "IS", rows: r, warnings: [...baseWarnings, ...warnings] };
    }
    case "RK": {
      const { rows: r, warnings } = parseRiskRows(rows, headerMap);
      return { template: "RK", rows: r, warnings: [...baseWarnings, ...warnings] };
    }
    case "AC": {
      const { rows: r, warnings } = parseActionRows(rows, headerMap);
      return { template: "AC", rows: r, warnings: [...baseWarnings, ...warnings] };
    }
    case "CR": {
      const { rows: r, warnings } = parseCRRows(rows, headerMap);
      return { template: "CR", rows: r, warnings: [...baseWarnings, ...warnings] };
    }
    case "TC": {
      const { rows: r, warnings } = parseTCRows(rows, headerMap);
      return { template: "TC", rows: r, warnings: [...baseWarnings, ...warnings] };
    }
  }
}

// ─── Geração de templates vazios ────────────────────────────────────────

const TEMPLATE_HEADERS: Record<ImportTemplate, string[]> = {
  TK: [
    "ID", "WBS", "Nome", "Início", "Fim", "Duração", "% Concluído",
    "Marco", "Resumo", "Pai", "Predecessores", "Responsável", "Equipe", "Área", "Frente",
  ],
  IS: [
    "ID", "Título", "Descrição", "Responsável", "Área",
    "Prioridade", "Severidade", "Status", "Workflow",
    "Data Criação", "Data Fechamento",
  ],
  RK: [
    "ID", "Título", "Descrição", "Responsável",
    "Probabilidade", "Impacto", "Mitigação", "Status",
  ],
  AC: [
    "ID", "Título", "Descrição", "Responsável",
    "Prazo", "Status", "Prioridade",
  ],
  CR: [
    "ID", "Nome", "Descrição", "Responsável", "Comitê",
    "Prioridade", "Status", "Data Limite",
  ],
  TC: [
    "Req ID", "Test ID", "Task ID", "Mapeamento", "Notas",
  ],
};

export function buildTemplateXlsx(template: ImportTemplate): Buffer {
  const headers = TEMPLATE_HEADERS[template];
  const ws = XLSX.utils.aoa_to_sheet([headers]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, template);
  const out = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  return out as Buffer;
}
