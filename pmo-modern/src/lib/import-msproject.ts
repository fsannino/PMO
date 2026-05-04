// Importador MS Project XML — substitui as queries .dqy do legado.
// Aceita o formato Project XML (export "Save As" do MS Project Desktop ou
// Project Server, schema 2003+).
//
// Mapeia <Tasks><Task><UID>, <Name>, <Start>, <Finish>, <PercentComplete>,
// <Milestone>, <Summary>, <OutlineLevel>, <OutlineNumber>,
// <PredecessorLink>{<PredecessorUID>, <Type>, <LinkLag>}.

import { XMLParser } from "fast-xml-parser";
import type { RowTask } from "./import-excel";

type RawPredecessor = {
  PredecessorUID?: string | number;
  Type?: string | number;
  LinkLag?: string | number;
};
type RawTask = {
  UID?: string | number;
  ID?: string | number;
  Name?: string;
  Start?: string;
  Finish?: string;
  Duration?: string; // formato PT.. ou número
  PercentComplete?: string | number;
  Milestone?: string | number;
  Summary?: string | number;
  OutlineLevel?: string | number;
  OutlineNumber?: string;
  PredecessorLink?: RawPredecessor | RawPredecessor[];
};

export type MspParseResult = {
  rows: RowTask[];
  warnings: string[];
};

const DURATION_RE = /^PT(\d+)H/;

function toBool(v: unknown): boolean {
  if (v == null) return false;
  if (typeof v === "boolean") return v;
  const s = String(v).trim().toLowerCase();
  return s === "1" || s === "true" || s === "yes";
}

function toNumber(v: unknown): number | undefined {
  if (v == null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function parseISO(s: unknown): Date | null {
  if (!s) return null;
  const d = new Date(String(s));
  return isNaN(d.getTime()) ? null : d;
}

function durationDaysFromMSP(s: unknown): number | undefined {
  if (!s) return undefined;
  const m = String(s).match(DURATION_RE);
  if (!m) return undefined;
  const hours = Number(m[1]);
  // MS Project assume 8h por dia útil (configurável). Usamos como aproximação.
  return Math.round(hours / 8);
}

function asArray<T>(x: T | T[] | undefined): T[] {
  if (x == null) return [];
  return Array.isArray(x) ? x : [x];
}

export function parseMSProjectXml(xml: string): MspParseResult {
  const parser = new XMLParser({
    ignoreAttributes: true,
    parseTagValue: false,
    trimValues: true,
  });
  const data = parser.parse(xml) as Record<string, unknown>;
  const project = data.Project as Record<string, unknown> | undefined;
  if (!project) return { rows: [], warnings: ["XML não tem nó <Project> raiz."] };

  const tasksWrap = project.Tasks as { Task?: RawTask | RawTask[] } | undefined;
  const rawTasks = asArray<RawTask>(tasksWrap?.Task);
  if (rawTasks.length === 0) return { rows: [], warnings: ["Nenhuma <Task> encontrada."] };

  const warnings: string[] = [];
  // mapeia UID para WBS para resolver predecessores como WBS quando possível
  const uidToOutline = new Map<string, string>();
  for (const t of rawTasks) {
    const uid = t.UID != null ? String(t.UID) : null;
    if (uid && t.OutlineNumber) uidToOutline.set(uid, String(t.OutlineNumber));
  }

  // Determina parent por OutlineLevel: o pai mais recente com nível < atual.
  const stack: { level: number; uid: string }[] = [];
  const uidToParent = new Map<string, string>();
  for (const t of rawTasks) {
    const uid = t.UID != null ? String(t.UID) : null;
    const lvl = toNumber(t.OutlineLevel) ?? 1;
    if (!uid) continue;
    while (stack.length && stack[stack.length - 1].level >= lvl) stack.pop();
    if (stack.length) uidToParent.set(uid, stack[stack.length - 1].uid);
    stack.push({ level: lvl, uid });
  }

  const rows: RowTask[] = [];
  for (const t of rawTasks) {
    const uid = t.UID != null ? String(t.UID) : null;
    if (!uid) {
      warnings.push("Tarefa sem UID — ignorada.");
      continue;
    }
    // MS Project costuma incluir uma "Project Summary Task" com UID=0 sem nome
    const name = (t.Name ?? "").trim();
    if (!name) {
      warnings.push(`UID ${uid}: ignorada (sem Name).`);
      continue;
    }
    const start = parseISO(t.Start);
    const end = parseISO(t.Finish);
    if (!start || !end) {
      warnings.push(`UID ${uid} ("${name}"): datas inválidas.`);
      continue;
    }

    const predecessors = asArray<RawPredecessor>(t.PredecessorLink)
      .map((p) => p.PredecessorUID != null ? String(p.PredecessorUID) : null)
      .filter((x): x is string => !!x);

    const parentUid = uidToParent.get(uid);

    rows.push({
      externalId: uid,
      wbs: t.OutlineNumber ? String(t.OutlineNumber) : undefined,
      name,
      startDate: start,
      endDate: end,
      durationDays: durationDaysFromMSP(t.Duration),
      percentDone: toNumber(t.PercentComplete) ?? 0,
      isMilestone: toBool(t.Milestone),
      isSummary: toBool(t.Summary),
      parentExternalId: parentUid,
      predecessorExternalIds: predecessors,
    });
  }

  return { rows, warnings };
}
