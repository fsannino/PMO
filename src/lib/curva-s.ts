// Curva S — calcula séries de progresso planejado vs realizado por período.
// Substitui as views DadosFlagCurvaS / FlagCurva e as functions
// RetornaBCWP/RetornaACWP do legado.
//
// Definições (ponderação por durationDays):
//
//  PLANEJADO acumulado em D:
//    Σ (task.durationDays × planejamentoNoDia(task, D))
//    onde planejamentoNoDia interpola linear entre baselineStart..baselineEnd
//    (cai para startDate..endDate se baseline ausente).
//
//  REALIZADO acumulado em D:
//    Σ (task.durationDays × percentDoneEm(task, D))
//    onde percentDoneEm é o último Measurement confirmado até D, ou 0.
//
//  BCWP = trabalho ganho até D = REALIZADO acumulado.
//  BCWS = trabalho planejado até D = PLANEJADO acumulado.
//  Variação de prazo SV = BCWP - BCWS (positivo = adiantado).
//
// Tudo normalizado para % do total (somatório de durationDays do projeto).

import { prisma } from "./db";

export type Granularity = "week" | "month";

export type SerieRow = {
  bucket: string;             // "2026-W08" ou "2026-02"
  date: Date;                 // data de início do bucket
  planejadoPct: number;       // 0..100
  realizadoPct: number;       // 0..100
  sv: number;                 // realizado - planejado (em %)
};

export type CurvaSResult = {
  granularity: Granularity;
  totalWork: number;
  series: SerieRow[];
  rangeStart: Date;
  rangeEnd: Date;
};

const DAY_MS = 86400000;

function dateOnly(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function startOfWeek(d: Date): Date {
  // segunda-feira como início
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const r = new Date(d);
  r.setDate(r.getDate() + diff);
  return dateOnly(r);
}
function startOfMonth(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}
function nextBucket(d: Date, g: Granularity): Date {
  const r = new Date(d);
  if (g === "week") r.setDate(r.getDate() + 7);
  else r.setMonth(r.getMonth() + 1);
  return r;
}
function bucketLabel(d: Date, g: Granularity): string {
  if (g === "month") return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  // ISO week-ish (simplificada): aaaa-Www
  const onejan = new Date(d.getFullYear(), 0, 1);
  const days = Math.floor((d.getTime() - onejan.getTime()) / DAY_MS);
  const week = Math.ceil((days + onejan.getDay() + 1) / 7);
  return `${d.getFullYear()}-W${String(week).padStart(2, "0")}`;
}

export async function computeCurvaS(
  projectId: string,
  granularity: Granularity = "week",
): Promise<CurvaSResult> {
  const tasks = await prisma.task.findMany({
    where: { projectId, deletedAt: null, isSummary: false },
    select: {
      id: true,
      startDate: true,
      endDate: true,
      baselineStart: true,
      baselineEnd: true,
      durationDays: true,
      percentDone: true,
      measurements: {
        where: { confirmed: true },
        select: { confirmedAt: true, percentDone: true },
        orderBy: { confirmedAt: "asc" },
      },
    },
  });

  if (tasks.length === 0) {
    const today = dateOnly(new Date());
    return { granularity, totalWork: 0, series: [], rangeStart: today, rangeEnd: today };
  }

  // intervalo: do menor baseline/start ao maior baseline/end (com hoje incluído)
  let min = new Date(8.64e15);
  let max = new Date(-8.64e15);
  for (const t of tasks) {
    const s = new Date(t.baselineStart ?? t.startDate);
    const e = new Date(t.baselineEnd ?? t.endDate);
    if (s < min) min = s;
    if (e > max) max = e;
  }
  const today = new Date();
  if (today > max) max = today;

  const totalWork = tasks.reduce((acc, t) => acc + (t.durationDays ?? 1), 0);

  const series: SerieRow[] = [];
  let cursor = granularity === "week" ? startOfWeek(min) : startOfMonth(min);
  const limit = granularity === "week" ? startOfWeek(max) : startOfMonth(max);

  while (cursor <= limit) {
    const D = new Date(cursor.getTime() + (granularity === "week" ? 6 : 27) * DAY_MS); // fim do bucket aprox.
    let planAcc = 0;
    let realAcc = 0;
    for (const t of tasks) {
      const w = t.durationDays ?? 1;
      // planejado ponderado linear
      const bs = t.baselineStart ? new Date(t.baselineStart) : new Date(t.startDate);
      const be = t.baselineEnd ? new Date(t.baselineEnd) : new Date(t.endDate);
      const planFrac =
        D < bs ? 0 :
        D >= be ? 1 :
        (D.getTime() - bs.getTime()) / Math.max(1, be.getTime() - bs.getTime());
      planAcc += w * planFrac;

      // realizado: último measurement confirmado até D ou (se nenhum) percentDone atual
      // se hoje > D, usar measurements anteriores; se D >= hoje, considerar atual.
      let realPct = 0;
      const lastBefore = t.measurements
        .filter((m) => m.confirmedAt && m.confirmedAt <= D)
        .pop();
      if (lastBefore) {
        realPct = lastBefore.percentDone;
      } else if (D >= today) {
        realPct = t.percentDone;
      }
      realAcc += w * (realPct / 100);
    }
    const planPct = totalWork ? (planAcc / totalWork) * 100 : 0;
    const realPct = totalWork ? (realAcc / totalWork) * 100 : 0;
    series.push({
      bucket: bucketLabel(cursor, granularity),
      date: new Date(cursor),
      planejadoPct: Math.round(planPct * 10) / 10,
      realizadoPct: Math.round(realPct * 10) / 10,
      sv: Math.round((realPct - planPct) * 10) / 10,
    });
    cursor = nextBucket(cursor, granularity);
  }

  return { granularity, totalWork, series, rangeStart: min, rangeEnd: max };
}

// ─── Previsto vs Realizado por equipe ───────────────────────────────────

export type EquipeRow = {
  equipe: string;
  totalWork: number;
  planejadoPct: number;
  realizadoPct: number;
};

export async function previstoVsRealizadoPorEquipe(projectId: string, today: Date = new Date()): Promise<EquipeRow[]> {
  const tasks = await prisma.task.findMany({
    where: { projectId, deletedAt: null, isSummary: false },
    select: {
      durationDays: true,
      percentDone: true,
      baselineStart: true,
      baselineEnd: true,
      startDate: true,
      endDate: true,
      equipe: { select: { code: true, name: true } },
    },
  });
  const map = new Map<string, { totalWork: number; planAcc: number; realAcc: number }>();
  for (const t of tasks) {
    const key = t.equipe?.code ?? "—";
    const w = t.durationDays ?? 1;
    const bs = t.baselineStart ? new Date(t.baselineStart) : new Date(t.startDate);
    const be = t.baselineEnd ? new Date(t.baselineEnd) : new Date(t.endDate);
    const planFrac =
      today < bs ? 0 :
      today >= be ? 1 :
      (today.getTime() - bs.getTime()) / Math.max(1, be.getTime() - bs.getTime());
    const cur = map.get(key) ?? { totalWork: 0, planAcc: 0, realAcc: 0 };
    cur.totalWork += w;
    cur.planAcc += w * planFrac;
    cur.realAcc += w * (t.percentDone / 100);
    map.set(key, cur);
  }
  return [...map.entries()].map(([equipe, v]) => ({
    equipe,
    totalWork: v.totalWork,
    planejadoPct: v.totalWork ? Math.round((v.planAcc / v.totalWork) * 100) : 0,
    realizadoPct: v.totalWork ? Math.round((v.realAcc / v.totalWork) * 100) : 0,
  })).sort((a, b) => a.equipe.localeCompare(b.equipe));
}

// ─── Heatmap de atrasos: semana × equipe ────────────────────────────────

export type HeatCell = { week: string; equipe: string; lateCount: number };

export async function heatmapAtrasos(projectId: string, today: Date = new Date()): Promise<{
  weeks: string[];
  equipes: string[];
  cells: HeatCell[];
}> {
  const tasks = await prisma.task.findMany({
    where: {
      projectId,
      deletedAt: null,
      isSummary: false,
      endDate: { lt: today },
      percentDone: { lt: 100 },
    },
    select: {
      endDate: true,
      equipe: { select: { code: true } },
    },
  });
  const weekSet = new Set<string>();
  const equipeSet = new Set<string>();
  const map = new Map<string, number>(); // "week|equipe" → count
  for (const t of tasks) {
    const w = bucketLabel(startOfWeek(t.endDate), "week");
    const e = t.equipe?.code ?? "—";
    weekSet.add(w);
    equipeSet.add(e);
    const key = `${w}|${e}`;
    map.set(key, (map.get(key) ?? 0) + 1);
  }
  const weeks = [...weekSet].sort();
  const equipes = [...equipeSet].sort();
  const cells: HeatCell[] = [];
  for (const w of weeks) {
    for (const e of equipes) {
      cells.push({ week: w, equipe: e, lateCount: map.get(`${w}|${e}`) ?? 0 });
    }
  }
  return { weeks, equipes, cells };
}
