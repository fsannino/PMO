// Scorecard — substitui as functions ScoreCardAndamento*, ScoreCardAtrasadas*
// do legado. Calcula KPIs de andamento e atraso em runtime (não persiste).

import { prisma } from "./db";

export type Scorecard = {
  totalTasks: number;
  totalWork: number;       // soma de duração (dias) das atividades
  weightedDone: number;    // soma de (durationDays * percentDone/100)
  andamentoPct: number;    // weightedDone / totalWork
  done: number;            // contagem de DONE
  inProgress: number;
  notStarted: number;
  delayed0to10: number;    // hoje > endDate, percent < 100, 0..10 dias atrasada
  delayedMore10: number;
};

function daysBetween(a: Date, b: Date): number {
  return Math.floor((b.getTime() - a.getTime()) / 86400000);
}

export async function projectScorecard(projectId: string, today: Date = new Date()): Promise<Scorecard> {
  const tasks = await prisma.task.findMany({
    where: { projectId, deletedAt: null, isSummary: false },
    select: {
      id: true,
      durationDays: true,
      percentDone: true,
      status: true,
      endDate: true,
    },
  });

  let totalWork = 0;
  let weightedDone = 0;
  let done = 0;
  let inProgress = 0;
  let notStarted = 0;
  let delayed0to10 = 0;
  let delayedMore10 = 0;

  for (const t of tasks) {
    const w = t.durationDays ?? 1;
    totalWork += w;
    weightedDone += (w * Math.max(0, Math.min(100, t.percentDone))) / 100;
    if (t.percentDone >= 100) done++;
    else if (t.percentDone > 0) inProgress++;
    else notStarted++;

    if (t.percentDone < 100) {
      const lateDays = daysBetween(t.endDate, today);
      if (lateDays > 0) {
        if (lateDays <= 10) delayed0to10++;
        else delayedMore10++;
      }
    }
  }

  const andamentoPct = totalWork > 0 ? Math.round((weightedDone / totalWork) * 100) : 0;
  return {
    totalTasks: tasks.length,
    totalWork,
    weightedDone,
    andamentoPct,
    done,
    inProgress,
    notStarted,
    delayed0to10,
    delayedMore10,
  };
}
