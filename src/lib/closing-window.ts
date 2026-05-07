// Janela HB (Hora de Bloqueio) — substitui a função VerificaHB() do legado
// e a tabela CONFIG_FECHAMENTO. Verifica se medições podem ser registradas
// no momento atual.
//
// Regra: para um projeto X, primeiro tenta config específica do projeto
// (projectId=X, dayOfWeek=N); se não existir, cai para a config global
// (projectId=null, dayOfWeek=N).

import { prisma } from "./db";

export type WindowStatus = {
  open: boolean;
  reason?: string;
  config?: { startTime: string; endTime: string; enabled: boolean; scope: "project" | "global" };
  now: Date;
};

const DAY_NAMES = ["domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado"];

function timeToMinutes(hhmm: string): number {
  const [h, m] = hhmm.split(":").map(Number);
  return h * 60 + (m ?? 0);
}

export async function getMeasurementWindow(
  projectId: string,
  when: Date = new Date(),
): Promise<WindowStatus> {
  const dayOfWeek = when.getDay();

  // tenta config específica do projeto
  const projConfig = await prisma.configFechamento.findFirst({
    where: { projectId, dayOfWeek },
  });
  // fallback global
  const config = projConfig
    ? { ...projConfig, scope: "project" as const }
    : await prisma.configFechamento
        .findFirst({ where: { projectId: null, dayOfWeek } })
        .then((c) => (c ? { ...c, scope: "global" as const } : null));

  if (!config) {
    return {
      open: false,
      now: when,
      reason: `Sem configuração de janela para ${DAY_NAMES[dayOfWeek]}.`,
    };
  }
  if (!config.enabled) {
    return {
      open: false,
      now: when,
      reason: `Janela fechada em ${DAY_NAMES[dayOfWeek]}.`,
      config: { startTime: config.startTime, endTime: config.endTime, enabled: false, scope: config.scope },
    };
  }

  const nowMin = when.getHours() * 60 + when.getMinutes();
  const startMin = timeToMinutes(config.startTime);
  const endMin = timeToMinutes(config.endTime);

  if (nowMin < startMin || nowMin > endMin) {
    return {
      open: false,
      now: when,
      reason: `Fora da janela ${config.startTime}–${config.endTime} de ${DAY_NAMES[dayOfWeek]}.`,
      config: { startTime: config.startTime, endTime: config.endTime, enabled: true, scope: config.scope },
    };
  }

  return {
    open: true,
    now: when,
    config: { startTime: config.startTime, endTime: config.endTime, enabled: true, scope: config.scope },
  };
}

export function periodOf(date: Date = new Date()): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}
