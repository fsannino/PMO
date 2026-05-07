// Configuração e helpers por módulo (CUT/GVI/GRF/TCP/TIN/PMO/LIGHT).
// Centraliza diferenças de UX/UI conforme Project.module.

import type { Module } from "./enums";

export type ModuleConfig = {
  code: Module;
  label: string;
  description: string;
  color: string;     // tone Tailwind: bg-* / text-*
  emoji: string;
  hints: string[];   // dicas operacionais mostradas no banner do projeto
  tabsHidden?: string[];  // labels de abas a esconder (ex.: ["Importar"] em LIGHT)
  measurementUnit?: "percent" | "cases";  // tipo de medição preferido
};

export const MODULE_CONFIG: Record<Module, ModuleConfig> = {
  PMO: {
    code: "PMO",
    label: "PMO Master",
    description: "Visão completa do escritório de projetos.",
    color: "bg-brand-50 text-brand-700",
    emoji: "🏢",
    hints: [],
    measurementUnit: "percent",
  },
  CUT: {
    code: "CUT",
    label: "CutOver",
    description: "Operações de transição e migração.",
    color: "bg-violet-100 text-violet-800",
    emoji: "🔁",
    hints: [
      "Janelas de operação são curtas — atenção redobrada às janelas HB.",
      "Atualizações de percentual em CutOver impactam o painel de transição em tempo real.",
    ],
    measurementUnit: "percent",
  },
  GVI: {
    code: "GVI",
    label: "Governança",
    description: "Saneamento e governança integrada.",
    color: "bg-emerald-100 text-emerald-800",
    emoji: "🛡️",
    hints: [
      "Use Issues para registrar pendências de saneamento.",
      "Estimativas devem refletir consenso da governança antes do registro.",
    ],
    measurementUnit: "percent",
  },
  GRF: {
    code: "GRF",
    label: "Gráfico",
    description: "Acompanhamento gráfico simplificado.",
    color: "bg-sky-100 text-sky-800",
    emoji: "📈",
    hints: [
      "Medição simplificada: apenas % sem horas detalhadas.",
    ],
    measurementUnit: "percent",
  },
  TCP: {
    code: "TCP",
    label: "Teste de Carga",
    description: "Validação de capacidade por casos de teste.",
    color: "bg-amber-100 text-amber-800",
    emoji: "⚡",
    hints: [
      "Cada tarefa representa um conjunto de casos de teste.",
      "Use o campo de horas como número de casos executados (informativo).",
    ],
    measurementUnit: "cases",
  },
  TIN: {
    code: "TIN",
    label: "Teste Integrado",
    description: "Validação de integração entre módulos.",
    color: "bg-orange-100 text-orange-800",
    emoji: "🔗",
    hints: [
      "Acompanhe percentual por dimensão (% de cases × % de cobertura).",
      "Use a aba Traceability (futura) para mapear requisitos × testes × tarefas.",
    ],
    measurementUnit: "cases",
  },
  LIGHT: {
    code: "LIGHT",
    label: "Light",
    description: "Versão simplificada (sem CRs e Traceability).",
    color: "bg-slate-100 text-slate-700",
    emoji: "🪶",
    hints: [
      "Módulo Light: foco em medição e relatórios; sem Change Requests e Traceability.",
    ],
    tabsHidden: ["CRs"],
    measurementUnit: "percent",
  },
};

export function getModuleConfig(module: string): ModuleConfig {
  if (module in MODULE_CONFIG) return MODULE_CONFIG[module as Module];
  return MODULE_CONFIG.PMO;
}
