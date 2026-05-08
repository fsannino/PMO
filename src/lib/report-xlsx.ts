// Export XLSX para os 6 relatórios — usa SheetJS (xlsx).

import * as XLSX from "xlsx";
import type {
  ReportDetalhadoRow,
  ReportConsolidadoRow,
  ReportCriticidadeRow,
  ReportCommentRow,
  ReportIssuesKPIRow,
  ReportAtrasadasRow,
  ReportHeader,
  ReportId,
} from "./reports";

function buildSheet(rows: Record<string, unknown>[], headers: string[]): XLSX.WorkSheet {
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  // largura aproximada por coluna
  ws["!cols"] = headers.map((h) => ({ wch: Math.max(10, Math.min(40, h.length + 4)) }));
  return ws;
}

function metadataSheet(title: string, header: ReportHeader): XLSX.WorkSheet {
  return XLSX.utils.aoa_to_sheet([
    ["Relatório:", title],
    ["Projeto:", `${header.projectCode} — ${header.projectName}`],
    ["Gerado em:", header.generatedAt],
    ["Gerado por:", header.generatedBy],
    [],
    ["KPI", "Valor"],
    ["Andamento", `${header.scorecard.andamentoPct}%`],
    ["Total de tarefas", header.scorecard.totalTasks],
    ["Atrasadas", header.scorecard.late],
  ]);
}

export function buildReportXlsx(
  reportId: ReportId,
  header: ReportHeader,
  data: unknown,
): Buffer {
  const wb = XLSX.utils.book_new();

  switch (reportId) {
    case "detalhado": {
      const rows = data as ReportDetalhadoRow[];
      const mapped = rows.map((r) => ({
        WBS: r.wbs, Nome: r.name, Início: r.start, Fim: r.end,
        "Duração (dias)": r.duration, "% Concluído": r.percent, Status: r.status,
        Responsável: r.assignee, Equipe: r.equipe, Área: r.area, "Dias atraso": r.daysLate,
      }));
      const headers = Object.keys(mapped[0] ?? {});
      XLSX.utils.book_append_sheet(wb, metadataSheet("Detalhado", header), "Resumo");
      XLSX.utils.book_append_sheet(wb, buildSheet(mapped, headers), "Detalhado");
      break;
    }
    case "consolidado": {
      const rows = data as ReportConsolidadoRow[];
      const mapped = rows.map((r) => ({
        Equipe: r.equipe, "Total tarefas": r.total, Concluídas: r.done,
        "Em andamento": r.inProgress, "Não iniciadas": r.notStarted,
        Atrasadas: r.late, "Trabalho (dias)": r.workDays, "% Médio": r.avgPercent,
      }));
      const headers = Object.keys(mapped[0] ?? {});
      XLSX.utils.book_append_sheet(wb, metadataSheet("Consolidado", header), "Resumo");
      XLSX.utils.book_append_sheet(wb, buildSheet(mapped, headers), "Consolidado");
      break;
    }
    case "criticidade": {
      const rows = data as ReportCriticidadeRow[];
      const mapped = rows.map((r) => ({
        Tipo: r.type, Título: r.title, Nível: r.level, Status: r.status, Owner: r.owner,
      }));
      const headers = Object.keys(mapped[0] ?? {});
      XLSX.utils.book_append_sheet(wb, metadataSheet("Criticidade", header), "Resumo");
      XLSX.utils.book_append_sheet(wb, buildSheet(mapped, headers), "Criticidade");
      break;
    }
    case "comentarios": {
      const rows = data as ReportCommentRow[];
      const mapped = rows.map((r) => ({
        Data: r.date, Autor: r.author, Tipo: r.type, Item: r.target, Comentário: r.body,
      }));
      const headers = Object.keys(mapped[0] ?? {});
      XLSX.utils.book_append_sheet(wb, metadataSheet("Comentários", header), "Resumo");
      XLSX.utils.book_append_sheet(wb, buildSheet(mapped, headers), "Comentários");
      break;
    }
    case "issues-kpi": {
      const rows = data as ReportIssuesKPIRow[];
      const mapped = rows.map((r) => ({
        Mês: r.month, Abertas: r.opened, Fechadas: r.closed, Saldo: r.open,
      }));
      const headers = Object.keys(mapped[0] ?? {});
      XLSX.utils.book_append_sheet(wb, metadataSheet("Issues KPI", header), "Resumo");
      XLSX.utils.book_append_sheet(wb, buildSheet(mapped, headers), "KPI");
      break;
    }
    case "atrasadas": {
      const rows = data as ReportAtrasadasRow[];
      const mapped = rows.map((r) => ({
        WBS: r.wbs, Nome: r.name, Fim: r.end, "Dias atraso": r.daysLate,
        "% Concluído": r.percent, Responsável: r.assignee, Equipe: r.equipe,
      }));
      const headers = Object.keys(mapped[0] ?? {});
      XLSX.utils.book_append_sheet(wb, metadataSheet("Atrasadas", header), "Resumo");
      XLSX.utils.book_append_sheet(wb, buildSheet(mapped, headers), "Atrasadas");
      break;
    }
  }
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer;
}
