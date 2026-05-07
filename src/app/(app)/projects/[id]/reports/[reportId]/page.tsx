import { prisma } from "@/lib/db";
import { notFound } from "next/navigation";
import Link from "next/link";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { ChevronLeft, FileSpreadsheet, FileText } from "lucide-react";
import {
  REPORTS,
  reportHeader,
  reportDetalhado,
  reportConsolidado,
  reportCriticidade,
  reportComentarios,
  reportIssuesKPI,
  reportAtrasadas,
  type ReportId,
} from "@/lib/reports";
import { StatusBadge } from "@/components/work-items/StatusBadge";

export const metadata = { title: "Preview de relatório — CollabZ" };

export default async function ReportPreviewPage({
  params,
}: {
  params: Promise<{ id: string; reportId: string }>;
}) {
  const { id, reportId: reportIdRaw } = await params;
  if (!(reportIdRaw in REPORTS)) notFound();
  const reportId = reportIdRaw as ReportId;
  const cfg = REPORTS[reportId];

  const project = await prisma.project.findUnique({ where: { id }, select: { id: true } });
  if (!project) notFound();

  const session = await getServerSession(authOptions);
  const header = await reportHeader(project.id, session?.user.name ?? "—");
  if (!header) notFound();

  let body: React.ReactNode;
  switch (reportId) {
    case "detalhado": {
      const rows = await reportDetalhado(project.id);
      body = (
        <Table headers={["WBS", "Tarefa", "Início", "Fim", "Dur.", "%", "Status", "Resp.", "Equipe", "Atraso"]}>
          {rows.map((r, i) => (
            <tr key={i} className="hover:bg-slate-50">
              <td className="px-2 py-1 font-mono text-xs">{r.wbs}</td>
              <td className="px-2 py-1">{r.name}</td>
              <td className="px-2 py-1 whitespace-nowrap">{r.start}</td>
              <td className="px-2 py-1 whitespace-nowrap">{r.end}</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.duration}d</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.percent}%</td>
              <td className="px-2 py-1"><StatusBadge value={r.status} size="xs" /></td>
              <td className="px-2 py-1 text-xs text-slate-600">{r.assignee}</td>
              <td className="px-2 py-1 text-xs text-slate-600">{r.equipe}</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.daysLate ? <span className="text-rose-700">{r.daysLate}d</span> : "—"}</td>
            </tr>
          ))}
        </Table>
      );
      break;
    }
    case "consolidado": {
      const rows = await reportConsolidado(project.id);
      body = (
        <Table headers={["Equipe", "Total", "Concl.", "Em curso", "Não iniciadas", "Atrasadas", "Trabalho (d)", "% Médio"]}>
          {rows.map((r, i) => (
            <tr key={i} className="hover:bg-slate-50">
              <td className="px-2 py-1 font-medium">{r.equipe}</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.total}</td>
              <td className="px-2 py-1 text-right tabular-nums text-emerald-700">{r.done}</td>
              <td className="px-2 py-1 text-right tabular-nums text-blue-700">{r.inProgress}</td>
              <td className="px-2 py-1 text-right tabular-nums text-slate-500">{r.notStarted}</td>
              <td className="px-2 py-1 text-right tabular-nums text-rose-700">{r.late}</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.workDays}</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.avgPercent}%</td>
            </tr>
          ))}
        </Table>
      );
      break;
    }
    case "criticidade": {
      const rows = await reportCriticidade(project.id);
      body = (
        <Table headers={["Tipo", "Título", "Nível", "Status", "Owner"]}>
          {rows.map((r, i) => (
            <tr key={i} className="hover:bg-slate-50">
              <td className="px-2 py-1"><span className={`rounded px-1.5 py-0.5 text-xs ${r.type === "ISSUE" ? "bg-amber-100 text-amber-800" : "bg-rose-100 text-rose-800"}`}>{r.type}</span></td>
              <td className="px-2 py-1">{r.title}</td>
              <td className="px-2 py-1"><StatusBadge value={r.level} size="xs" /></td>
              <td className="px-2 py-1"><StatusBadge value={r.status} size="xs" /></td>
              <td className="px-2 py-1 text-xs text-slate-600">{r.owner}</td>
            </tr>
          ))}
        </Table>
      );
      break;
    }
    case "comentarios": {
      const rows = await reportComentarios(project.id);
      body = (
        <Table headers={["Data", "Autor", "Tipo", "Item", "Comentário"]}>
          {rows.map((r, i) => (
            <tr key={i} className="hover:bg-slate-50">
              <td className="px-2 py-1 whitespace-nowrap text-xs text-slate-500">{r.date}</td>
              <td className="px-2 py-1 text-xs">{r.author}</td>
              <td className="px-2 py-1 text-xs"><span className="rounded bg-slate-100 px-1.5 py-0.5">{r.type}</span></td>
              <td className="px-2 py-1 text-xs">{r.target}</td>
              <td className="px-2 py-1 whitespace-pre-wrap">{r.body}</td>
            </tr>
          ))}
        </Table>
      );
      break;
    }
    case "issues-kpi": {
      const rows = await reportIssuesKPI(project.id);
      body = (
        <Table headers={["Mês", "Abertas", "Fechadas", "Saldo aberto"]}>
          {rows.map((r, i) => (
            <tr key={i} className="hover:bg-slate-50">
              <td className="px-2 py-1 font-mono">{r.month}</td>
              <td className="px-2 py-1 text-right tabular-nums text-amber-700">{r.opened}</td>
              <td className="px-2 py-1 text-right tabular-nums text-emerald-700">{r.closed}</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.open}</td>
            </tr>
          ))}
        </Table>
      );
      break;
    }
    case "atrasadas": {
      const rows = await reportAtrasadas(project.id);
      body = (
        <Table headers={["WBS", "Nome", "Fim", "Atraso", "%", "Resp.", "Equipe"]}>
          {rows.map((r, i) => (
            <tr key={i} className="hover:bg-slate-50">
              <td className="px-2 py-1 font-mono text-xs">{r.wbs}</td>
              <td className="px-2 py-1">{r.name}</td>
              <td className="px-2 py-1 whitespace-nowrap">{r.end}</td>
              <td className="px-2 py-1 text-right tabular-nums text-rose-700">{r.daysLate}d</td>
              <td className="px-2 py-1 text-right tabular-nums">{r.percent}%</td>
              <td className="px-2 py-1 text-xs text-slate-600">{r.assignee}</td>
              <td className="px-2 py-1 text-xs text-slate-600">{r.equipe}</td>
            </tr>
          ))}
        </Table>
      );
      break;
    }
  }

  return (
    <div className="space-y-4">
      <Link href={`/projects/${project.id}/reports`} className="inline-flex items-center gap-1 text-sm text-brand-700 hover:underline">
        <ChevronLeft className="h-3.5 w-3.5" /> voltar para galeria
      </Link>

      <div className="flex flex-wrap items-end justify-between gap-3 border-b pb-3">
        <div>
          <h2 className="text-xl font-semibold">{cfg.label}</h2>
          <p className="text-xs text-slate-500">{cfg.desc}</p>
          <p className="mt-1 text-xs text-slate-400">
            Gerado em {header.generatedAt} · Andamento {header.scorecard.andamentoPct}% ·{" "}
            {header.scorecard.totalTasks} tarefas · {header.scorecard.late} atrasadas
          </p>
        </div>
        <div className="flex gap-2">
          <a
            href={`/api/reports/${project.id}/${reportId}?format=xlsx`}
            download
            className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium hover:bg-slate-50"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" /> Baixar XLSX
          </a>
          <a
            href={`/api/reports/${project.id}/${reportId}?format=pdf`}
            download
            className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium hover:bg-slate-50"
          >
            <FileText className="h-3.5 w-3.5 text-rose-600" /> Baixar PDF
          </a>
        </div>
      </div>

      <div className="overflow-x-auto rounded-lg border bg-white">{body}</div>
    </div>
  );
}

function Table({ headers, children }: { headers: string[]; children: React.ReactNode }) {
  return (
    <table className="w-full text-sm">
      <thead className="bg-slate-50 text-xs uppercase text-slate-500">
        <tr>
          {headers.map((h) => (
            <th key={h} className="px-2 py-2 text-left">{h}</th>
          ))}
        </tr>
      </thead>
      <tbody className="divide-y divide-slate-100">{children}</tbody>
    </table>
  );
}
