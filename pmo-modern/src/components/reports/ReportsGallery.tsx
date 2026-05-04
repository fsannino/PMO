"use client";

import { Download, FileSpreadsheet, FileText } from "lucide-react";
import { REPORTS, type ReportId } from "@/lib/reports";

const ICONS: Record<ReportId, string> = {
  detalhado: "📋",
  consolidado: "📊",
  criticidade: "⚠️",
  comentarios: "💬",
  "issues-kpi": "📈",
  atrasadas: "⏰",
};

export function ReportsGallery({ projectId }: { projectId: string }) {
  const ids = Object.keys(REPORTS) as ReportId[];
  return (
    <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
      {ids.map((id) => {
        const r = REPORTS[id];
        return (
          <div key={id} className="rounded-lg border bg-white p-4 shadow-sm transition hover:shadow">
            <div className="mb-2 flex items-start gap-2">
              <span className="text-2xl">{ICONS[id]}</span>
              <div>
                <h3 className="font-semibold">{r.label}</h3>
                <p className="mt-1 text-xs text-slate-500">{r.desc}</p>
              </div>
            </div>
            <div className="mt-3 flex gap-2 border-t pt-3">
              <a
                href={`/api/reports/${projectId}/${id}?format=xlsx`}
                download
                className="inline-flex flex-1 items-center justify-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium hover:bg-slate-50"
              >
                <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
                XLSX
              </a>
              <a
                href={`/api/reports/${projectId}/${id}?format=pdf`}
                download
                className="inline-flex flex-1 items-center justify-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium hover:bg-slate-50"
              >
                <FileText className="h-3.5 w-3.5 text-rose-600" />
                PDF
              </a>
              <a
                href={`/projects/${projectId}/reports/${id}`}
                className="inline-flex items-center justify-center gap-1.5 rounded-md bg-brand-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-brand-700"
              >
                <Download className="h-3.5 w-3.5" />
                Preview
              </a>
            </div>
          </div>
        );
      })}
    </div>
  );
}
