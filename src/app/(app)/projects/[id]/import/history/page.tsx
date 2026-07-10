import { prisma } from "@/lib/db";
import { notFound } from "next/navigation";
import Link from "next/link";
import { ChevronLeft } from "lucide-react";

export const metadata = { title: "Histórico de importações" };

const SOURCE_LABEL: Record<string, string> = {
  EXCEL: "Excel",
  MSPROJECT_XML: "MS Project (XML)",
  MANUAL: "Manual",
};

const MODE_LABEL: Record<string, string> = {
  REPLACE: "Substituir",
  MERGE: "Mesclar",
  ONLY_NEW: "Apenas novos",
};

export default async function ImportHistoryPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const project = await prisma.project.findUnique({ where: { id }, select: { id: true } });
  if (!project) notFound();

  const logs = await prisma.importLog.findMany({
    where: { projectId: project.id },
    include: { user: { select: { name: true, email: true } } },
    orderBy: { createdAt: "desc" },
    take: 100,
  });

  return (
    <div className="space-y-4">
      <Link
        href={`/projects/${project.id}/import`}
        className="inline-flex items-center gap-1 text-sm text-brand-700 hover:underline"
      >
        <ChevronLeft className="h-3.5 w-3.5" /> voltar para Importar
      </Link>

      <div className="overflow-hidden rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Data</th>
              <th className="px-3 py-2 text-left">Usuário</th>
              <th className="px-3 py-2 text-left">Origem</th>
              <th className="px-3 py-2 text-left">Template</th>
              <th className="px-3 py-2 text-left">Modo</th>
              <th className="px-3 py-2 text-left">Arquivo</th>
              <th className="px-3 py-2 text-right">Registros</th>
              <th className="px-3 py-2 text-left">Avisos</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {logs.length === 0 && (
              <tr>
                <td colSpan={8} className="px-4 py-6 text-center text-slate-500">
                  Sem histórico de importações.
                </td>
              </tr>
            )}
            {logs.map((l) => (
              <tr key={l.id} className="hover:bg-slate-50">
                <td className="px-3 py-2 whitespace-nowrap">{l.createdAt.toLocaleString("pt-BR")}</td>
                <td className="px-3 py-2 text-xs text-slate-600">{l.user?.name ?? "—"}</td>
                <td className="px-3 py-2">
                  <span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{SOURCE_LABEL[l.source] ?? l.source}</span>
                </td>
                <td className="px-3 py-2">
                  <span className="rounded bg-brand-50 px-2 py-0.5 text-xs text-brand-700">{l.template ?? "—"}</span>
                </td>
                <td className="px-3 py-2 text-xs">{MODE_LABEL[l.mode] ?? l.mode}</td>
                <td className="px-3 py-2 font-mono text-xs">{l.filename}</td>
                <td className="px-3 py-2 text-right tabular-nums">{l.recordCount}</td>
                <td className="px-3 py-2">
                  {l.warnings ? (
                    <details className="cursor-pointer text-xs text-amber-700">
                      <summary>{l.warnings.split("\n").length} aviso(s)</summary>
                      <pre className="mt-1 max-h-40 overflow-auto whitespace-pre-wrap rounded bg-amber-50 p-2 text-[11px]">
                        {l.warnings}
                      </pre>
                    </details>
                  ) : (
                    <span className="text-xs text-slate-400">—</span>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
