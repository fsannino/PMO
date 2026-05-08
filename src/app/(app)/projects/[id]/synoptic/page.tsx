import { prisma } from "@/lib/db";
import { notFound } from "next/navigation";
import { computeCurvaS, previstoVsRealizadoPorEquipe, heatmapAtrasos } from "@/lib/curva-s";
import { CurvaSChart } from "@/components/synoptic/CurvaSChart";
import { PrevRealBars } from "@/components/synoptic/PrevRealBars";
import { AtrasoHeatmap } from "@/components/synoptic/AtrasoHeatmap";
import { SynopticToolbar } from "@/components/synoptic/SynopticToolbar";

export const metadata = { title: "Painel Sinóptico" };

export default async function SynopticPage({
  params,
  searchParams,
}: {
  params: Promise<{ id: string }>;
  searchParams: Promise<{ g?: string }>;
}) {
  const { id } = await params;
  const { g } = await searchParams;
  const project = await prisma.project.findUnique({
    where: { id },
    select: { id: true },
  });
  if (!project) notFound();

  const granularity = g === "month" ? "month" : "week";

  const [curva, equipes, heat] = await Promise.all([
    computeCurvaS(project.id, granularity),
    previstoVsRealizadoPorEquipe(project.id),
    heatmapAtrasos(project.id),
  ]);

  const last = curva.series[curva.series.length - 1];
  const sv = last?.sv ?? 0;

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h2 className="text-lg font-semibold">Curva S — Previsto vs Realizado</h2>
          <p className="text-xs text-slate-500">
            Total de trabalho: {curva.totalWork} dia(s) · Última semana:{" "}
            <span className={sv >= 0 ? "text-emerald-700" : "text-rose-700"}>
              SV {sv >= 0 ? "+" : ""}{sv.toFixed(1)}%
            </span>
          </p>
        </div>
        <SynopticToolbar projectId={project.id} />
      </div>

      <section className="rounded-lg border bg-white p-4">
        <CurvaSChart series={curva.series} />
      </section>

      <section className="grid gap-4 lg:grid-cols-2">
        <div className="rounded-lg border bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Previsto vs Realizado por equipe
          </h3>
          <PrevRealBars data={equipes} />
        </div>
        <div className="rounded-lg border bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Heatmap de atrasos (semana × equipe)
          </h3>
          <AtrasoHeatmap weeks={heat.weeks} equipes={heat.equipes} cells={heat.cells} />
        </div>
      </section>

      <p className="text-xs text-slate-500">
        Cálculos em tempo real a partir das tarefas e medições confirmadas.
        Substitui as views <code>DadosFlagCurvaS</code>, <code>DadosPainelSinoptico*</code>
        e as functions <code>RetornaBCWP</code>/<code>RetornaACWP</code> do legado.
      </p>
    </div>
  );
}
