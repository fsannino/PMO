"use client";

import { cn } from "@/lib/utils";
import type { HeatCell } from "@/lib/curva-s";

export function AtrasoHeatmap({
  weeks,
  equipes,
  cells,
}: {
  weeks: string[];
  equipes: string[];
  cells: HeatCell[];
}) {
  if (weeks.length === 0 || equipes.length === 0) {
    return (
      <div className="grid h-40 place-items-center rounded border bg-slate-50 text-sm text-slate-500">
        Nenhuma tarefa atrasada — 🎉
      </div>
    );
  }
  const max = Math.max(1, ...cells.map((c) => c.lateCount));
  const get = (w: string, e: string) =>
    cells.find((c) => c.week === w && c.equipe === e)?.lateCount ?? 0;

  return (
    <div className="overflow-x-auto">
      <table className="border-separate border-spacing-1 text-xs">
        <thead>
          <tr>
            <th className="px-2 py-1 text-left font-medium text-slate-500">Semana \ Equipe</th>
            {equipes.map((e) => (
              <th key={e} className="px-2 py-1 text-center font-mono text-slate-600">{e}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {weeks.map((w) => (
            <tr key={w}>
              <td className="px-2 py-1 text-left font-mono text-slate-600">{w}</td>
              {equipes.map((e) => {
                const v = get(w, e);
                const intensity = v / max; // 0..1
                const bg =
                  v === 0 ? "bg-slate-50" :
                  intensity < 0.34 ? "bg-amber-100" :
                  intensity < 0.67 ? "bg-orange-300" : "bg-rose-500";
                const text = v === 0 ? "text-slate-300" : intensity > 0.5 ? "text-white" : "text-slate-900";
                return (
                  <td
                    key={e}
                    className={cn("h-8 w-12 rounded text-center font-medium", bg, text)}
                    title={`${w} · ${e}: ${v} tarefa(s) atrasada(s)`}
                  >
                    {v || ""}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
