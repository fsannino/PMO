"use client";

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  Legend,
} from "recharts";
import type { EquipeRow } from "@/lib/curva-s";

export function PrevRealBars({ data }: { data: EquipeRow[] }) {
  if (data.length === 0) {
    return (
      <div className="grid h-40 place-items-center rounded border bg-slate-50 text-sm text-slate-500">
        Sem dados.
      </div>
    );
  }
  return (
    <ResponsiveContainer width="100%" height={320}>
      <BarChart data={data} margin={{ top: 10, right: 16, left: -8, bottom: 0 }}>
        <CartesianGrid stroke="#f1f5f9" />
        <XAxis dataKey="equipe" tick={{ fontSize: 11, fill: "#475569" }} />
        <YAxis
          tick={{ fontSize: 11, fill: "#475569" }}
          domain={[0, 100]}
          tickFormatter={(v) => `${v}%`}
        />
        <Tooltip
          formatter={(v: number) => `${v.toFixed(0)}%`}
          contentStyle={{ borderRadius: 6, fontSize: 12 }}
        />
        <Legend wrapperStyle={{ fontSize: 12 }} />
        <Bar dataKey="planejadoPct" name="Planejado" fill="#cbd5e1" />
        <Bar dataKey="realizadoPct" name="Realizado" fill="#2f6feb" />
      </BarChart>
    </ResponsiveContainer>
  );
}
