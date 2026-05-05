"use client";

import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  Legend,
  ReferenceLine,
} from "recharts";
import type { SerieRow } from "@/lib/curva-s";

export function CurvaSChart({ series }: { series: SerieRow[] }) {
  if (series.length === 0) {
    return <Empty msg="Sem dados para Curva S." />;
  }
  return (
    <ResponsiveContainer width="100%" height={320}>
      <LineChart data={series} margin={{ top: 10, right: 16, left: -8, bottom: 0 }}>
        <CartesianGrid stroke="#f1f5f9" />
        <XAxis dataKey="bucket" tick={{ fontSize: 11, fill: "#475569" }} />
        <YAxis
          tick={{ fontSize: 11, fill: "#475569" }}
          domain={[0, 100]}
          tickFormatter={(v) => `${v}%`}
        />
        <Tooltip
          formatter={(v: number) => `${v.toFixed(1)}%`}
          labelStyle={{ color: "#0f172a", fontWeight: 600 }}
          contentStyle={{ borderRadius: 6, fontSize: 12 }}
        />
        <Legend wrapperStyle={{ fontSize: 12 }} />
        <ReferenceLine y={100} stroke="#94a3b8" strokeDasharray="3 3" />
        <Line
          type="monotone"
          dataKey="planejadoPct"
          name="Planejado (BCWS)"
          stroke="#94a3b8"
          strokeWidth={2}
          dot={false}
        />
        <Line
          type="monotone"
          dataKey="realizadoPct"
          name="Realizado (BCWP)"
          stroke="#2f6feb"
          strokeWidth={2.5}
          dot={false}
        />
      </LineChart>
    </ResponsiveContainer>
  );
}

function Empty({ msg }: { msg: string }) {
  return (
    <div className="grid h-40 place-items-center rounded border bg-slate-50 text-sm text-slate-500">
      {msg}
    </div>
  );
}
