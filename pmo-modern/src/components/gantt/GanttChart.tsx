"use client";

import { useMemo, useState } from "react";
import { cn } from "@/lib/utils";

export type GanttTask = {
  id: string;
  name: string;
  wbs: string | null;
  startDate: Date | string;
  endDate: Date | string;
  baselineStart: Date | string | null;
  baselineEnd: Date | string | null;
  percentDone: number;
  status: string;
  isMilestone: boolean;
  isSummary: boolean;
  parentId: string | null;
  predecessorIds: string[];
};

type Zoom = "day" | "week" | "month";

const ZOOM_PX: Record<Zoom, number> = {
  day: 24,    // 24px por dia
  week: 8,    // 8px por dia (≈56px/semana)
  month: 3,   // 3px por dia (≈90px/mês)
};

const ROW_HEIGHT = 28;
const HEADER_HEIGHT = 48;
const LEFT_PANEL_WIDTH = 280;

const DAY_MS = 24 * 60 * 60 * 1000;

const STATUS_FILL: Record<string, string> = {
  NOT_STARTED: "#cbd5e1",
  IN_PROGRESS: "#3b82f6",
  DELAYED: "#f59e0b",
  DONE: "#10b981",
  CANCELLED: "#f43f5e",
};

function dateOnly(d: Date | string): Date {
  const x = new Date(d);
  return new Date(x.getFullYear(), x.getMonth(), x.getDate());
}

function diffDays(a: Date, b: Date): number {
  return Math.round((dateOnly(b).getTime() - dateOnly(a).getTime()) / DAY_MS);
}

function fmtBR(d: Date): string {
  return d.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
}

export function GanttChart({ tasks }: { tasks: GanttTask[] }) {
  const [zoom, setZoom] = useState<Zoom>("week");

  const { rangeStart, rangeEnd, totalDays } = useMemo(() => {
    if (tasks.length === 0) {
      const today = dateOnly(new Date());
      return { rangeStart: today, rangeEnd: today, totalDays: 1 };
    }
    let min = dateOnly(tasks[0].startDate);
    let max = dateOnly(tasks[0].endDate);
    for (const t of tasks) {
      const s = dateOnly(t.startDate);
      const e = dateOnly(t.endDate);
      if (s < min) min = s;
      if (e > max) max = e;
      if (t.baselineStart) {
        const bs = dateOnly(t.baselineStart);
        if (bs < min) min = bs;
      }
      if (t.baselineEnd) {
        const be = dateOnly(t.baselineEnd);
        if (be > max) max = be;
      }
    }
    // padding
    min = new Date(min.getTime() - 3 * DAY_MS);
    max = new Date(max.getTime() + 3 * DAY_MS);
    return { rangeStart: min, rangeEnd: max, totalDays: diffDays(min, max) + 1 };
  }, [tasks]);

  const pxPerDay = ZOOM_PX[zoom];
  const chartWidth = totalDays * pxPerDay;
  const chartHeight = tasks.length * ROW_HEIGHT;

  const taskById = useMemo(() => new Map(tasks.map((t) => [t.id, t])), [tasks]);
  const indexById = useMemo(() => {
    const m = new Map<string, number>();
    tasks.forEach((t, i) => m.set(t.id, i));
    return m;
  }, [tasks]);

  // Geração de marcas no eixo X (uma marca por unidade do zoom)
  const ticks = useMemo(() => {
    const arr: { x: number; label: string; major: boolean }[] = [];
    if (zoom === "day") {
      for (let i = 0; i <= totalDays; i++) {
        const date = new Date(rangeStart.getTime() + i * DAY_MS);
        arr.push({
          x: i * pxPerDay,
          label: fmtBR(date),
          major: date.getDate() === 1,
        });
      }
    } else if (zoom === "week") {
      // segunda-feira como início
      for (let i = 0; i <= totalDays; i++) {
        const date = new Date(rangeStart.getTime() + i * DAY_MS);
        if (date.getDay() === 1 || i === 0) {
          arr.push({
            x: i * pxPerDay,
            label: fmtBR(date),
            major: date.getDate() <= 7,
          });
        }
      }
    } else {
      // month: marca o dia 1 de cada mês
      for (let i = 0; i <= totalDays; i++) {
        const date = new Date(rangeStart.getTime() + i * DAY_MS);
        if (date.getDate() === 1 || i === 0) {
          arr.push({
            x: i * pxPerDay,
            label: date.toLocaleDateString("pt-BR", { month: "short", year: "numeric" }),
            major: date.getMonth() === 0,
          });
        }
      }
    }
    return arr;
  }, [rangeStart, totalDays, pxPerDay, zoom]);

  // marca de hoje
  const todayX = useMemo(() => {
    const today = dateOnly(new Date());
    if (today < rangeStart || today > rangeEnd) return null;
    return diffDays(rangeStart, today) * pxPerDay;
  }, [rangeStart, rangeEnd, pxPerDay]);

  return (
    <div className="rounded-lg border bg-white">
      {/* Toolbar */}
      <div className="flex items-center justify-between border-b px-3 py-2">
        <div className="text-xs text-slate-500">
          {tasks.length} tarefas · {fmtBR(rangeStart)} → {fmtBR(rangeEnd)}
        </div>
        <div className="flex items-center gap-1">
          <span className="text-xs text-slate-500 mr-1">Zoom:</span>
          {(["day", "week", "month"] as Zoom[]).map((z) => (
            <button
              key={z}
              onClick={() => setZoom(z)}
              className={cn(
                "rounded px-2 py-1 text-xs",
                z === zoom ? "bg-brand-600 text-white" : "bg-slate-100 text-slate-700 hover:bg-slate-200",
              )}
            >
              {z === "day" ? "Dia" : z === "week" ? "Semana" : "Mês"}
            </button>
          ))}
        </div>
      </div>

      <div className="overflow-auto" style={{ maxHeight: "70vh" }}>
        <div className="flex" style={{ width: LEFT_PANEL_WIDTH + chartWidth, minWidth: "100%" }}>
          {/* Painel esquerdo */}
          <div className="sticky left-0 z-20 shrink-0 border-r bg-white" style={{ width: LEFT_PANEL_WIDTH }}>
            <div
              className="border-b bg-slate-50 px-3 text-xs font-semibold uppercase tracking-wide text-slate-500"
              style={{ height: HEADER_HEIGHT, lineHeight: `${HEADER_HEIGHT}px` }}
            >
              Tarefa
            </div>
            {tasks.map((t) => (
              <div
                key={t.id}
                className={cn(
                  "flex items-center gap-1 truncate border-b px-3 text-xs",
                  t.isSummary && "bg-slate-50 font-medium",
                )}
                style={{ height: ROW_HEIGHT }}
                title={t.name}
              >
                {t.wbs && <span className="font-mono text-[10px] text-slate-500">{t.wbs}</span>}
                <span className="truncate">{t.name}</span>
              </div>
            ))}
          </div>

          {/* Painel direito (SVG) */}
          <div className="relative">
            <svg
              width={chartWidth}
              height={HEADER_HEIGHT + chartHeight}
              style={{ display: "block" }}
            >
              {/* fundo de header */}
              <rect x={0} y={0} width={chartWidth} height={HEADER_HEIGHT} fill="#f8fafc" />
              {/* grid vertical + labels */}
              {ticks.map((t, i) => (
                <g key={i}>
                  <line
                    x1={t.x}
                    y1={0}
                    x2={t.x}
                    y2={HEADER_HEIGHT + chartHeight}
                    stroke={t.major ? "#cbd5e1" : "#e2e8f0"}
                    strokeWidth={t.major ? 1 : 0.5}
                  />
                  <text
                    x={t.x + 4}
                    y={HEADER_HEIGHT - 8}
                    fontSize={10}
                    fill="#475569"
                  >
                    {t.label}
                  </text>
                </g>
              ))}

              {/* linha do hoje */}
              {todayX !== null && (
                <line
                  x1={todayX}
                  y1={0}
                  x2={todayX}
                  y2={HEADER_HEIGHT + chartHeight}
                  stroke="#dc2626"
                  strokeWidth={1.5}
                  strokeDasharray="3 3"
                />
              )}

              {/* linhas alternadas */}
              {tasks.map((_, i) => (
                <rect
                  key={`row-${i}`}
                  x={0}
                  y={HEADER_HEIGHT + i * ROW_HEIGHT}
                  width={chartWidth}
                  height={ROW_HEIGHT}
                  fill={i % 2 === 0 ? "#fff" : "#fafafa"}
                />
              ))}

              {/* setas de dependência */}
              {tasks.map((t) =>
                t.predecessorIds.map((pid) => {
                  const pred = taskById.get(pid);
                  const predIdx = indexById.get(pid);
                  const succIdx = indexById.get(t.id);
                  if (!pred || predIdx === undefined || succIdx === undefined) return null;
                  const x1 = diffDays(rangeStart, dateOnly(pred.endDate)) * pxPerDay;
                  const y1 = HEADER_HEIGHT + predIdx * ROW_HEIGHT + ROW_HEIGHT / 2;
                  const x2 = diffDays(rangeStart, dateOnly(t.startDate)) * pxPerDay;
                  const y2 = HEADER_HEIGHT + succIdx * ROW_HEIGHT + ROW_HEIGHT / 2;
                  const elbow = Math.max(x1 + 6, x2 - 6);
                  const path = `M${x1},${y1} L${elbow},${y1} L${elbow},${y2} L${x2 - 4},${y2}`;
                  return (
                    <g key={`${pid}-${t.id}`}>
                      <path d={path} stroke="#94a3b8" strokeWidth={1} fill="none" />
                      <polygon
                        points={`${x2 - 4},${y2 - 3} ${x2},${y2} ${x2 - 4},${y2 + 3}`}
                        fill="#94a3b8"
                      />
                    </g>
                  );
                }),
              )}

              {/* barras / marcos / resumos */}
              {tasks.map((t, i) => {
                const xStart = diffDays(rangeStart, dateOnly(t.startDate)) * pxPerDay;
                const xEnd = diffDays(rangeStart, dateOnly(t.endDate)) * pxPerDay;
                const w = Math.max(2, xEnd - xStart);
                const y = HEADER_HEIGHT + i * ROW_HEIGHT + 5;
                const h = ROW_HEIGHT - 10;

                if (t.isMilestone) {
                  const cx = xStart;
                  const cy = HEADER_HEIGHT + i * ROW_HEIGHT + ROW_HEIGHT / 2;
                  const s = 7;
                  return (
                    <polygon
                      key={t.id}
                      points={`${cx},${cy - s} ${cx + s},${cy} ${cx},${cy + s} ${cx - s},${cy}`}
                      fill={t.percentDone >= 100 ? "#10b981" : "#f59e0b"}
                      stroke="#1f2937"
                      strokeWidth={1}
                    >
                      <title>{`${t.name} — ${fmtBR(dateOnly(t.startDate))}`}</title>
                    </polygon>
                  );
                }

                if (t.isSummary) {
                  return (
                    <g key={t.id}>
                      <rect x={xStart} y={y + h / 2 - 3} width={w} height={6} fill="#1f2937" rx={1} />
                      <polygon
                        points={`${xStart},${y + h / 2 + 3} ${xStart + 4},${y + h / 2 + 8} ${xStart + 8},${y + h / 2 + 3}`}
                        fill="#1f2937"
                      />
                      <polygon
                        points={`${xEnd},${y + h / 2 + 3} ${xEnd - 4},${y + h / 2 + 8} ${xEnd - 8},${y + h / 2 + 3}`}
                        fill="#1f2937"
                      />
                      <title>{`${t.name} — ${fmtBR(dateOnly(t.startDate))} → ${fmtBR(dateOnly(t.endDate))}`}</title>
                    </g>
                  );
                }

                // baseline (atrás)
                const blStart = t.baselineStart ? diffDays(rangeStart, dateOnly(t.baselineStart)) * pxPerDay : null;
                const blEnd = t.baselineEnd ? diffDays(rangeStart, dateOnly(t.baselineEnd)) * pxPerDay : null;

                const fill = STATUS_FILL[t.status] ?? "#3b82f6";
                const progressW = (w * Math.max(0, Math.min(100, t.percentDone))) / 100;

                return (
                  <g key={t.id}>
                    {blStart !== null && blEnd !== null && (
                      <rect
                        x={blStart}
                        y={y + h - 3}
                        width={Math.max(2, blEnd - blStart)}
                        height={3}
                        fill="#94a3b8"
                        opacity={0.6}
                      />
                    )}
                    <rect x={xStart} y={y} width={w} height={h} fill={fill} opacity={0.25} rx={2} />
                    <rect x={xStart} y={y} width={progressW} height={h} fill={fill} rx={2} />
                    <title>
                      {`${t.name}\n${fmtBR(dateOnly(t.startDate))} → ${fmtBR(dateOnly(t.endDate))}\n${Math.round(t.percentDone)}% · ${t.status}`}
                    </title>
                  </g>
                );
              })}
            </svg>
          </div>
        </div>
      </div>

      {/* Legenda */}
      <div className="flex flex-wrap items-center gap-3 border-t bg-slate-50 px-3 py-2 text-xs text-slate-600">
        <Legend color="#3b82f6" label="Em andamento" />
        <Legend color="#10b981" label="Concluído" />
        <Legend color="#f59e0b" label="Atrasado" />
        <Legend color="#cbd5e1" label="Não iniciado" />
        <span className="ml-2 inline-flex items-center gap-1">
          <span className="inline-block h-1 w-4 bg-slate-400 opacity-60" /> Baseline
        </span>
        <span className="ml-2 inline-flex items-center gap-1">
          <svg width={10} height={10}><polygon points="5,0 10,5 5,10 0,5" fill="#f59e0b" stroke="#1f2937" strokeWidth={1} /></svg>
          Marco
        </span>
        <span className="ml-auto text-xs text-rose-600">— linha vermelha tracejada = hoje</span>
      </div>
    </div>
  );
}

function Legend({ color, label }: { color: string; label: string }) {
  return (
    <span className="inline-flex items-center gap-1">
      <span className="inline-block h-3 w-3 rounded-sm" style={{ backgroundColor: color }} />
      {label}
    </span>
  );
}
