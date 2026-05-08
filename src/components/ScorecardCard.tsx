import type { Scorecard } from "@/lib/scorecard";
import { cn } from "@/lib/utils";

export function ScorecardCard({
  title,
  subtitle,
  data,
  compact,
}: {
  title: string;
  subtitle?: string;
  data: Scorecard;
  compact?: boolean;
}) {
  return (
    <div className="rounded-lg border bg-white p-4 shadow-sm">
      <div className="flex items-baseline justify-between gap-2">
        <div>
          <h3 className="text-sm font-semibold uppercase tracking-wide text-slate-500">{title}</h3>
          {subtitle && <p className="text-xs text-slate-500">{subtitle}</p>}
        </div>
        <span className="text-3xl font-semibold tabular-nums text-brand-700">
          {data.andamentoPct}%
        </span>
      </div>

      <ProgressBar value={data.andamentoPct} className="mt-3" />

      <dl className={cn("mt-4 grid gap-2 text-xs", compact ? "grid-cols-2" : "grid-cols-2 sm:grid-cols-4")}>
        <Stat label="Tarefas" value={data.totalTasks} />
        <Stat label="Concluídas" value={data.done} tone="emerald" />
        <Stat label="Em curso" value={data.inProgress} tone="blue" />
        <Stat label="Não iniciadas" value={data.notStarted} tone="slate" />
        <Stat label="Atrasadas 0–10d" value={data.delayed0to10} tone="amber" />
        <Stat label="Atrasadas +10d" value={data.delayedMore10} tone="rose" />
      </dl>
    </div>
  );
}

function ProgressBar({ value, className }: { value: number; className?: string }) {
  const v = Math.max(0, Math.min(100, value));
  return (
    <div className={cn("h-2 w-full overflow-hidden rounded-full bg-slate-100", className)}>
      <div
        className="h-full bg-brand-500 transition-all"
        style={{ width: `${v}%` }}
      />
    </div>
  );
}

const TONES: Record<string, string> = {
  emerald: "text-emerald-700",
  blue: "text-blue-700",
  amber: "text-amber-700",
  rose: "text-rose-700",
  slate: "text-slate-700",
};

function Stat({ label, value, tone = "slate" }: { label: string; value: number; tone?: string }) {
  return (
    <div>
      <dt className="text-slate-500">{label}</dt>
      <dd className={cn("text-base font-semibold tabular-nums", TONES[tone])}>{value}</dd>
    </div>
  );
}
