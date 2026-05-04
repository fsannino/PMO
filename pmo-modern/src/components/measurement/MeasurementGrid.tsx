"use client";

import { useState, useTransition } from "react";
import { CheckCircle2, Lock, Save } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Input } from "@/components/ui/Field";
import { upsertMeasurement } from "@/lib/actions/measurement";
import { cn } from "@/lib/utils";

export type MeasurementRow = {
  id: string;            // task id
  wbs: string | null;
  name: string;
  startDate: string;
  endDate: string;
  taskPercent: number;          // valor atual da tarefa
  draftPercent: number | null;  // medição não-confirmada do user
  draftHours: number | null;
  draftComment: string | null;
  locked: boolean;              // já tem MeasurementLock no período
  daysLate: number | null;      // > 0 atrasada
};

export function MeasurementGrid({
  projectId,
  rows,
  windowOpen,
  windowReason,
}: {
  projectId: string;
  rows: MeasurementRow[];
  windowOpen: boolean;
  windowReason?: string;
}) {
  const [filter, setFilter] = useState("");
  const [showLocked, setShowLocked] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [pending, start] = useTransition();
  void projectId;

  // estado local para editar antes de salvar
  const [drafts, setDrafts] = useState<Record<string, { percent: string; hours: string; comment: string }>>(() => {
    const obj: Record<string, { percent: string; hours: string; comment: string }> = {};
    for (const r of rows) {
      obj[r.id] = {
        percent: String(Math.round(r.draftPercent ?? r.taskPercent)),
        hours: r.draftHours == null ? "" : String(r.draftHours),
        comment: r.draftComment ?? "",
      };
    }
    return obj;
  });

  const filtered = rows.filter((r) => {
    if (!showLocked && r.locked) return false;
    if (filter && !r.name.toLowerCase().includes(filter.toLowerCase())) return false;
    return true;
  });

  function update(id: string, patch: Partial<{ percent: string; hours: string; comment: string }>) {
    setDrafts((s) => ({ ...s, [id]: { ...s[id], ...patch } }));
  }

  async function save(id: string, confirm: boolean) {
    setError(null);
    setBusyId(id);
    const d = drafts[id];
    start(async () => {
      const res = await upsertMeasurement({
        taskId: id,
        percentDone: Number(d.percent || 0),
        hoursWorked: d.hours ? Number(d.hours) : null,
        comment: d.comment || null,
        confirm,
      });
      setBusyId(null);
      if (!res.ok) setError(res.error);
    });
  }

  return (
    <div className="space-y-4">
      {!windowOpen && (
        <div className="rounded border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-900">
          ⚠️ <strong>Janela de medição fechada.</strong>{" "}
          {windowReason ?? "Você pode salvar rascunho, mas não confirmar."}
        </div>
      )}

      <div className="flex flex-wrap items-center gap-2">
        <Input
          placeholder="Buscar tarefa…"
          className="max-w-xs"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        />
        <label className="ml-2 flex items-center gap-1 text-sm text-slate-600">
          <input type="checkbox" checked={showLocked} onChange={(e) => setShowLocked(e.target.checked)} />
          incluir confirmadas
        </label>
        <span className="ml-auto text-xs text-slate-500">
          {filtered.length} de {rows.length}
        </span>
      </div>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>
      )}

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">WBS</th>
              <th className="px-3 py-2 text-left">Tarefa</th>
              <th className="px-3 py-2 text-left">Fim</th>
              <th className="px-3 py-2 text-right">Atual</th>
              <th className="px-3 py-2 text-left">% Realizado</th>
              <th className="px-3 py-2 text-left">Horas</th>
              <th className="px-3 py-2 text-left">Comentário</th>
              <th className="px-3 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.map((r) => {
              const d = drafts[r.id];
              const dirty = d && (d.percent !== String(Math.round(r.draftPercent ?? r.taskPercent)));
              return (
                <tr key={r.id} className={cn("hover:bg-slate-50", r.locked && "opacity-60")}>
                  <td className="px-3 py-2 font-mono text-xs">{r.wbs ?? "—"}</td>
                  <td className="px-3 py-2">
                    <div className="flex items-center gap-1">
                      {r.locked && <Lock className="h-3 w-3 text-emerald-600" />}
                      <span>{r.name}</span>
                    </div>
                  </td>
                  <td className="px-3 py-2 whitespace-nowrap">
                    {new Date(r.endDate).toLocaleDateString("pt-BR")}
                    {r.daysLate != null && r.daysLate > 0 && (
                      <span className="ml-1 rounded bg-rose-100 px-1.5 py-0.5 text-[10px] font-medium text-rose-700">
                        +{r.daysLate}d
                      </span>
                    )}
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">{Math.round(r.taskPercent)}%</td>
                  <td className="px-3 py-2">
                    <Input
                      type="number"
                      min={0}
                      max={100}
                      value={d.percent}
                      disabled={r.locked}
                      onChange={(e) => update(r.id, { percent: e.target.value })}
                      className="w-20 text-right tabular-nums"
                    />
                  </td>
                  <td className="px-3 py-2">
                    <Input
                      type="number"
                      min={0}
                      step="0.5"
                      value={d.hours}
                      disabled={r.locked}
                      onChange={(e) => update(r.id, { hours: e.target.value })}
                      className="w-20 text-right tabular-nums"
                      placeholder="—"
                    />
                  </td>
                  <td className="px-3 py-2">
                    <Input
                      value={d.comment}
                      disabled={r.locked}
                      onChange={(e) => update(r.id, { comment: e.target.value })}
                      maxLength={200}
                      className="w-full"
                      placeholder="opcional"
                    />
                  </td>
                  <td className="px-3 py-2 text-right">
                    {r.locked ? (
                      <span className="text-xs text-emerald-700">confirmada</span>
                    ) : (
                      <div className="flex justify-end gap-1">
                        <Button
                          size="sm"
                          variant="secondary"
                          disabled={pending && busyId === r.id}
                          onClick={() => save(r.id, false)}
                          title="Salvar rascunho"
                        >
                          <Save className="h-3.5 w-3.5" />
                        </Button>
                        <Button
                          size="sm"
                          variant="primary"
                          disabled={!windowOpen || (pending && busyId === r.id)}
                          onClick={() => save(r.id, true)}
                          title={windowOpen ? "Confirmar" : "Janela fechada"}
                        >
                          <CheckCircle2 className="h-3.5 w-3.5" />
                        </Button>
                      </div>
                    )}
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={8} className="px-4 py-6 text-center text-slate-500">
                  Nenhuma tarefa medível.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <p className="text-xs text-slate-500">
        Ao <strong>confirmar</strong>, a medição é gravada com lock para o período corrente
        (mês). O percentual é propagado para a tarefa e registrado no histórico.
      </p>
    </div>
  );
}
