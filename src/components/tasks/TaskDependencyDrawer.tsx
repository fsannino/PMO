"use client";

import { useState, useTransition } from "react";
import { X, Plus, Trash2 } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Select, Input } from "@/components/ui/Field";
import { createDependency, deleteDependency } from "@/lib/actions/tasks";
import { DependencyTypes } from "@/lib/enums";
import type { TaskRow } from "./TaskTreeView";

export function TaskDependencyDrawer({
  task,
  allTasks,
  canWrite,
  onClose,
}: {
  task: TaskRow;
  allTasks: TaskRow[];
  canWrite: boolean;
  onClose: () => void;
}) {
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();
  const [predId, setPredId] = useState("");
  const [type, setType] = useState("FS");
  const [lag, setLag] = useState(0);

  function add() {
    if (!predId) return;
    setError(null);
    start(async () => {
      const res = await createDependency({ predecessorId: predId, successorId: task.id, type, lagDays: lag });
      if (res.ok) {
        setPredId("");
        setLag(0);
        onClose(); // forçar refresh da lista
      } else {
        setError(res.error);
      }
    });
  }

  function remove(depId: string) {
    setError(null);
    start(async () => {
      const res = await deleteDependency(depId);
      if (res.ok) onClose();
      else setError(res.error);
    });
  }

  // tarefas já em uso como predecessor
  const usedIds = new Set(task.predecessors.map((p) => p.predecessorId));
  const available = allTasks.filter((t) => !usedIds.has(t.id));

  return (
    <div className="fixed inset-0 z-40 flex justify-end bg-black/30">
      <div className="flex h-full w-full max-w-md flex-col bg-white shadow-xl">
        <div className="flex items-center justify-between border-b px-4 py-3">
          <h3 className="font-semibold">Predecessores — {task.name}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700">
            <X className="h-4 w-4" />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-4">
          {error && <p className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</p>}

          <ul className="space-y-2">
            {task.predecessors.length === 0 && (
              <li className="rounded border border-dashed border-slate-300 p-3 text-center text-sm text-slate-500">
                Sem predecessores.
              </li>
            )}
            {task.predecessors.map((p) => (
              <li key={p.id} className="flex items-center justify-between rounded border bg-slate-50 px-3 py-2 text-sm">
                <span>
                  {p.predecessor.wbs && <span className="mr-1 font-mono text-xs text-slate-500">{p.predecessor.wbs}</span>}
                  {p.predecessor.name}
                </span>
                {canWrite && (
                  <button
                    onClick={() => remove(p.id)}
                    className="text-slate-400 hover:text-rose-600"
                    disabled={pending}
                    title="Remover"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                )}
              </li>
            ))}
          </ul>

          {canWrite && (
            <div className="rounded border bg-white p-3 space-y-2">
              <h4 className="text-sm font-medium">Adicionar predecessor</h4>
              <Field label="Tarefa">
                <Select value={predId} onChange={(e) => setPredId(e.target.value)}>
                  <option value="">— selecione —</option>
                  {available.map((t) => (
                    <option key={t.id} value={t.id}>
                      {t.wbs ? `${t.wbs} ` : ""}{t.name}
                    </option>
                  ))}
                </Select>
              </Field>
              <div className="grid grid-cols-2 gap-2">
                <Field label="Tipo">
                  <Select value={type} onChange={(e) => setType(e.target.value)}>
                    {DependencyTypes.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </Select>
                </Field>
                <Field label="Lag (dias)">
                  <Input type="number" value={lag} onChange={(e) => setLag(Number(e.target.value))} />
                </Field>
              </div>
              <Button onClick={add} disabled={pending || !predId}>
                <Plus className="h-4 w-4" /> Adicionar
              </Button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
