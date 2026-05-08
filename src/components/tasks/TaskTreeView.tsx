"use client";

import { useMemo, useState, useTransition } from "react";
import {
  ChevronDown,
  ChevronRight,
  Pencil,
  Plus,
  Trash2,
  Undo2,
  Diamond,
  History as HistoryIcon,
  GitBranch,
} from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Input } from "@/components/ui/Field";
import { TaskFormDialog, type TaskFormDefaults } from "./TaskFormDialog";
import { TaskHistoryDrawer } from "./TaskHistoryDrawer";
import { TaskDependencyDrawer } from "./TaskDependencyDrawer";
import { upsertTask, softDeleteTask, restoreTask } from "@/lib/actions/tasks";
import { buildTree, flattenTree, type TaskNodeBase } from "@/lib/task-tree";
import { cn } from "@/lib/utils";

export type TaskRow = TaskNodeBase & {
  externalId: string | null;
  description: string | null;
  baselineStart: Date | string | null;
  baselineEnd: Date | string | null;
  assigneeId: string | null;
  areaId: string | null;
  frenteId: string | null;
  equipeId: string | null;
  assignee: { name: string } | null;
  equipe: { code: string } | null;
  area: { code: string } | null;
  predecessors: { id: string; predecessorId: string; predecessor: { name: string; wbs: string | null } }[];
};

export type Lookup = { id: string; code?: string; name: string; email?: string };

const STATUS_COLORS: Record<string, string> = {
  NOT_STARTED: "bg-slate-100 text-slate-600",
  IN_PROGRESS: "bg-blue-100 text-blue-700",
  DELAYED: "bg-amber-100 text-amber-700",
  DONE: "bg-emerald-100 text-emerald-700",
  CANCELLED: "bg-rose-100 text-rose-700",
};

function fmtDate(d: Date | string): string {
  return new Date(d).toLocaleDateString("pt-BR");
}

export function TaskTreeView({
  projectId,
  rows,
  users,
  areas,
  frentes,
  equipes,
  canWrite,
}: {
  projectId: string;
  rows: TaskRow[];
  users: Lookup[];
  areas: Lookup[];
  frentes: Lookup[];
  equipes: Lookup[];
  canWrite: boolean;
}) {
  const [showDeleted, setShowDeleted] = useState(false);
  const [filter, setFilter] = useState("");
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const [editing, setEditing] = useState<TaskFormDefaults | "new" | null>(null);
  const [historyTask, setHistoryTask] = useState<TaskRow | null>(null);
  const [depTask, setDepTask] = useState<TaskRow | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();

  const filteredRows = useMemo(() => {
    return rows.filter((r) => {
      if (!showDeleted && r.deletedAt) return false;
      if (filter && !r.name.toLowerCase().includes(filter.toLowerCase())) return false;
      return true;
    });
  }, [rows, showDeleted, filter]);

  const tree = useMemo(() => buildTree(filteredRows), [filteredRows]);
  const flat = useMemo(() => flattenTree(tree), [tree]);

  // visíveis = não estão sob algum colapsado
  const visible = useMemo(() => {
    const hiddenAncestors = new Set<string>();
    return flat.filter((n) => {
      // se algum ancestral está em collapsed, esta linha fica oculta
      let cur = n.parentId;
      while (cur) {
        if (collapsed.has(cur)) return false;
        const p = flat.find((x) => x.id === cur);
        cur = p?.parentId ?? null;
      }
      void hiddenAncestors;
      return true;
    });
  }, [flat, collapsed]);

  function toggle(id: string) {
    setCollapsed((s) => {
      const next = new Set(s);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function quickPercent(task: TaskRow, value: number) {
    setError(null);
    start(async () => {
      const res = await upsertTask({
        id: task.id,
        projectId,
        parentId: task.parentId,
        wbs: task.wbs,
        externalId: task.externalId,
        name: task.name,
        description: task.description,
        startDate: new Date(task.startDate).toISOString().slice(0, 10),
        endDate: new Date(task.endDate).toISOString().slice(0, 10),
        baselineStart: task.baselineStart ? new Date(task.baselineStart).toISOString().slice(0, 10) : null,
        baselineEnd: task.baselineEnd ? new Date(task.baselineEnd).toISOString().slice(0, 10) : null,
        percentDone: value,
        isMilestone: task.isMilestone,
        isSummary: task.isSummary,
        assigneeId: task.assigneeId,
        areaId: task.areaId,
        frenteId: task.frenteId,
        equipeId: task.equipeId,
      });
      if (!res.ok) setError(res.error);
    });
  }

  function handleDelete(task: TaskRow) {
    if (!confirm(`Excluir tarefa "${task.name}"? Pode ser restaurada.`)) return;
    setError(null);
    start(async () => {
      const res = await softDeleteTask(task.id);
      if (!res.ok) setError(res.error);
    });
  }

  function handleRestore(task: TaskRow) {
    setError(null);
    start(async () => {
      const res = await restoreTask(task.id);
      if (!res.ok) setError(res.error);
    });
  }

  const taskMap = useMemo(() => new Map(rows.map((r) => [r.id, r])), [rows]);

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <Input
          placeholder="Buscar tarefa…"
          className="max-w-xs"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        />
        <label className="ml-2 flex items-center gap-1 text-sm text-slate-600">
          <input
            type="checkbox"
            checked={showDeleted}
            onChange={(e) => setShowDeleted(e.target.checked)}
          />
          incluir excluídas
        </label>
        <div className="ml-auto flex gap-2">
          {canWrite && (
            <Button onClick={() => setEditing("new")} disabled={pending}>
              <Plus className="h-4 w-4" /> Nova tarefa
            </Button>
          )}
        </div>
      </div>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">WBS</th>
              <th className="px-3 py-2 text-left">Tarefa</th>
              <th className="px-3 py-2 text-left">Início</th>
              <th className="px-3 py-2 text-left">Fim</th>
              <th className="px-3 py-2 text-left">Dur.</th>
              <th className="px-3 py-2 text-left">%</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Resp.</th>
              <th className="px-3 py-2 text-left">Equipe</th>
              <th className="px-3 py-2 text-left">Pred.</th>
              <th className="px-3 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {visible.length === 0 && (
              <tr>
                <td colSpan={11} className="px-4 py-8 text-center text-slate-500">
                  Nenhuma tarefa.
                </td>
              </tr>
            )}
            {visible.map((n) => {
              const task = n;
              const hasChildren = n.children.length > 0;
              const collapsedHere = collapsed.has(n.id);
              return (
                <tr
                  key={n.id}
                  className={cn(
                    "hover:bg-slate-50",
                    n.deletedAt && "opacity-50",
                    n.isSummary && "font-medium",
                  )}
                >
                  <td className="px-3 py-1 font-mono text-xs">{n.wbs ?? "—"}</td>
                  <td className="px-3 py-1">
                    <div
                      className="flex items-center gap-1"
                      style={{ paddingLeft: `${n.depth * 14}px` }}
                    >
                      {hasChildren ? (
                        <button
                          onClick={() => toggle(n.id)}
                          className="text-slate-400 hover:text-slate-700"
                        >
                          {collapsedHere ? (
                            <ChevronRight className="h-4 w-4" />
                          ) : (
                            <ChevronDown className="h-4 w-4" />
                          )}
                        </button>
                      ) : (
                        <span className="inline-block w-4" />
                      )}
                      {n.isMilestone && <Diamond className="h-3 w-3 text-amber-600" fill="currentColor" />}
                      <span className={cn(n.isSummary && "text-brand-700")}>{n.name}</span>
                    </div>
                  </td>
                  <td className="px-3 py-1 whitespace-nowrap">{fmtDate(n.startDate)}</td>
                  <td className="px-3 py-1 whitespace-nowrap">{fmtDate(n.endDate)}</td>
                  <td className="px-3 py-1 text-right">{n.durationDays ?? "—"}d</td>
                  <td className="px-3 py-1">
                    <PercentCell
                      value={n.percentDone}
                      disabled={!canWrite || !!n.deletedAt || n.isSummary}
                      onChange={(v) => quickPercent(taskMap.get(n.id)!, v)}
                    />
                  </td>
                  <td className="px-3 py-1">
                    <span className={cn("rounded px-2 py-0.5 text-xs", STATUS_COLORS[n.status] ?? "bg-slate-100")}>
                      {n.status}
                    </span>
                  </td>
                  <td className="px-3 py-1 text-xs text-slate-600">
                    {taskMap.get(n.id)?.assignee?.name ?? "—"}
                  </td>
                  <td className="px-3 py-1 text-xs text-slate-600">
                    {taskMap.get(n.id)?.equipe?.code ?? "—"}
                  </td>
                  <td className="px-3 py-1 text-xs">
                    {taskMap.get(n.id)?.predecessors.length ? (
                      <span className="rounded bg-slate-100 px-1.5 py-0.5 font-mono">
                        {taskMap.get(n.id)?.predecessors.length}
                      </span>
                    ) : "—"}
                  </td>
                  <td className="px-3 py-1 text-right">
                    <div className="flex justify-end gap-0.5">
                      <Button size="sm" variant="ghost" title="Histórico" onClick={() => setHistoryTask(taskMap.get(n.id)!)}>
                        <HistoryIcon className="h-3.5 w-3.5" />
                      </Button>
                      {canWrite && !n.deletedAt && (
                        <>
                          <Button size="sm" variant="ghost" title="Predecessores" onClick={() => setDepTask(taskMap.get(n.id)!)}>
                            <GitBranch className="h-3.5 w-3.5" />
                          </Button>
                          <Button
                            size="sm"
                            variant="ghost"
                            title="Editar"
                            onClick={() => setEditing(taskToFormDefaults(taskMap.get(n.id)!))}
                          >
                            <Pencil className="h-3.5 w-3.5" />
                          </Button>
                          <Button
                            size="sm"
                            variant="ghost"
                            title="Excluir"
                            onClick={() => handleDelete(taskMap.get(n.id)!)}
                          >
                            <Trash2 className="h-3.5 w-3.5 text-rose-600" />
                          </Button>
                        </>
                      )}
                      {canWrite && n.deletedAt && (
                        <Button
                          size="sm"
                          variant="ghost"
                          title="Restaurar"
                          onClick={() => handleRestore(taskMap.get(n.id)!)}
                        >
                          <Undo2 className="h-3.5 w-3.5 text-emerald-600" />
                        </Button>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {editing && (
        <TaskFormDialog
          projectId={projectId}
          tasks={rows}
          users={users}
          areas={areas}
          frentes={frentes}
          equipes={equipes}
          defaults={editing === "new" ? null : editing}
          onClose={() => setEditing(null)}
        />
      )}

      {historyTask && (
        <TaskHistoryDrawer
          taskId={historyTask.id}
          taskName={historyTask.name}
          onClose={() => setHistoryTask(null)}
        />
      )}

      {depTask && (
        <TaskDependencyDrawer
          task={depTask}
          allTasks={rows.filter((r) => !r.deletedAt && r.id !== depTask.id && !r.isSummary)}
          canWrite={canWrite}
          onClose={() => setDepTask(null)}
        />
      )}
    </div>
  );
}

function PercentCell({
  value,
  disabled,
  onChange,
}: {
  value: number;
  disabled: boolean;
  onChange: (v: number) => void;
}) {
  const [v, setV] = useState(String(Math.round(value)));
  const [editing, setEditing] = useState(false);

  if (disabled || !editing) {
    return (
      <button
        type="button"
        onClick={() => !disabled && setEditing(true)}
        className={cn(
          "block w-14 cursor-text rounded text-right tabular-nums",
          disabled && "cursor-not-allowed text-slate-400",
        )}
      >
        {Math.round(value)}%
      </button>
    );
  }
  return (
    <input
      type="number"
      autoFocus
      min={0}
      max={100}
      value={v}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => {
        setEditing(false);
        const n = Number(v);
        if (!isNaN(n) && n !== value) onChange(Math.max(0, Math.min(100, n)));
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        if (e.key === "Escape") {
          setV(String(Math.round(value)));
          setEditing(false);
        }
      }}
      className="w-14 rounded border border-slate-300 px-1 text-right tabular-nums"
    />
  );
}

function taskToFormDefaults(task: TaskRow): TaskFormDefaults {
  return {
    id: task.id,
    parentId: task.parentId,
    wbs: task.wbs,
    externalId: task.externalId,
    name: task.name,
    description: task.description,
    startDate: new Date(task.startDate).toISOString().slice(0, 10),
    endDate: new Date(task.endDate).toISOString().slice(0, 10),
    baselineStart: task.baselineStart ? new Date(task.baselineStart).toISOString().slice(0, 10) : null,
    baselineEnd: task.baselineEnd ? new Date(task.baselineEnd).toISOString().slice(0, 10) : null,
    percentDone: task.percentDone,
    isMilestone: task.isMilestone,
    isSummary: task.isSummary,
    assigneeId: task.assigneeId,
    areaId: task.areaId,
    frenteId: task.frenteId,
    equipeId: task.equipeId,
  };
}
