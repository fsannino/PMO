"use client";

import { useEffect, useState } from "react";
import { X, History as HistoryIcon } from "lucide-react";

type HistoryEntry = {
  id: string;
  action: string;
  field: string | null;
  oldValue: string | null;
  newValue: string | null;
  note: string | null;
  createdAt: string;
  user: { name: string } | null;
};

const ACTION_BADGE: Record<string, string> = {
  CREATED: "bg-emerald-100 text-emerald-700",
  UPDATED: "bg-blue-100 text-blue-700",
  DELETED: "bg-rose-100 text-rose-700",
  RESTORED: "bg-amber-100 text-amber-700",
  IMPORTED: "bg-violet-100 text-violet-700",
  MEASURED: "bg-indigo-100 text-indigo-700",
};

function fmt(s: string) {
  return new Date(s).toLocaleString("pt-BR");
}

function shortVal(v: string | null) {
  if (!v) return "—";
  // se parecer ISO date longa, encurta
  if (/^\d{4}-\d{2}-\d{2}T/.test(v)) return new Date(v).toLocaleDateString("pt-BR");
  if (v.length > 60) return v.slice(0, 60) + "…";
  return v;
}

export function TaskHistoryDrawer({
  taskId,
  taskName,
  onClose,
}: {
  taskId: string;
  taskName: string;
  onClose: () => void;
}) {
  const [entries, setEntries] = useState<HistoryEntry[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setEntries(null);
    setError(null);
    fetch(`/api/tasks/${taskId}/history`)
      .then((r) => r.json())
      .then((data) => setEntries(data.entries))
      .catch((e) => setError(String(e)));
  }, [taskId]);

  return (
    <div className="fixed inset-0 z-40 flex justify-end bg-black/30">
      <div className="flex h-full w-full max-w-md flex-col bg-white shadow-xl">
        <div className="flex items-center justify-between border-b px-4 py-3">
          <div className="flex items-center gap-2">
            <HistoryIcon className="h-4 w-4 text-slate-500" />
            <h3 className="font-semibold">Histórico — {taskName}</h3>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700">
            <X className="h-4 w-4" />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto p-4">
          {error && <p className="text-sm text-rose-700">Erro: {error}</p>}
          {!entries && !error && <p className="text-sm text-slate-500">Carregando…</p>}
          {entries && entries.length === 0 && (
            <p className="text-sm text-slate-500">Sem registros de histórico.</p>
          )}
          <ul className="space-y-3">
            {entries?.map((e) => (
              <li key={e.id} className="rounded border border-slate-200 p-3 text-sm">
                <div className="flex items-center justify-between gap-2">
                  <span
                    className={`inline-block rounded px-2 py-0.5 text-xs font-medium ${
                      ACTION_BADGE[e.action] ?? "bg-slate-100"
                    }`}
                  >
                    {e.action}
                  </span>
                  <span className="text-xs text-slate-500">{fmt(e.createdAt)}</span>
                </div>
                {e.note && <p className="mt-1.5 text-slate-700">{e.note}</p>}
                {e.field && (
                  <p className="mt-1.5 text-xs text-slate-600">
                    <strong>{e.field}:</strong>{" "}
                    <span className="text-rose-600 line-through">{shortVal(e.oldValue)}</span>
                    <span className="mx-1">→</span>
                    <span className="text-emerald-700">{shortVal(e.newValue)}</span>
                  </p>
                )}
                <p className="mt-1.5 text-xs text-slate-500">
                  por {e.user?.name ?? "—"}
                </p>
              </li>
            ))}
          </ul>
        </div>
      </div>
    </div>
  );
}
