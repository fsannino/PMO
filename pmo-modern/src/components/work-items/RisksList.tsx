"use client";

import { useMemo, useState, useTransition } from "react";
import { Plus, Pencil, Trash2 } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select, Textarea } from "@/components/ui/Field";
import { StatusBadge } from "./StatusBadge";
import { WorkItemDrawer } from "./WorkItemDrawer";
import { CommentsThread, type CommentItem } from "./CommentsThread";
import { upsertRisk, deleteRisk } from "@/lib/actions/work-items";
import { RiskStatuses } from "@/lib/enums";

export type RiskRow = {
  id: string;
  title: string;
  description: string | null;
  ownerId: string | null;
  probability: number;
  impact: number;
  exposure: number;
  mitigation: string | null;
  status: string;
  owner: { name: string } | null;
  comments: CommentItem[];
};

type Lookup = { id: string; name: string };

function expoBadge(e: number) {
  if (e >= 0.5) return "bg-rose-100 text-rose-800";
  if (e >= 0.25) return "bg-amber-100 text-amber-800";
  return "bg-slate-100 text-slate-700";
}

export function RisksList({
  projectId,
  rows,
  users,
  canWrite,
  currentUserId,
  isAdmin,
}: {
  projectId: string;
  rows: RiskRow[];
  users: Lookup[];
  canWrite: boolean;
  currentUserId: string;
  isAdmin: boolean;
}) {
  const [filter, setFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [editing, setEditing] = useState<RiskRow | "new" | null>(null);
  const [openId, setOpenId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();

  const filtered = useMemo(() => rows.filter((r) => {
    if (filter && !r.title.toLowerCase().includes(filter.toLowerCase())) return false;
    if (statusFilter && r.status !== statusFilter) return false;
    return true;
  }), [rows, filter, statusFilter]);

  const open = openId ? rows.find((r) => r.id === openId) : null;

  function submit(formData: FormData) {
    setError(null);
    const data = {
      id: editing !== "new" ? editing?.id : undefined,
      projectId,
      title: formData.get("title"),
      description: (formData.get("description") as string) || null,
      ownerId: (formData.get("ownerId") as string) || null,
      probability: Number(formData.get("probability") ?? 0.5),
      impact: Number(formData.get("impact") ?? 0.5),
      mitigation: (formData.get("mitigation") as string) || null,
      status: formData.get("status") || "IDENTIFIED",
    };
    start(async () => {
      const res = await upsertRisk(data);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }

  function remove(it: RiskRow) {
    if (!confirm(`Excluir "${it.title}"?`)) return;
    start(async () => {
      const res = await deleteRisk(it.id);
      if (!res.ok) setError(res.error);
    });
  }

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <Input placeholder="Buscar…" className="max-w-xs" value={filter} onChange={(e) => setFilter(e.target.value)} />
        <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="max-w-[12rem]">
          <option value="">Todos os status</option>
          {RiskStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
        </Select>
        <span className="ml-auto text-xs text-slate-500">{filtered.length} de {rows.length}</span>
        {canWrite && (
          <Button onClick={() => setEditing("new")}>
            <Plus className="h-4 w-4" /> Novo risco
          </Button>
        )}
      </div>

      {error && <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>}

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Título</th>
              <th className="px-3 py-2 text-right">Prob.</th>
              <th className="px-3 py-2 text-right">Imp.</th>
              <th className="px-3 py-2 text-right">Expo.</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-right">Coments</th>
              <th className="px-3 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.map((r) => (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-3 py-2">
                  <button onClick={() => setOpenId(r.id)} className="text-left font-medium text-brand-700 hover:underline">
                    {r.title}
                  </button>
                </td>
                <td className="px-3 py-2 text-right tabular-nums">{r.probability.toFixed(2)}</td>
                <td className="px-3 py-2 text-right tabular-nums">{r.impact.toFixed(2)}</td>
                <td className="px-3 py-2 text-right">
                  <span className={`inline-block rounded px-2 py-0.5 text-xs font-medium tabular-nums ${expoBadge(r.exposure)}`}>
                    {r.exposure.toFixed(2)}
                  </span>
                </td>
                <td className="px-3 py-2"><StatusBadge value={r.status} /></td>
                <td className="px-3 py-2 text-xs text-slate-600">{r.owner?.name ?? "—"}</td>
                <td className="px-3 py-2 text-right tabular-nums text-xs">{r.comments.length}</td>
                <td className="px-3 py-2 text-right">
                  {canWrite && (
                    <div className="flex justify-end gap-1">
                      <Button size="sm" variant="ghost" onClick={() => setEditing(r)}><Pencil className="h-3.5 w-3.5" /></Button>
                      <Button size="sm" variant="ghost" onClick={() => remove(r)}><Trash2 className="h-3.5 w-3.5 text-rose-600" /></Button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={8} className="px-4 py-6 text-center text-slate-500">Sem riscos.</td></tr>}
          </tbody>
        </table>
      </div>

      {editing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <form action={submit} className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-lg bg-white shadow-xl">
            <div className="border-b px-5 py-3 font-semibold">{editing === "new" ? "Novo risco" : `Editar — ${editing.title}`}</div>
            <div className="space-y-3 p-5">
              <Field label="Título" required>
                <Input name="title" required defaultValue={editing === "new" ? "" : editing.title} maxLength={200} />
              </Field>
              <Field label="Descrição">
                <Textarea name="description" rows={3} defaultValue={editing === "new" ? "" : editing.description ?? ""} />
              </Field>
              <div className="grid gap-3 sm:grid-cols-3">
                <Field label="Probabilidade (0..1)" required>
                  <Input name="probability" type="number" min={0} max={1} step="0.05" defaultValue={editing === "new" ? 0.5 : editing.probability} />
                </Field>
                <Field label="Impacto (0..1)" required>
                  <Input name="impact" type="number" min={0} max={1} step="0.05" defaultValue={editing === "new" ? 0.5 : editing.impact} />
                </Field>
                <Field label="Status">
                  <Select name="status" defaultValue={editing === "new" ? "IDENTIFIED" : editing.status}>
                    {RiskStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
                  </Select>
                </Field>
              </div>
              <Field label="Mitigação">
                <Textarea name="mitigation" rows={2} defaultValue={editing === "new" ? "" : editing.mitigation ?? ""} />
              </Field>
              <Field label="Owner">
                <Select name="ownerId" defaultValue={editing === "new" ? "" : editing.ownerId ?? ""}>
                  <option value="">—</option>
                  {users.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                </Select>
              </Field>
            </div>
            <div className="flex justify-end gap-2 border-t bg-slate-50 px-5 py-3">
              <Button type="button" variant="secondary" onClick={() => setEditing(null)} disabled={pending}>Cancelar</Button>
              <Button type="submit" disabled={pending}>{pending ? "Salvando…" : "Salvar"}</Button>
            </div>
          </form>
        </div>
      )}

      {open && (
        <WorkItemDrawer
          title={open.title}
          subtitle={
            <span className="flex items-center gap-2">
              <StatusBadge value={open.status} />
              <span className="font-mono">P×I={open.exposure.toFixed(2)}</span>
            </span>
          }
          onClose={() => setOpenId(null)}
        >
          <div className="space-y-4">
            {open.description && <p className="whitespace-pre-wrap text-sm text-slate-700">{open.description}</p>}
            {open.mitigation && (
              <div className="rounded border border-emerald-200 bg-emerald-50 p-3 text-sm">
                <div className="text-xs font-semibold uppercase text-emerald-800">Mitigação</div>
                <p className="mt-1 whitespace-pre-wrap text-emerald-900">{open.mitigation}</p>
              </div>
            )}
            <dl className="grid grid-cols-3 gap-y-1 text-xs text-slate-600">
              <dt>Probabilidade</dt><dd className="col-span-2 tabular-nums">{open.probability.toFixed(2)}</dd>
              <dt>Impacto</dt><dd className="col-span-2 tabular-nums">{open.impact.toFixed(2)}</dd>
              <dt>Exposição</dt><dd className="col-span-2 tabular-nums">{open.exposure.toFixed(2)}</dd>
              <dt>Owner</dt><dd className="col-span-2">{open.owner?.name ?? "—"}</dd>
            </dl>
            <CommentsThread comments={open.comments} target={{ riskId: open.id }} currentUserId={currentUserId} isAdmin={isAdmin} />
          </div>
        </WorkItemDrawer>
      )}
    </div>
  );
}
