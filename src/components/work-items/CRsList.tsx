"use client";

import { useMemo, useState, useTransition } from "react";
import { Plus, Pencil, Trash2, ListPlus } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select, Textarea } from "@/components/ui/Field";
import { StatusBadge } from "./StatusBadge";
import { WorkItemDrawer } from "./WorkItemDrawer";
import { CommentsThread, type CommentItem } from "./CommentsThread";
import { upsertChangeRequest, deleteChangeRequest, addCRIncrement } from "@/lib/actions/work-items";
import { ChangeRequestStatuses, Priorities } from "@/lib/enums";

export type CRRow = {
  id: string;
  name: string;
  description: string | null;
  ownerId: string | null;
  comiteId: string | null;
  priority: string;
  status: string;
  dueDate: Date | string | null;
  owner: { name: string } | null;
  comite: { code: string; name: string } | null;
  comments: CommentItem[];
  increments: { id: string; sequence: number; description: string; appliedAt: Date | string | null }[];
};

type Lookup = { id: string; name: string; code?: string };

function fmtDate(d: Date | string | null) { return d ? new Date(d).toLocaleDateString("pt-BR") : "—"; }
function inputDate(d: Date | string | null) { return d ? new Date(d).toISOString().slice(0, 10) : ""; }

export function CRsList({
  projectId,
  rows,
  users,
  comites,
  canWrite,
  currentUserId,
  isAdmin,
}: {
  projectId: string;
  rows: CRRow[];
  users: Lookup[];
  comites: Lookup[];
  canWrite: boolean;
  currentUserId: string;
  isAdmin: boolean;
}) {
  const [filter, setFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [editing, setEditing] = useState<CRRow | "new" | null>(null);
  const [openId, setOpenId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [newIncrement, setNewIncrement] = useState("");
  const [pending, start] = useTransition();

  const filtered = useMemo(() => rows.filter((r) => {
    if (filter && !r.name.toLowerCase().includes(filter.toLowerCase())) return false;
    if (statusFilter && r.status !== statusFilter) return false;
    return true;
  }), [rows, filter, statusFilter]);

  const open = openId ? rows.find((r) => r.id === openId) : null;

  function submit(formData: FormData) {
    setError(null);
    const data = {
      id: editing !== "new" ? editing?.id : undefined,
      projectId,
      name: formData.get("name"),
      description: (formData.get("description") as string) || null,
      ownerId: (formData.get("ownerId") as string) || null,
      comiteId: (formData.get("comiteId") as string) || null,
      priority: formData.get("priority") || "MEDIUM",
      status: formData.get("status") || "OPEN",
      dueDate: (formData.get("dueDate") as string) || null,
    };
    start(async () => {
      const res = await upsertChangeRequest(data);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }
  function remove(it: CRRow) {
    if (!confirm(`Excluir "${it.name}"?`)) return;
    start(async () => {
      const res = await deleteChangeRequest(it.id);
      if (!res.ok) setError(res.error);
    });
  }
  function addInc() {
    if (!open || !newIncrement.trim()) return;
    setError(null);
    start(async () => {
      const res = await addCRIncrement({ changeRequestId: open.id, description: newIncrement });
      if (res.ok) setNewIncrement("");
      else setError(res.error);
    });
  }

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <Input placeholder="Buscar…" className="max-w-xs" value={filter} onChange={(e) => setFilter(e.target.value)} />
        <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="max-w-[14rem]">
          <option value="">Todos os status</option>
          {ChangeRequestStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
        </Select>
        <span className="ml-auto text-xs text-slate-500">{filtered.length} de {rows.length}</span>
        {canWrite && (
          <Button onClick={() => setEditing("new")}>
            <Plus className="h-4 w-4" /> Novo CR
          </Button>
        )}
      </div>

      {error && <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>}

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Nome</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Prioridade</th>
              <th className="px-3 py-2 text-left">Comitê</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-left">Prazo</th>
              <th className="px-3 py-2 text-right">Inc.</th>
              <th className="px-3 py-2 text-right">Coments</th>
              <th className="px-3 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.map((r) => (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-3 py-2">
                  <button onClick={() => setOpenId(r.id)} className="text-left font-medium text-brand-700 hover:underline">
                    {r.name}
                  </button>
                </td>
                <td className="px-3 py-2"><StatusBadge value={r.status} /></td>
                <td className="px-3 py-2"><StatusBadge value={r.priority} /></td>
                <td className="px-3 py-2 text-xs text-slate-600">{r.comite?.code ?? "—"}</td>
                <td className="px-3 py-2 text-xs text-slate-600">{r.owner?.name ?? "—"}</td>
                <td className="px-3 py-2 whitespace-nowrap">{fmtDate(r.dueDate)}</td>
                <td className="px-3 py-2 text-right tabular-nums text-xs">{r.increments.length}</td>
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
            {filtered.length === 0 && <tr><td colSpan={9} className="px-4 py-6 text-center text-slate-500">Sem CRs.</td></tr>}
          </tbody>
        </table>
      </div>

      {editing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <form action={submit} className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-lg bg-white shadow-xl">
            <div className="border-b px-5 py-3 font-semibold">{editing === "new" ? "Novo CR" : `Editar — ${editing.name}`}</div>
            <div className="space-y-3 p-5">
              <Field label="Nome" required>
                <Input name="name" required defaultValue={editing === "new" ? "" : editing.name} maxLength={200} />
              </Field>
              <Field label="Descrição">
                <Textarea name="description" rows={3} defaultValue={editing === "new" ? "" : editing.description ?? ""} />
              </Field>
              <div className="grid gap-3 sm:grid-cols-3">
                <Field label="Status">
                  <Select name="status" defaultValue={editing === "new" ? "OPEN" : editing.status}>
                    {ChangeRequestStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
                  </Select>
                </Field>
                <Field label="Prioridade">
                  <Select name="priority" defaultValue={editing === "new" ? "MEDIUM" : editing.priority}>
                    {Priorities.map((p) => <option key={p} value={p}>{p}</option>)}
                  </Select>
                </Field>
                <Field label="Prazo">
                  <Input name="dueDate" type="date" defaultValue={editing === "new" ? "" : inputDate(editing.dueDate)} />
                </Field>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <Field label="Owner">
                  <Select name="ownerId" defaultValue={editing === "new" ? "" : editing.ownerId ?? ""}>
                    <option value="">—</option>
                    {users.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                  </Select>
                </Field>
                <Field label="Comitê">
                  <Select name="comiteId" defaultValue={editing === "new" ? "" : editing.comiteId ?? ""}>
                    <option value="">—</option>
                    {comites.map((c) => <option key={c.id} value={c.id}>{c.code} — {c.name}</option>)}
                  </Select>
                </Field>
              </div>
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
          title={open.name}
          subtitle={<span className="flex items-center gap-2"><StatusBadge value={open.status} /><StatusBadge value={open.priority} />{open.comite && <span>· comitê: {open.comite.code}</span>}</span>}
          onClose={() => setOpenId(null)}
        >
          <div className="space-y-4">
            {open.description && <p className="whitespace-pre-wrap text-sm text-slate-700">{open.description}</p>}

            <div>
              <h4 className="mb-2 text-sm font-semibold uppercase tracking-wide text-slate-500">
                Incrementos ({open.increments.length})
              </h4>
              <ul className="space-y-1">
                {open.increments.length === 0 && <li className="rounded border border-dashed p-2 text-center text-xs text-slate-500">Sem incrementos.</li>}
                {open.increments.map((inc) => (
                  <li key={inc.id} className="rounded border bg-slate-50 p-2 text-sm">
                    <div className="flex items-center gap-2 text-xs text-slate-500">
                      <span className="font-mono">#{inc.sequence}</span>
                      {inc.appliedAt && <span>· aplicado em {new Date(inc.appliedAt).toLocaleDateString("pt-BR")}</span>}
                    </div>
                    <p className="mt-0.5 whitespace-pre-wrap">{inc.description}</p>
                  </li>
                ))}
              </ul>
              {canWrite && (
                <div className="mt-2 flex gap-2">
                  <Input
                    value={newIncrement}
                    onChange={(e) => setNewIncrement(e.target.value)}
                    placeholder="Descreva o próximo incremento…"
                    maxLength={2000}
                  />
                  <Button size="sm" onClick={addInc} disabled={pending || !newIncrement.trim()}>
                    <ListPlus className="h-3.5 w-3.5" />
                  </Button>
                </div>
              )}
            </div>

            <CommentsThread comments={open.comments} target={{ changeRequestId: open.id }} currentUserId={currentUserId} isAdmin={isAdmin} />
          </div>
        </WorkItemDrawer>
      )}
    </div>
  );
}
