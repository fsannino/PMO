"use client";

import { useMemo, useState, useTransition } from "react";
import { Plus, Pencil, Trash2 } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select, Textarea } from "@/components/ui/Field";
import { StatusBadge } from "./StatusBadge";
import { WorkItemDrawer } from "./WorkItemDrawer";
import { CommentsThread, type CommentItem } from "./CommentsThread";
import { upsertActionItem, deleteActionItem } from "@/lib/actions/work-items";
import { ActionStatuses, Priorities } from "@/lib/enums";

export type ActionRow = {
  id: string;
  title: string;
  description: string | null;
  ownerId: string | null;
  dueDate: Date | string | null;
  status: string;
  priority: string;
  owner: { name: string } | null;
  comments: CommentItem[];
};

type Lookup = { id: string; name: string };

function fmtDate(d: Date | string | null) {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("pt-BR");
}
function inputDate(d: Date | string | null) {
  if (!d) return "";
  return new Date(d).toISOString().slice(0, 10);
}
function isLate(due: Date | string | null, status: string) {
  if (!due) return false;
  if (status === "DONE" || status === "CANCELLED") return false;
  return new Date(due) < new Date();
}

export function ActionsList({
  projectId,
  rows,
  users,
  canWrite,
  currentUserId,
  isAdmin,
}: {
  projectId: string;
  rows: ActionRow[];
  users: Lookup[];
  canWrite: boolean;
  currentUserId: string;
  isAdmin: boolean;
}) {
  const [filter, setFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [editing, setEditing] = useState<ActionRow | "new" | null>(null);
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
      dueDate: (formData.get("dueDate") as string) || null,
      status: formData.get("status") || "OPEN",
      priority: formData.get("priority") || "MEDIUM",
    };
    start(async () => {
      const res = await upsertActionItem(data);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }

  function remove(it: ActionRow) {
    if (!confirm(`Excluir "${it.title}"?`)) return;
    start(async () => {
      const res = await deleteActionItem(it.id);
      if (!res.ok) setError(res.error);
    });
  }

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <Input placeholder="Buscar…" className="max-w-xs" value={filter} onChange={(e) => setFilter(e.target.value)} />
        <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="max-w-[12rem]">
          <option value="">Todos os status</option>
          {ActionStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
        </Select>
        <span className="ml-auto text-xs text-slate-500">{filtered.length} de {rows.length}</span>
        {canWrite && (
          <Button onClick={() => setEditing("new")}>
            <Plus className="h-4 w-4" /> Nova action
          </Button>
        )}
      </div>

      {error && <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>}

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Título</th>
              <th className="px-3 py-2 text-left">Prazo</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Prioridade</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-right">Coments</th>
              <th className="px-3 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.map((r) => {
              const late = isLate(r.dueDate, r.status);
              return (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-3 py-2">
                  <button onClick={() => setOpenId(r.id)} className="text-left font-medium text-brand-700 hover:underline">
                    {r.title}
                  </button>
                </td>
                <td className={`px-3 py-2 whitespace-nowrap ${late ? "text-rose-700 font-medium" : ""}`}>
                  {fmtDate(r.dueDate)}
                  {late && <span className="ml-1 text-xs">⚠</span>}
                </td>
                <td className="px-3 py-2"><StatusBadge value={r.status} /></td>
                <td className="px-3 py-2"><StatusBadge value={r.priority} /></td>
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
            );})}
            {filtered.length === 0 && <tr><td colSpan={7} className="px-4 py-6 text-center text-slate-500">Sem actions.</td></tr>}
          </tbody>
        </table>
      </div>

      {editing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <form action={submit} className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-lg bg-white shadow-xl">
            <div className="border-b px-5 py-3 font-semibold">{editing === "new" ? "Nova action" : `Editar — ${editing.title}`}</div>
            <div className="space-y-3 p-5">
              <Field label="Título" required>
                <Input name="title" required defaultValue={editing === "new" ? "" : editing.title} maxLength={200} />
              </Field>
              <Field label="Descrição">
                <Textarea name="description" rows={3} defaultValue={editing === "new" ? "" : editing.description ?? ""} />
              </Field>
              <div className="grid gap-3 sm:grid-cols-3">
                <Field label="Prazo">
                  <Input name="dueDate" type="date" defaultValue={editing === "new" ? "" : inputDate(editing.dueDate)} />
                </Field>
                <Field label="Status">
                  <Select name="status" defaultValue={editing === "new" ? "OPEN" : editing.status}>
                    {ActionStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
                  </Select>
                </Field>
                <Field label="Prioridade">
                  <Select name="priority" defaultValue={editing === "new" ? "MEDIUM" : editing.priority}>
                    {Priorities.map((p) => <option key={p} value={p}>{p}</option>)}
                  </Select>
                </Field>
              </div>
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
          subtitle={<span className="flex items-center gap-2"><StatusBadge value={open.status} /><StatusBadge value={open.priority} />{open.dueDate && <span>· prazo: {fmtDate(open.dueDate)}</span>}</span>}
          onClose={() => setOpenId(null)}
        >
          <div className="space-y-4">
            {open.description && <p className="whitespace-pre-wrap text-sm text-slate-700">{open.description}</p>}
            <dl className="grid grid-cols-2 gap-y-1 text-xs text-slate-600">
              <dt>Owner</dt><dd>{open.owner?.name ?? "—"}</dd>
              <dt>Prazo</dt><dd>{fmtDate(open.dueDate)}</dd>
            </dl>
            <CommentsThread comments={open.comments} target={{ actionId: open.id }} currentUserId={currentUserId} isAdmin={isAdmin} />
          </div>
        </WorkItemDrawer>
      )}
    </div>
  );
}
