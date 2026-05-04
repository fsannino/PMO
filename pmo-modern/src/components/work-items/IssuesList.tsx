"use client";

import { useMemo, useState, useTransition } from "react";
import { Plus, Pencil, Trash2 } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select, Textarea } from "@/components/ui/Field";
import { StatusBadge } from "./StatusBadge";
import { WorkItemDrawer } from "./WorkItemDrawer";
import { CommentsThread, type CommentItem } from "./CommentsThread";
import { upsertIssue, deleteIssue } from "@/lib/actions/work-items";
import { Severities, Priorities, IssueStatuses } from "@/lib/enums";

export type IssueRow = {
  id: string;
  title: string;
  description: string | null;
  ownerId: string | null;
  taskId: string | null;
  areaId: string | null;
  severity: string;
  priority: string;
  status: string;
  openedAt: Date | string;
  closedAt: Date | string | null;
  owner: { name: string } | null;
  area: { code: string } | null;
  task: { name: string; wbs: string | null } | null;
  comments: CommentItem[];
};

type Lookup = { id: string; name: string; code?: string };

export function IssuesList({
  projectId,
  rows,
  users,
  areas,
  tasks,
  canWrite,
  currentUserId,
  isAdmin,
}: {
  projectId: string;
  rows: IssueRow[];
  users: Lookup[];
  areas: Lookup[];
  tasks: { id: string; name: string; wbs: string | null }[];
  canWrite: boolean;
  currentUserId: string;
  isAdmin: boolean;
}) {
  const [filter, setFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("");
  const [editing, setEditing] = useState<IssueRow | "new" | null>(null);
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
      taskId: (formData.get("taskId") as string) || null,
      areaId: (formData.get("areaId") as string) || null,
      severity: formData.get("severity") || "MEDIUM",
      priority: formData.get("priority") || "MEDIUM",
      status: formData.get("status") || "OPEN",
    };
    start(async () => {
      const res = await upsertIssue(data);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }

  function remove(it: IssueRow) {
    if (!confirm(`Excluir "${it.title}"?`)) return;
    start(async () => {
      const res = await deleteIssue(it.id);
      if (!res.ok) setError(res.error);
    });
  }

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <Input placeholder="Buscar…" className="max-w-xs" value={filter} onChange={(e) => setFilter(e.target.value)} />
        <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="max-w-[12rem]">
          <option value="">Todos os status</option>
          {IssueStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
        </Select>
        <span className="ml-auto text-xs text-slate-500">{filtered.length} de {rows.length}</span>
        {canWrite && (
          <Button onClick={() => setEditing("new")}>
            <Plus className="h-4 w-4" /> Nova issue
          </Button>
        )}
      </div>

      {error && <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>}

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Título</th>
              <th className="px-3 py-2 text-left">Severidade</th>
              <th className="px-3 py-2 text-left">Prioridade</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-left">Área</th>
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
                  {r.task && (
                    <div className="text-xs text-slate-500">↳ {r.task.wbs ?? ""} {r.task.name}</div>
                  )}
                </td>
                <td className="px-3 py-2"><StatusBadge value={r.severity} /></td>
                <td className="px-3 py-2"><StatusBadge value={r.priority} /></td>
                <td className="px-3 py-2"><StatusBadge value={r.status} /></td>
                <td className="px-3 py-2 text-xs text-slate-600">{r.owner?.name ?? "—"}</td>
                <td className="px-3 py-2 text-xs text-slate-600">{r.area?.code ?? "—"}</td>
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
            {filtered.length === 0 && <tr><td colSpan={8} className="px-4 py-6 text-center text-slate-500">Sem issues.</td></tr>}
          </tbody>
        </table>
      </div>

      {editing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <form action={submit} className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-lg bg-white shadow-xl">
            <div className="border-b px-5 py-3 font-semibold">{editing === "new" ? "Nova issue" : `Editar — ${editing.title}`}</div>
            <div className="space-y-3 p-5">
              <Field label="Título" required>
                <Input name="title" required defaultValue={editing === "new" ? "" : editing.title} maxLength={200} />
              </Field>
              <Field label="Descrição">
                <Textarea name="description" rows={3} defaultValue={editing === "new" ? "" : editing.description ?? ""} />
              </Field>
              <div className="grid gap-3 sm:grid-cols-3">
                <Field label="Severidade">
                  <Select name="severity" defaultValue={editing === "new" ? "MEDIUM" : editing.severity}>
                    {Severities.map((s) => <option key={s} value={s}>{s}</option>)}
                  </Select>
                </Field>
                <Field label="Prioridade">
                  <Select name="priority" defaultValue={editing === "new" ? "MEDIUM" : editing.priority}>
                    {Priorities.map((p) => <option key={p} value={p}>{p}</option>)}
                  </Select>
                </Field>
                <Field label="Status">
                  <Select name="status" defaultValue={editing === "new" ? "OPEN" : editing.status}>
                    {IssueStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
                  </Select>
                </Field>
              </div>
              <div className="grid gap-3 sm:grid-cols-3">
                <Field label="Owner">
                  <Select name="ownerId" defaultValue={editing === "new" ? "" : editing.ownerId ?? ""}>
                    <option value="">—</option>
                    {users.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                  </Select>
                </Field>
                <Field label="Área">
                  <Select name="areaId" defaultValue={editing === "new" ? "" : editing.areaId ?? ""}>
                    <option value="">—</option>
                    {areas.map((a) => <option key={a.id} value={a.id}>{a.code}</option>)}
                  </Select>
                </Field>
                <Field label="Tarefa">
                  <Select name="taskId" defaultValue={editing === "new" ? "" : editing.taskId ?? ""}>
                    <option value="">—</option>
                    {tasks.map((t) => <option key={t.id} value={t.id}>{t.wbs ? `${t.wbs} ` : ""}{t.name}</option>)}
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
          title={open.title}
          subtitle={
            <span className="flex items-center gap-2">
              <StatusBadge value={open.status} />
              <StatusBadge value={open.severity} />
              <StatusBadge value={open.priority} />
              {open.task && <span>· tarefa: {open.task.name}</span>}
            </span>
          }
          onClose={() => setOpenId(null)}
        >
          <div className="space-y-4">
            {open.description && <p className="whitespace-pre-wrap text-sm text-slate-700">{open.description}</p>}
            <dl className="grid grid-cols-2 gap-y-1 text-xs text-slate-600">
              <dt>Owner</dt><dd>{open.owner?.name ?? "—"}</dd>
              <dt>Área</dt><dd>{open.area?.code ?? "—"}</dd>
              <dt>Aberto em</dt><dd>{new Date(open.openedAt).toLocaleString("pt-BR")}</dd>
              {open.closedAt && (<><dt>Fechado em</dt><dd>{new Date(open.closedAt).toLocaleString("pt-BR")}</dd></>)}
            </dl>
            <CommentsThread
              comments={open.comments}
              target={{ issueId: open.id }}
              currentUserId={currentUserId}
              isAdmin={isAdmin}
            />
          </div>
        </WorkItemDrawer>
      )}
    </div>
  );
}
