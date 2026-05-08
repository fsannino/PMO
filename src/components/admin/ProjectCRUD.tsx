"use client";

import { useState, useTransition } from "react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select, Textarea } from "@/components/ui/Field";
import { Pencil, Plus, Trash2, X } from "lucide-react";
import { upsertProject, deleteProject } from "@/lib/actions/projects";
import { Modules, ProjectStatuses } from "@/lib/enums";

type Lookup = { id: string; code: string; name: string };
type User = { id: string; name: string; email: string };
export type ProjectRow = {
  id: string;
  code: string;
  name: string;
  description: string | null;
  module: string;
  status: string;
  priority: number;
  ownerId: string;
  unidadeId: string | null;
  governancaId: string | null;
  startDate: Date | string;
  endDate: Date | string;
  baselineDate: Date | string | null;
  owner: { name: string };
  unidade: { code: string } | null;
};

function fmtDate(d: Date | string | null): string {
  if (!d) return "";
  const date = typeof d === "string" ? new Date(d) : d;
  return date.toISOString().slice(0, 10);
}

function fmtBR(d: Date | string | null): string {
  if (!d) return "—";
  const date = typeof d === "string" ? new Date(d) : d;
  return date.toLocaleDateString("pt-BR");
}

export function ProjectCRUD({
  rows,
  users,
  unidades,
  governancas,
}: {
  rows: ProjectRow[];
  users: User[];
  unidades: Lookup[];
  governancas: Lookup[];
}) {
  const [editing, setEditing] = useState<ProjectRow | "new" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();
  const [filter, setFilter] = useState("");

  const filtered = rows.filter(
    (r) =>
      r.code.toLowerCase().includes(filter.toLowerCase()) ||
      r.name.toLowerCase().includes(filter.toLowerCase()),
  );

  function submit(formData: FormData) {
    setError(null);
    start(async () => {
      const res = await upsertProject(formData);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }

  function handleDelete(id: string, code: string) {
    if (!confirm(`Excluir projeto "${code}"? Esta ação irá apagar tarefas, issues, riscos e CRs vinculados.`)) return;
    setError(null);
    start(async () => {
      const res = await deleteProject(id);
      if (!res.ok) setError(res.error);
    });
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2">
        <Input
          placeholder="Buscar por código ou nome…"
          className="max-w-xs"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        />
        <Button onClick={() => setEditing("new")}>
          <Plus className="h-4 w-4" /> Novo projeto
        </Button>
      </div>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      {editing && (
        <form action={submit} className="rounded-lg border bg-white p-4 shadow-sm">
          <div className="mb-3 flex items-center justify-between">
            <h3 className="font-medium">
              {editing === "new" ? "Novo projeto" : `Editar — ${editing.code}`}
            </h3>
            <button type="button" onClick={() => setEditing(null)} className="text-slate-400 hover:text-slate-700">
              <X className="h-4 w-4" />
            </button>
          </div>
          {editing !== "new" && <input type="hidden" name="id" value={editing.id} />}
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Código" required hint="UPPERCASE; sem espaço.">
              <Input name="code" required defaultValue={editing === "new" ? "" : editing.code} maxLength={30} />
            </Field>
            <Field label="Nome" required>
              <Input name="name" required defaultValue={editing === "new" ? "" : editing.name} maxLength={200} />
            </Field>
            <Field label="Módulo" required>
              <Select name="module" defaultValue={editing === "new" ? "PMO" : editing.module}>
                {Modules.map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </Select>
            </Field>
            <Field label="Status" required>
              <Select name="status" defaultValue={editing === "new" ? "ACTIVE" : editing.status}>
                {ProjectStatuses.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </Select>
            </Field>
            <Field label="Prioridade" hint="0 (mais alta) … 99 (mais baixa).">
              <Input name="priority" type="number" min={0} max={99} defaultValue={editing === "new" ? 0 : editing.priority} />
            </Field>
            <Field label="Owner" required>
              <Select name="ownerId" defaultValue={editing === "new" ? "" : editing.ownerId} required>
                <option value="">— selecione —</option>
                {users.map((u) => (
                  <option key={u.id} value={u.id}>
                    {u.name} ({u.email})
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Unidade">
              <Select name="unidadeId" defaultValue={editing === "new" ? "" : editing.unidadeId ?? ""}>
                <option value="">— nenhuma —</option>
                {unidades.map((u) => (
                  <option key={u.id} value={u.id}>{u.code} — {u.name}</option>
                ))}
              </Select>
            </Field>
            <Field label="Governança">
              <Select name="governancaId" defaultValue={editing === "new" ? "" : editing.governancaId ?? ""}>
                <option value="">— nenhuma —</option>
                {governancas.map((g) => (
                  <option key={g.id} value={g.id}>{g.code} — {g.name}</option>
                ))}
              </Select>
            </Field>
            <Field label="Início" required>
              <Input name="startDate" type="date" required defaultValue={editing === "new" ? "" : fmtDate(editing.startDate)} />
            </Field>
            <Field label="Fim" required>
              <Input name="endDate" type="date" required defaultValue={editing === "new" ? "" : fmtDate(editing.endDate)} />
            </Field>
            <Field label="Baseline">
              <Input name="baselineDate" type="date" defaultValue={editing === "new" ? "" : fmtDate(editing.baselineDate)} />
            </Field>
          </div>
          <Field label="Descrição">
            <Textarea name="description" defaultValue={editing === "new" ? "" : editing.description ?? ""} rows={3} />
          </Field>
          <div className="mt-4 flex justify-end gap-2">
            <Button type="button" variant="secondary" onClick={() => setEditing(null)} disabled={pending}>
              Cancelar
            </Button>
            <Button type="submit" disabled={pending}>
              {pending ? "Salvando…" : "Salvar"}
            </Button>
          </div>
        </form>
      )}

      <div className="overflow-hidden rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Código</th>
              <th className="px-3 py-2 text-left">Nome</th>
              <th className="px-3 py-2 text-left">Módulo</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-left">Unidade</th>
              <th className="px-3 py-2 text-left">Início</th>
              <th className="px-3 py-2 text-left">Fim</th>
              <th className="px-3 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.map((r) => (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-3 py-2 font-mono text-xs">{r.code}</td>
                <td className="px-3 py-2">{r.name}</td>
                <td className="px-3 py-2">
                  <span className="rounded bg-brand-50 px-2 py-0.5 text-xs text-brand-700">{r.module}</span>
                </td>
                <td className="px-3 py-2">
                  <span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{r.status}</span>
                </td>
                <td className="px-3 py-2 text-slate-600">{r.owner.name}</td>
                <td className="px-3 py-2 text-slate-600">{r.unidade?.code ?? "—"}</td>
                <td className="px-3 py-2">{fmtBR(r.startDate)}</td>
                <td className="px-3 py-2">{fmtBR(r.endDate)}</td>
                <td className="px-3 py-2 text-right">
                  <div className="flex justify-end gap-1">
                    <Button size="sm" variant="ghost" onClick={() => setEditing(r)}>
                      <Pencil className="h-3.5 w-3.5" />
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => handleDelete(r.id, r.code)}>
                      <Trash2 className="h-3.5 w-3.5 text-rose-600" />
                    </Button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={9} className="px-4 py-6 text-center text-slate-500">
                  Nenhum projeto.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
