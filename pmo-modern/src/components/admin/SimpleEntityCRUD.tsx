"use client";

import { useState, useTransition } from "react";
import { Button } from "@/components/ui/Button";
import { Field, Input } from "@/components/ui/Field";
import { Pencil, Trash2, Plus, X } from "lucide-react";
import {
  upsertSimpleEntity,
  deleteSimpleEntity,
  type SimpleEntityKind,
} from "@/lib/actions/simple-entity";

type Row = { id: string; code: string; name: string; active: boolean };

export function SimpleEntityCRUD({
  kind,
  rows,
  labelSingular,
}: {
  kind: SimpleEntityKind;
  rows: Row[];
  labelSingular: string;
}) {
  const [editing, setEditing] = useState<Row | "new" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();
  const [filter, setFilter] = useState("");

  const filtered = rows.filter(
    (r) => r.code.toLowerCase().includes(filter.toLowerCase()) || r.name.toLowerCase().includes(filter.toLowerCase()),
  );

  function handleSubmit(formData: FormData) {
    setError(null);
    start(async () => {
      const res = await upsertSimpleEntity(kind, formData);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }

  function handleDelete(id: string, code: string) {
    if (!confirm(`Excluir "${code}"? Esta ação não pode ser desfeita.`)) return;
    setError(null);
    start(async () => {
      const res = await deleteSimpleEntity(kind, id);
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
          <Plus className="h-4 w-4" /> Novo {labelSingular}
        </Button>
      </div>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      {editing && (
        <form
          action={handleSubmit}
          className="rounded-lg border bg-white p-4 shadow-sm"
        >
          <div className="mb-3 flex items-center justify-between">
            <h3 className="font-medium">
              {editing === "new" ? `Novo ${labelSingular}` : `Editar ${labelSingular}`}
            </h3>
            <button type="button" onClick={() => setEditing(null)} className="text-slate-400 hover:text-slate-700">
              <X className="h-4 w-4" />
            </button>
          </div>
          {editing !== "new" && <input type="hidden" name="id" value={editing.id} />}
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Código" required hint="Letras maiúsculas, números, _ e -. Convertido para UPPERCASE.">
              <Input
                name="code"
                required
                defaultValue={editing === "new" ? "" : editing.code}
                pattern="[A-Za-z0-9_\-]+"
                maxLength={20}
              />
            </Field>
            <Field label="Nome" required>
              <Input name="name" required defaultValue={editing === "new" ? "" : editing.name} maxLength={120} />
            </Field>
          </div>
          <label className="mt-3 flex items-center gap-2 text-sm">
            <input
              type="checkbox"
              name="active"
              defaultChecked={editing === "new" ? true : editing.active}
            />
            Ativo
          </label>
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
              <th className="px-4 py-2 text-left">Código</th>
              <th className="px-4 py-2 text-left">Nome</th>
              <th className="px-4 py-2 text-left">Ativo</th>
              <th className="px-4 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.length === 0 && (
              <tr>
                <td colSpan={4} className="px-4 py-6 text-center text-slate-500">
                  Nenhum registro.
                </td>
              </tr>
            )}
            {filtered.map((r) => (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-4 py-2 font-mono text-xs">{r.code}</td>
                <td className="px-4 py-2">{r.name}</td>
                <td className="px-4 py-2">
                  <span className={r.active ? "text-emerald-700" : "text-slate-400"}>
                    {r.active ? "Sim" : "Não"}
                  </span>
                </td>
                <td className="px-4 py-2 text-right">
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
          </tbody>
        </table>
      </div>
    </div>
  );
}
