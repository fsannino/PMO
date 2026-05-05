"use client";

import { useState, useTransition } from "react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select } from "@/components/ui/Field";
import { Pencil, Plus, X, KeyRound, Power } from "lucide-react";
import { upsertUser, toggleUserActive, resetPassword } from "@/lib/actions/users";
import type { Role } from "@/lib/enums";
import { Roles } from "@/lib/enums";

type Lookup = { id: string; code: string; name: string };
export type UserRow = {
  id: string;
  email: string;
  name: string;
  role: string;
  active: boolean;
  unidadeId: string | null;
  areaId: string | null;
  unidade?: { code: string } | null;
  area?: { code: string } | null;
};

export function UserCRUD({
  rows,
  unidades,
  areas,
}: {
  rows: UserRow[];
  unidades: Lookup[];
  areas: Lookup[];
}) {
  const [editing, setEditing] = useState<UserRow | "new" | null>(null);
  const [resetTarget, setResetTarget] = useState<UserRow | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();
  const [filter, setFilter] = useState("");

  const filtered = rows.filter(
    (r) =>
      r.email.toLowerCase().includes(filter.toLowerCase()) ||
      r.name.toLowerCase().includes(filter.toLowerCase()),
  );

  function submitUpsert(formData: FormData) {
    setError(null);
    start(async () => {
      const res = await upsertUser(formData);
      if (res.ok) setEditing(null);
      else setError(res.error);
    });
  }

  function toggle(id: string) {
    setError(null);
    start(async () => {
      const res = await toggleUserActive(id);
      if (!res.ok) setError(res.error);
    });
  }

  function submitReset(formData: FormData) {
    if (!resetTarget) return;
    setError(null);
    const newPwd = String(formData.get("password") ?? "");
    start(async () => {
      const res = await resetPassword(resetTarget.id, newPwd);
      if (res.ok) setResetTarget(null);
      else setError(res.error);
    });
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2">
        <Input
          placeholder="Buscar por nome ou e-mail…"
          className="max-w-xs"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        />
        <Button onClick={() => setEditing("new")}>
          <Plus className="h-4 w-4" /> Novo usuário
        </Button>
      </div>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      {editing && (
        <form action={submitUpsert} className="rounded-lg border bg-white p-4 shadow-sm">
          <div className="mb-3 flex items-center justify-between">
            <h3 className="font-medium">
              {editing === "new" ? "Novo usuário" : `Editar — ${editing.email}`}
            </h3>
            <button type="button" onClick={() => setEditing(null)} className="text-slate-400 hover:text-slate-700">
              <X className="h-4 w-4" />
            </button>
          </div>
          {editing !== "new" && <input type="hidden" name="id" value={editing.id} />}
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Nome" required>
              <Input name="name" required defaultValue={editing === "new" ? "" : editing.name} />
            </Field>
            <Field label="E-mail" required>
              <Input
                name="email"
                type="email"
                required
                defaultValue={editing === "new" ? "" : editing.email}
              />
            </Field>
            <Field label="Papel" required>
              <Select name="role" defaultValue={editing === "new" ? "MEMBER" : editing.role}>
                {Roles.map((r) => (
                  <option key={r} value={r}>
                    {r}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label={editing === "new" ? "Senha (mín 6)" : "Nova senha (deixe vazio p/ manter)"} required={editing === "new"}>
              <Input
                name="password"
                type="password"
                minLength={6}
                required={editing === "new"}
                placeholder={editing === "new" ? "" : "••••••"}
              />
            </Field>
            <Field label="Unidade">
              <Select name="unidadeId" defaultValue={editing === "new" ? "" : editing.unidadeId ?? ""}>
                <option value="">— nenhuma —</option>
                {unidades.map((u) => (
                  <option key={u.id} value={u.id}>
                    {u.code} — {u.name}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Área">
              <Select name="areaId" defaultValue={editing === "new" ? "" : editing.areaId ?? ""}>
                <option value="">— nenhuma —</option>
                {areas.map((a) => (
                  <option key={a.id} value={a.id}>
                    {a.code} — {a.name}
                  </option>
                ))}
              </Select>
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

      {resetTarget && (
        <form action={submitReset} className="rounded-lg border border-amber-200 bg-amber-50 p-4">
          <div className="mb-2 flex items-center justify-between">
            <h3 className="font-medium text-amber-900">
              Resetar senha — {resetTarget.email}
            </h3>
            <button type="button" onClick={() => setResetTarget(null)} className="text-amber-600 hover:text-amber-900">
              <X className="h-4 w-4" />
            </button>
          </div>
          <Field label="Nova senha (mín 6)" required>
            <Input name="password" type="password" minLength={6} required />
          </Field>
          <div className="mt-3 flex justify-end gap-2">
            <Button type="button" variant="secondary" onClick={() => setResetTarget(null)} disabled={pending}>
              Cancelar
            </Button>
            <Button type="submit" disabled={pending}>
              {pending ? "Salvando…" : "Resetar"}
            </Button>
          </div>
        </form>
      )}

      <div className="overflow-hidden rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-4 py-2 text-left">Nome</th>
              <th className="px-4 py-2 text-left">E-mail</th>
              <th className="px-4 py-2 text-left">Papel</th>
              <th className="px-4 py-2 text-left">Unidade / Área</th>
              <th className="px-4 py-2 text-left">Ativo</th>
              <th className="px-4 py-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.map((r) => (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-4 py-2">{r.name}</td>
                <td className="px-4 py-2 font-mono text-xs">{r.email}</td>
                <td className="px-4 py-2">
                  <span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{r.role}</span>
                </td>
                <td className="px-4 py-2 text-xs text-slate-600">
                  {r.unidade?.code ?? "—"} / {r.area?.code ?? "—"}
                </td>
                <td className="px-4 py-2">
                  <span className={r.active ? "text-emerald-700" : "text-rose-600"}>
                    {r.active ? "Sim" : "Não"}
                  </span>
                </td>
                <td className="px-4 py-2 text-right">
                  <div className="flex justify-end gap-1">
                    <Button size="sm" variant="ghost" onClick={() => setEditing(r)} title="Editar">
                      <Pencil className="h-3.5 w-3.5" />
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => setResetTarget(r)} title="Resetar senha">
                      <KeyRound className="h-3.5 w-3.5 text-amber-700" />
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => toggle(r.id)} title={r.active ? "Desativar" : "Ativar"}>
                      <Power className={r.active ? "h-3.5 w-3.5 text-rose-600" : "h-3.5 w-3.5 text-emerald-600"} />
                    </Button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={6} className="px-4 py-6 text-center text-slate-500">
                  Nenhum usuário.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
