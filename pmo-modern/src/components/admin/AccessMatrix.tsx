"use client";

import { useMemo, useState, useTransition } from "react";
import { useRouter } from "next/navigation";
import { Button } from "@/components/ui/Button";
import { Field, Select } from "@/components/ui/Field";
import { Save } from "lucide-react";
import { saveAccessMatrix } from "@/lib/actions/access";
import { Modules, type Module } from "@/lib/enums";

type ProjectOpt = { id: string; code: string; name: string; module: string };
type UserOpt = { id: string; name: string; email: string; role: string; active: boolean };
type AccessRow = { userId: string; projectId: string; module: string; canRead: boolean; canWrite: boolean; canAdmin: boolean };

type RowState = { userId: string; canRead: boolean; canWrite: boolean; canAdmin: boolean };

export function AccessMatrix({
  projects,
  users,
  accesses,
}: {
  projects: ProjectOpt[];
  users: UserOpt[];
  accesses: AccessRow[];
}) {
  const router = useRouter();
  const [projectId, setProjectId] = useState(projects[0]?.id ?? "");
  const [module, setModule] = useState<Module>(
    (projects[0]?.module as Module) ?? "PMO",
  );
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();

  // estado das linhas baseado nos acessos atuais
  const initialRows: RowState[] = useMemo(() => {
    return users
      .filter((u) => u.active)
      .map((u) => {
        const a = accesses.find(
          (x) => x.userId === u.id && x.projectId === projectId && x.module === module,
        );
        return {
          userId: u.id,
          canRead: a?.canRead ?? false,
          canWrite: a?.canWrite ?? false,
          canAdmin: a?.canAdmin ?? false,
        };
      });
  }, [users, accesses, projectId, module]);
  const [rows, setRows] = useState<RowState[]>(initialRows);

  // ao mudar projeto/módulo, recarrega rows
  useMemo(() => setRows(initialRows), [initialRows]);

  function update(userId: string, key: keyof Omit<RowState, "userId">, value: boolean) {
    setRows((rs) =>
      rs.map((r) => {
        if (r.userId !== userId) return r;
        const next = { ...r, [key]: value };
        // canWrite implica canRead; canAdmin implica canRead+canWrite
        if (key === "canWrite" && value) next.canRead = true;
        if (key === "canAdmin" && value) {
          next.canRead = true;
          next.canWrite = true;
        }
        if (key === "canRead" && !value) {
          next.canWrite = false;
          next.canAdmin = false;
        }
        return next;
      }),
    );
  }

  function save() {
    setError(null);
    start(async () => {
      const res = await saveAccessMatrix({ projectId, module, rows });
      if (res.ok) router.refresh();
      else setError(res.error);
    });
  }

  return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Projeto">
          <Select value={projectId} onChange={(e) => setProjectId(e.target.value)}>
            {projects.map((p) => (
              <option key={p.id} value={p.id}>
                {p.code} — {p.name}
              </option>
            ))}
          </Select>
        </Field>
        <Field label="Módulo">
          <Select value={module} onChange={(e) => setModule(e.target.value as Module)}>
            {Modules.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </Select>
        </Field>
      </div>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      <div className="overflow-hidden rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Usuário</th>
              <th className="px-3 py-2 text-left">Papel</th>
              <th className="px-3 py-2 text-center">Ler</th>
              <th className="px-3 py-2 text-center">Escrever</th>
              <th className="px-3 py-2 text-center">Admin</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {users.filter((u) => u.active).map((u) => {
              const r = rows.find((x) => x.userId === u.id)!;
              const isAdmin = u.role === "ADMIN";
              return (
                <tr key={u.id} className={isAdmin ? "bg-amber-50/40" : ""}>
                  <td className="px-3 py-2">
                    <div className="font-medium">{u.name}</div>
                    <div className="text-xs text-slate-500">{u.email}</div>
                  </td>
                  <td className="px-3 py-2">
                    <span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{u.role}</span>
                    {isAdmin && (
                      <div className="text-xs text-amber-700 mt-1">bypass total</div>
                    )}
                  </td>
                  <td className="px-3 py-2 text-center">
                    <input
                      type="checkbox"
                      disabled={isAdmin}
                      checked={isAdmin ? true : r.canRead}
                      onChange={(e) => update(u.id, "canRead", e.target.checked)}
                    />
                  </td>
                  <td className="px-3 py-2 text-center">
                    <input
                      type="checkbox"
                      disabled={isAdmin}
                      checked={isAdmin ? true : r.canWrite}
                      onChange={(e) => update(u.id, "canWrite", e.target.checked)}
                    />
                  </td>
                  <td className="px-3 py-2 text-center">
                    <input
                      type="checkbox"
                      disabled={isAdmin}
                      checked={isAdmin ? true : r.canAdmin}
                      onChange={(e) => update(u.id, "canAdmin", e.target.checked)}
                    />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="flex justify-end">
        <Button onClick={save} disabled={pending}>
          <Save className="h-4 w-4" /> {pending ? "Salvando…" : "Salvar matriz"}
        </Button>
      </div>
    </div>
  );
}
