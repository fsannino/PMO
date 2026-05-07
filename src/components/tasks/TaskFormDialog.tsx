"use client";

import { useState, useTransition } from "react";
import { X } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select, Textarea } from "@/components/ui/Field";
import { upsertTask } from "@/lib/actions/tasks";
import type { TaskRow, Lookup } from "./TaskTreeView";

export type TaskFormDefaults = {
  id?: string;
  parentId: string | null;
  wbs: string | null;
  externalId: string | null;
  name: string;
  description: string | null;
  startDate: string;
  endDate: string;
  baselineStart: string | null;
  baselineEnd: string | null;
  percentDone: number;
  isMilestone: boolean;
  isSummary: boolean;
  assigneeId: string | null;
  areaId: string | null;
  frenteId: string | null;
  equipeId: string | null;
};

export function TaskFormDialog({
  projectId,
  tasks,
  users,
  areas,
  frentes,
  equipes,
  defaults,
  onClose,
}: {
  projectId: string;
  tasks: TaskRow[];
  users: Lookup[];
  areas: Lookup[];
  frentes: Lookup[];
  equipes: Lookup[];
  defaults: TaskFormDefaults | null;
  onClose: () => void;
}) {
  const isNew = !defaults;
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();
  const [isSummary, setIsSummary] = useState(defaults?.isSummary ?? false);
  const [isMilestone, setIsMilestone] = useState(defaults?.isMilestone ?? false);

  function handleSubmit(formData: FormData) {
    setError(null);
    const data = {
      id: defaults?.id,
      projectId,
      parentId: (formData.get("parentId") as string) || null,
      wbs: (formData.get("wbs") as string) || null,
      externalId: (formData.get("externalId") as string) || null,
      name: formData.get("name"),
      description: (formData.get("description") as string) || null,
      startDate: formData.get("startDate"),
      endDate: formData.get("endDate"),
      baselineStart: (formData.get("baselineStart") as string) || null,
      baselineEnd: (formData.get("baselineEnd") as string) || null,
      percentDone: Number(formData.get("percentDone") ?? 0),
      isMilestone: formData.get("isMilestone") === "on",
      isSummary: formData.get("isSummary") === "on",
      assigneeId: (formData.get("assigneeId") as string) || null,
      areaId: (formData.get("areaId") as string) || null,
      frenteId: (formData.get("frenteId") as string) || null,
      equipeId: (formData.get("equipeId") as string) || null,
    };
    start(async () => {
      const res = await upsertTask(data);
      if (res.ok) onClose();
      else setError(res.error);
    });
  }

  // possíveis pais: tarefas-resumo (e a si mesmo é filtrado no submit)
  const summaryOptions = tasks.filter((t) => !t.deletedAt && (t.isSummary || t.id === defaults?.parentId));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
      <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-lg bg-white shadow-xl">
        <form action={handleSubmit}>
          <div className="flex items-center justify-between border-b px-5 py-3">
            <h2 className="font-semibold">{isNew ? "Nova tarefa" : `Editar — ${defaults?.name}`}</h2>
            <button type="button" onClick={onClose} className="text-slate-400 hover:text-slate-700">
              <X className="h-4 w-4" />
            </button>
          </div>

          <div className="space-y-3 p-5">
            {error && (
              <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>
            )}

            <Field label="Nome" required>
              <Input name="name" required defaultValue={defaults?.name ?? ""} maxLength={200} />
            </Field>

            <div className="grid gap-3 sm:grid-cols-2">
              <Field label="WBS / EAP">
                <Input name="wbs" defaultValue={defaults?.wbs ?? ""} maxLength={50} placeholder="1.2.3" />
              </Field>
              <Field label="ID externo (MS Project / Excel)">
                <Input name="externalId" defaultValue={defaults?.externalId ?? ""} maxLength={50} />
              </Field>
              <Field label="Tarefa pai (resumo)">
                <Select name="parentId" defaultValue={defaults?.parentId ?? ""}>
                  <option value="">— raiz —</option>
                  {summaryOptions.map((t) => (
                    <option key={t.id} value={t.id}>
                      {t.wbs ? `${t.wbs} ` : ""}{t.name}
                    </option>
                  ))}
                </Select>
              </Field>
              <div className="flex items-end gap-4 pb-1">
                <label className="flex items-center gap-2 text-sm">
                  <input
                    type="checkbox"
                    name="isSummary"
                    checked={isSummary}
                    onChange={(e) => {
                      setIsSummary(e.target.checked);
                      if (e.target.checked) setIsMilestone(false);
                    }}
                  />
                  Resumo
                </label>
                <label className="flex items-center gap-2 text-sm">
                  <input
                    type="checkbox"
                    name="isMilestone"
                    checked={isMilestone}
                    onChange={(e) => {
                      setIsMilestone(e.target.checked);
                      if (e.target.checked) setIsSummary(false);
                    }}
                  />
                  Marco
                </label>
              </div>
            </div>

            <div className="grid gap-3 sm:grid-cols-2">
              <Field label="Início" required>
                <Input name="startDate" type="date" required defaultValue={defaults?.startDate ?? ""} />
              </Field>
              <Field label="Fim" required>
                <Input name="endDate" type="date" required defaultValue={defaults?.endDate ?? ""} />
              </Field>
              <Field label="Baseline início">
                <Input name="baselineStart" type="date" defaultValue={defaults?.baselineStart ?? ""} />
              </Field>
              <Field label="Baseline fim">
                <Input name="baselineEnd" type="date" defaultValue={defaults?.baselineEnd ?? ""} />
              </Field>
            </div>

            <div className="grid gap-3 sm:grid-cols-2">
              <Field label="% Realizado">
                <Input
                  name="percentDone"
                  type="number"
                  min={0}
                  max={100}
                  defaultValue={defaults?.percentDone ?? 0}
                  disabled={isSummary}
                />
              </Field>
              <Field label="Responsável">
                <Select name="assigneeId" defaultValue={defaults?.assigneeId ?? ""}>
                  <option value="">— sem responsável —</option>
                  {users.map((u) => (
                    <option key={u.id} value={u.id}>{u.name}</option>
                  ))}
                </Select>
              </Field>
              <Field label="Equipe">
                <Select name="equipeId" defaultValue={defaults?.equipeId ?? ""}>
                  <option value="">—</option>
                  {equipes.map((e) => (
                    <option key={e.id} value={e.id}>{e.code} — {e.name}</option>
                  ))}
                </Select>
              </Field>
              <Field label="Área">
                <Select name="areaId" defaultValue={defaults?.areaId ?? ""}>
                  <option value="">—</option>
                  {areas.map((a) => (
                    <option key={a.id} value={a.id}>{a.code} — {a.name}</option>
                  ))}
                </Select>
              </Field>
              <Field label="Frente">
                <Select name="frenteId" defaultValue={defaults?.frenteId ?? ""}>
                  <option value="">—</option>
                  {frentes.map((f) => (
                    <option key={f.id} value={f.id}>{f.code} — {f.name}</option>
                  ))}
                </Select>
              </Field>
            </div>

            <Field label="Descrição">
              <Textarea name="description" rows={3} defaultValue={defaults?.description ?? ""} maxLength={2000} />
            </Field>
          </div>

          <div className="flex justify-end gap-2 border-t bg-slate-50 px-5 py-3">
            <Button type="button" variant="secondary" onClick={onClose} disabled={pending}>Cancelar</Button>
            <Button type="submit" disabled={pending}>{pending ? "Salvando…" : "Salvar"}</Button>
          </div>
        </form>
      </div>
    </div>
  );
}
