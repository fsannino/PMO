"use client";

import { useMemo, useState, useTransition } from "react";
import { useRouter } from "next/navigation";
import { Button } from "@/components/ui/Button";
import { Field, Input, Select } from "@/components/ui/Field";
import { Save } from "lucide-react";
import { saveConfigFechamento } from "@/lib/actions/config-fechamento";

const DAYS = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];

type Project = { id: string; code: string; name: string };
export type Config = {
  projectId: string | null;
  dayOfWeek: number;
  enabled: boolean;
  startTime: string;
  endTime: string;
};
type DayState = { dayOfWeek: number; enabled: boolean; startTime: string; endTime: string };

const DEFAULTS: DayState[] = Array.from({ length: 7 }, (_, i) => ({
  dayOfWeek: i,
  enabled: i >= 1 && i <= 5,
  startTime: "08:00",
  endTime: "18:00",
}));

export function ConfigFechamentoEditor({
  projects,
  configs,
}: {
  projects: Project[];
  configs: Config[];
}) {
  const router = useRouter();
  const [scope, setScope] = useState<string>("global"); // "global" | projectId
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();

  const days: DayState[] = useMemo(() => {
    const projectId = scope === "global" ? null : scope;
    const result = DEFAULTS.map((d) => {
      const found = configs.find(
        (c) => (c.projectId ?? "global") === (projectId ?? "global") && c.dayOfWeek === d.dayOfWeek,
      );
      return found
        ? { dayOfWeek: d.dayOfWeek, enabled: found.enabled, startTime: found.startTime, endTime: found.endTime }
        : d;
    });
    return result;
  }, [scope, configs]);

  const [state, setState] = useState<DayState[]>(days);

  // sincroniza quando scope muda
  useMemo(() => setState(days), [days]);

  function update(idx: number, patch: Partial<DayState>) {
    setState((s) => s.map((d, i) => (i === idx ? { ...d, ...patch } : d)));
  }

  function save() {
    setError(null);
    const projectId = scope === "global" ? null : scope;
    start(async () => {
      const res = await saveConfigFechamento({ projectId, days: state });
      if (res.ok) router.refresh();
      else setError(res.error);
    });
  }

  return (
    <div className="space-y-4">
      <Field label="Escopo da configuração" hint="A configuração GLOBAL é o fallback quando o projeto não tem configuração própria.">
        <Select value={scope} onChange={(e) => setScope(e.target.value)}>
          <option value="global">🌐 Global (fallback)</option>
          {projects.map((p) => (
            <option key={p.id} value={p.id}>
              {p.code} — {p.name}
            </option>
          ))}
        </Select>
      </Field>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      <div className="overflow-hidden rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Dia</th>
              <th className="px-3 py-2 text-center">Aberto</th>
              <th className="px-3 py-2 text-left">Início</th>
              <th className="px-3 py-2 text-left">Fim</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {state.map((d, idx) => (
              <tr key={d.dayOfWeek}>
                <td className="px-3 py-2 font-medium">{DAYS[d.dayOfWeek]}</td>
                <td className="px-3 py-2 text-center">
                  <input
                    type="checkbox"
                    checked={d.enabled}
                    onChange={(e) => update(idx, { enabled: e.target.checked })}
                  />
                </td>
                <td className="px-3 py-2">
                  <Input
                    type="time"
                    value={d.startTime}
                    onChange={(e) => update(idx, { startTime: e.target.value })}
                    disabled={!d.enabled}
                    className="max-w-[8rem]"
                  />
                </td>
                <td className="px-3 py-2">
                  <Input
                    type="time"
                    value={d.endTime}
                    onChange={(e) => update(idx, { endTime: e.target.value })}
                    disabled={!d.enabled}
                    className="max-w-[8rem]"
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex justify-end">
        <Button onClick={save} disabled={pending}>
          <Save className="h-4 w-4" /> {pending ? "Salvando…" : "Salvar"}
        </Button>
      </div>
    </div>
  );
}
