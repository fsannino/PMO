"use client";

import { useState, useTransition } from "react";
import { Upload, FileSpreadsheet, FileCode2, Download, AlertTriangle, CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Field, Select } from "@/components/ui/Field";
import { previewImport, confirmImport } from "@/lib/actions/import";
import { ImportTemplates, type ImportTemplate } from "@/lib/enums";

type Source = "EXCEL" | "MSPROJECT_XML";
type Mode = "REPLACE" | "MERGE" | "ONLY_NEW";

type Preview = {
  template: ImportTemplate;
  source: Source;
  rowCount: number;
  sample: Record<string, unknown>[];
  warnings: string[];
  payload: string;
  filename: string;
};

const TEMPLATE_LABELS: Record<ImportTemplate, string> = {
  TK: "TK — Tarefas (cronograma)",
  IS: "IS — Issues",
  RK: "RK — Riscos",
  AC: "AC — Actions",
  CR: "CR — Change Requests",
  TC: "TC — Traceability",
};

const MODE_LABELS: Record<Mode, { label: string; desc: string }> = {
  MERGE: { label: "Mesclar", desc: "Atualiza existentes (por ID externo) e cria novos. Recomendado." },
  ONLY_NEW: { label: "Apenas novos", desc: "Cria apenas registros que ainda não existem." },
  REPLACE: { label: "Substituir tudo", desc: "⚠️ Apaga todos os registros do tipo no projeto antes de importar." },
};

export function ImportPanel({ projectId }: { projectId: string }) {
  const [source, setSource] = useState<Source>("EXCEL");
  const [hint, setHint] = useState<ImportTemplate | "">("");
  const [mode, setMode] = useState<Mode>("MERGE");
  const [file, setFile] = useState<File | null>(null);
  const [preview, setPreview] = useState<Preview | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [pending, start] = useTransition();
  const [dragOver, setDragOver] = useState(false);

  function reset() {
    setFile(null);
    setPreview(null);
    setError(null);
    setSuccess(null);
  }

  async function handleFile(f: File | null) {
    setError(null);
    setSuccess(null);
    setPreview(null);
    if (!f) return;
    setFile(f);

    // detecta source automaticamente pela extensão
    const ext = f.name.split(".").pop()?.toLowerCase() ?? "";
    const isXml = ext === "xml";
    setSource(isXml ? "MSPROJECT_XML" : "EXCEL");

    // ler como base64
    const buf = await f.arrayBuffer();
    const base64 = bufferToBase64(buf);

    start(async () => {
      const res = await previewImport({
        projectId,
        filename: f.name,
        source: isXml ? "MSPROJECT_XML" : "EXCEL",
        templateHint: hint || undefined,
        base64,
      });
      if (!res.ok) {
        setError(res.error);
        return;
      }
      setPreview({
        template: res.template,
        source: res.source,
        rowCount: res.rowCount,
        sample: res.sample as Record<string, unknown>[],
        warnings: res.warnings,
        payload: res.payload,
        filename: f.name,
      });
    });
  }

  function handleDrop(e: React.DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  }

  function handleConfirm() {
    if (!preview) return;
    setError(null);
    start(async () => {
      const res = await confirmImport({
        projectId,
        source: preview.source,
        mode,
        payload: preview.payload,
      });
      if (!res.ok) {
        setError(res.error);
      } else {
        setSuccess(`Importação concluída: ${res.recordCount} registro(s) processado(s).`);
        setPreview(null);
        setFile(null);
      }
    });
  }

  return (
    <div className="space-y-6">
      {/* Templates baixáveis */}
      <section className="rounded-lg border bg-white p-4">
        <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
          Templates Excel
        </h3>
        <div className="flex flex-wrap gap-2">
          {ImportTemplates.map((t) => (
            <a
              key={t}
              href={`/api/import/template/${t}`}
              className="inline-flex items-center gap-1.5 rounded border border-slate-300 bg-white px-3 py-1.5 text-xs hover:bg-slate-50"
              download
            >
              <Download className="h-3.5 w-3.5" />
              {t}.xlsx
            </a>
          ))}
        </div>
        <p className="mt-2 text-xs text-slate-500">
          Use os mesmos cabeçalhos do template; aceitamos PT-BR e EN.
        </p>
      </section>

      {/* Upload + opções */}
      <section className="rounded-lg border bg-white p-4">
        <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
          Carregar arquivo
        </h3>

        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Tipo de template (opcional, auto-detectado)">
            <Select value={hint} onChange={(e) => setHint(e.target.value as ImportTemplate | "")}>
              <option value="">— auto-detectar —</option>
              {ImportTemplates.map((t) => (
                <option key={t} value={t}>{TEMPLATE_LABELS[t]}</option>
              ))}
            </Select>
          </Field>
          <Field label="Modo de importação" hint={MODE_LABELS[mode].desc}>
            <Select value={mode} onChange={(e) => setMode(e.target.value as Mode)}>
              {(Object.keys(MODE_LABELS) as Mode[]).map((m) => (
                <option key={m} value={m}>{MODE_LABELS[m].label}</option>
              ))}
            </Select>
          </Field>
        </div>

        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
          className={`mt-4 flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed p-8 text-center transition ${
            dragOver ? "border-brand-500 bg-brand-50" : "border-slate-300 bg-slate-50"
          }`}
        >
          <Upload className="h-8 w-8 text-slate-400" />
          <p className="text-sm text-slate-600">
            Arraste um arquivo <strong>.xlsx</strong> ou <strong>.xml</strong> (MS Project) aqui,
            <br />ou clique no botão abaixo.
          </p>
          <label className="mt-2 inline-block cursor-pointer rounded-md bg-brand-600 px-4 py-2 text-sm font-medium text-white hover:bg-brand-700">
            Selecionar arquivo
            <input
              type="file"
              accept=".xlsx,.xls,.xml"
              className="hidden"
              onChange={(e) => handleFile(e.target.files?.[0] ?? null)}
            />
          </label>
          {file && <p className="text-xs text-slate-500">{file.name} · {(file.size / 1024).toFixed(1)} KB</p>}
        </div>
      </section>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</div>
      )}
      {success && (
        <div className="flex items-start gap-2 rounded border border-emerald-200 bg-emerald-50 px-3 py-2 text-sm text-emerald-700">
          <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0" />
          <div>
            {success}
            <div className="mt-1">
              <Button size="sm" variant="secondary" onClick={reset}>Nova importação</Button>
            </div>
          </div>
        </div>
      )}
      {pending && !preview && <div className="text-sm text-slate-500">Processando…</div>}

      {/* Preview + confirmar */}
      {preview && (
        <section className="rounded-lg border bg-white p-4">
          <div className="mb-3 flex flex-wrap items-center gap-3">
            {preview.source === "EXCEL" ? (
              <FileSpreadsheet className="h-5 w-5 text-emerald-600" />
            ) : (
              <FileCode2 className="h-5 w-5 text-violet-600" />
            )}
            <h3 className="font-semibold">Preview</h3>
            <span className="rounded bg-brand-50 px-2 py-0.5 text-xs font-medium text-brand-700">
              {TEMPLATE_LABELS[preview.template]}
            </span>
            <span className="text-sm text-slate-600">
              {preview.rowCount} linha(s) — exibindo {Math.min(50, preview.rowCount)}
            </span>
            <span className="ml-auto text-xs text-slate-500">{preview.filename}</span>
          </div>

          {preview.warnings.length > 0 && (
            <div className="mb-3 rounded border border-amber-200 bg-amber-50 p-2 text-xs text-amber-800">
              <div className="mb-1 flex items-center gap-1 font-medium">
                <AlertTriangle className="h-3.5 w-3.5" /> {preview.warnings.length} aviso(s):
              </div>
              <ul className="ml-4 list-disc space-y-0.5">
                {preview.warnings.slice(0, 10).map((w, i) => <li key={i}>{w}</li>)}
                {preview.warnings.length > 10 && (
                  <li className="text-amber-700">… e mais {preview.warnings.length - 10}.</li>
                )}
              </ul>
            </div>
          )}

          {preview.sample.length > 0 ? (
            <div className="overflow-x-auto rounded border">
              <table className="w-full text-xs">
                <thead className="bg-slate-50 text-slate-500">
                  <tr>
                    {Object.keys(preview.sample[0]).map((k) => (
                      <th key={k} className="whitespace-nowrap px-2 py-1 text-left font-medium">{k}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {preview.sample.map((row, i) => (
                    <tr key={i} className="hover:bg-slate-50">
                      {Object.entries(row).map(([k, v]) => (
                        <td key={k} className="whitespace-nowrap px-2 py-1">
                          {formatCell(v)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-sm text-slate-500">Nenhuma linha válida.</p>
          )}

          <div className="mt-4 flex justify-end gap-2 border-t pt-3">
            <Button variant="secondary" onClick={reset} disabled={pending}>Cancelar</Button>
            <Button onClick={handleConfirm} disabled={pending || preview.rowCount === 0}>
              {pending ? "Importando…" : `Confirmar ${MODE_LABELS[mode].label.toLowerCase()}`}
            </Button>
          </div>
        </section>
      )}
    </div>
  );
}

function formatCell(v: unknown): string {
  if (v == null || v === "") return "—";
  if (Array.isArray(v)) return v.join(", ");
  if (v instanceof Date) return v.toLocaleDateString("pt-BR");
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T/.test(v)) return new Date(v).toLocaleDateString("pt-BR");
  if (typeof v === "boolean") return v ? "Sim" : "Não";
  return String(v);
}

function bufferToBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}
