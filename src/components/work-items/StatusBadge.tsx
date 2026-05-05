import { cn } from "@/lib/utils";

const TONES: Record<string, string> = {
  // status
  OPEN: "bg-amber-100 text-amber-800",
  IN_PROGRESS: "bg-blue-100 text-blue-800",
  RESOLVED: "bg-emerald-100 text-emerald-800",
  CLOSED: "bg-slate-100 text-slate-700",
  IDENTIFIED: "bg-amber-100 text-amber-800",
  ANALYZING: "bg-blue-100 text-blue-800",
  MITIGATING: "bg-violet-100 text-violet-800",
  ACCEPTED: "bg-slate-100 text-slate-700",
  DONE: "bg-emerald-100 text-emerald-800",
  CANCELLED: "bg-rose-100 text-rose-800",
  UNDER_REVIEW: "bg-blue-100 text-blue-800",
  APPROVED: "bg-emerald-100 text-emerald-800",
  IN_IMPLEMENTATION: "bg-violet-100 text-violet-800",
  IMPLEMENTED: "bg-emerald-100 text-emerald-800",
  REJECTED: "bg-rose-100 text-rose-800",
  // severity / priority
  LOW: "bg-slate-100 text-slate-700",
  MEDIUM: "bg-amber-100 text-amber-800",
  HIGH: "bg-orange-100 text-orange-800",
  CRITICAL: "bg-rose-100 text-rose-800",
};

const LABELS: Record<string, string> = {
  OPEN: "Aberto",
  IN_PROGRESS: "Em andamento",
  RESOLVED: "Resolvido",
  CLOSED: "Fechado",
  IDENTIFIED: "Identificado",
  ANALYZING: "Em análise",
  MITIGATING: "Mitigando",
  ACCEPTED: "Aceito",
  DONE: "Concluído",
  CANCELLED: "Cancelado",
  UNDER_REVIEW: "Em revisão",
  APPROVED: "Aprovado",
  IN_IMPLEMENTATION: "Em implementação",
  IMPLEMENTED: "Implementado",
  REJECTED: "Rejeitado",
  LOW: "Baixa",
  MEDIUM: "Média",
  HIGH: "Alta",
  CRITICAL: "Crítica",
};

export function StatusBadge({ value, size = "sm" }: { value: string; size?: "xs" | "sm" }) {
  const tone = TONES[value] ?? "bg-slate-100 text-slate-700";
  const label = LABELS[value] ?? value;
  return (
    <span
      className={cn(
        "inline-block rounded font-medium",
        size === "xs" ? "px-1.5 py-0.5 text-[10px]" : "px-2 py-0.5 text-xs",
        tone,
      )}
    >
      {label}
    </span>
  );
}
