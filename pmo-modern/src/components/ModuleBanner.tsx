import { getModuleConfig } from "@/lib/modules";
import { Sparkles } from "lucide-react";

export function ModuleBanner({ module }: { module: string }) {
  const cfg = getModuleConfig(module);
  if (cfg.hints.length === 0) return null;
  return (
    <div className={`rounded-md border border-transparent ${cfg.color} px-3 py-2 text-xs`}>
      <div className="flex items-center gap-1.5 font-medium">
        <Sparkles className="h-3 w-3" />
        Módulo {cfg.emoji} {cfg.label} — {cfg.description}
      </div>
      <ul className="mt-1 ml-4 list-disc space-y-0.5">
        {cfg.hints.map((h, i) => <li key={i}>{h}</li>)}
      </ul>
    </div>
  );
}
