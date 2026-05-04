"use client";

import { useRouter, useSearchParams } from "next/navigation";
import { useTransition } from "react";
import { Button } from "@/components/ui/Button";
import { cn } from "@/lib/utils";

export function SynopticToolbar({ projectId }: { projectId: string }) {
  const router = useRouter();
  const params = useSearchParams();
  const granularity = (params.get("g") ?? "week") as "week" | "month";
  const [pending, start] = useTransition();

  function setG(g: "week" | "month") {
    const usp = new URLSearchParams(params.toString());
    usp.set("g", g);
    start(() => router.push(`/projects/${projectId}/synoptic?${usp.toString()}`));
  }

  return (
    <div className="flex items-center gap-2">
      <span className="text-xs text-slate-500">Granularidade:</span>
      {(["week", "month"] as const).map((g) => (
        <Button
          key={g}
          size="sm"
          variant={granularity === g ? "primary" : "secondary"}
          onClick={() => setG(g)}
          disabled={pending}
          className={cn(granularity === g && "shadow-sm")}
        >
          {g === "week" ? "Semana" : "Mês"}
        </Button>
      ))}
    </div>
  );
}
