import { Sparkles } from "lucide-react";

export function ComingSoon({
  title,
  description,
  session,
}: {
  title: string;
  description: string;
  session: number;
}) {
  return (
    <div className="space-y-6">
      <header>
        <h1 className="text-2xl font-semibold">{title}</h1>
      </header>
      <div className="rounded-lg border bg-white p-8 text-center shadow-sm">
        <Sparkles className="mx-auto h-12 w-12 text-brand-500" />
        <h2 className="mt-4 text-lg font-semibold">Em breve</h2>
        <p className="mx-auto mt-2 max-w-md text-sm text-slate-600">{description}</p>
        <p className="mt-4 text-xs text-slate-500">
          Implementação prevista para a <strong>Sessão {session}</strong> do plano
          (<code className="rounded bg-slate-100 px-1">pmo-modern/PLAN.md</code>).
        </p>
      </div>
    </div>
  );
}
