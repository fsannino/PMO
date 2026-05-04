import { prisma } from "@/lib/db";
import { notFound } from "next/navigation";
import Link from "next/link";

export default async function ProjectDetailPage({
  params,
}: {
  params: { id: string };
}) {
  const project = await prisma.project.findUnique({
    where: { id: params.id },
    include: {
      unidade: true,
      governanca: true,
      owner: { select: { name: true, email: true } },
      _count: {
        select: {
          tasks: true,
          issues: true,
          risks: true,
          actions: true,
          changeRequests: true,
        },
      },
    },
  });
  if (!project) notFound();

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-2 text-sm text-slate-500">
        <span className="rounded bg-slate-100 px-2 py-0.5 text-xs font-medium">
          {project.status}
        </span>
        {project.description && <span>· {project.description}</span>}
      </div>

      <section className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <Stat label="Tarefas" value={project._count.tasks} href={`/projects/${project.id}/tasks`} />
        <Stat label="Issues" value={project._count.issues} />
        <Stat label="Riscos" value={project._count.risks} />
        <Stat label="Change Requests" value={project._count.changeRequests} />
      </section>

      <section className="grid gap-4 md:grid-cols-2">
        <div className="rounded-lg border bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Detalhes
          </h3>
          <dl className="grid grid-cols-2 gap-y-2 text-sm">
            <dt className="text-slate-500">Unidade</dt>
            <dd>{project.unidade?.name ?? "—"}</dd>
            <dt className="text-slate-500">Governança</dt>
            <dd>{project.governanca?.name ?? "—"}</dd>
            <dt className="text-slate-500">Owner</dt>
            <dd>{project.owner.name}</dd>
            <dt className="text-slate-500">Início</dt>
            <dd>{project.startDate.toLocaleDateString("pt-BR")}</dd>
            <dt className="text-slate-500">Fim</dt>
            <dd>{project.endDate.toLocaleDateString("pt-BR")}</dd>
            <dt className="text-slate-500">Baseline</dt>
            <dd>{project.baselineDate?.toLocaleDateString("pt-BR") ?? "—"}</dd>
            <dt className="text-slate-500">Prioridade</dt>
            <dd>{project.priority}</dd>
          </dl>
        </div>
        <div className="rounded-lg border bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Próximos passos (sessões futuras)
          </h3>
          <ul className="space-y-2 text-sm text-slate-600">
            <li>📥 Importação Excel/MS Project — Sessão 5</li>
            <li>📏 Medição + janela HB — Sessão 6</li>
            <li>📈 Painel sinóptico + Curva S — Sessão 7</li>
            <li>🐞 Issues / Risks / CRs — Sessão 8</li>
            <li>📊 Relatórios — Sessão 9</li>
          </ul>
        </div>
      </section>
    </div>
  );
}

function Stat({ label, value, href }: { label: string; value: number; href?: string }) {
  const inner = (
    <div className="rounded-lg border bg-white p-4 transition hover:shadow-sm">
      <div className="text-sm text-slate-500">{label}</div>
      <div className="mt-1 text-2xl font-semibold">{value}</div>
    </div>
  );
  return href ? <Link href={href}>{inner}</Link> : inner;
}
