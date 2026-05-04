import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import Link from "next/link";

export default async function ProjectDetailPage({
  params,
}: {
  params: { id: string };
}) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;

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

  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");

  const ok = await hasAccess(session.user.id, project.id, moduleParsed.data, "read");
  if (!ok) redirect("/forbidden");

  return (
    <div className="space-y-6">
      <header>
        <div className="flex items-baseline gap-3">
          <span className="font-mono text-sm text-slate-500">{project.code}</span>
          <span className="rounded bg-brand-50 px-2 py-0.5 text-xs font-medium text-brand-700">
            {project.module}
          </span>
          <span className="rounded bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-700">
            {project.status}
          </span>
        </div>
        <h1 className="mt-1 text-2xl font-semibold">{project.name}</h1>
        {project.description && (
          <p className="mt-2 text-sm text-slate-600">{project.description}</p>
        )}
      </header>

      <section className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <Stat label="Tarefas" value={project._count.tasks} />
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
          </dl>
        </div>
        <div className="rounded-lg border bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Próximos passos (sessões futuras)
          </h3>
          <ul className="space-y-2 text-sm text-slate-600">
            <li>📋 Cronograma + Gantt — Sessão 4</li>
            <li>📥 Importação Excel/MS Project — Sessão 5</li>
            <li>📏 Medição + janela HB — Sessão 6</li>
            <li>📈 Painel sinóptico + Curva S — Sessão 7</li>
            <li>🐞 Issues / Risks / CRs — Sessão 8</li>
            <li>📊 Relatórios — Sessão 9</li>
          </ul>
          <p className="mt-3 text-xs text-slate-400">
            Você está vendo este projeto porque tem acesso de leitura via{" "}
            {session.user.role === "ADMIN" ? "ADMIN bypass" : "matriz Access"}.
          </p>
        </div>
      </section>

      <Link
        href="/projects"
        className="inline-block text-sm text-brand-700 hover:underline"
      >
        ← voltar para a lista
      </Link>
    </div>
  );
}

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-lg border bg-white p-4">
      <div className="text-sm text-slate-500">{label}</div>
      <div className="mt-1 text-2xl font-semibold">{value}</div>
    </div>
  );
}
