import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import Link from "next/link";
import { FolderKanban, AlertTriangle, Bug, GitPullRequestArrow } from "lucide-react";
import { projectScorecard } from "@/lib/scorecard";
import { cn } from "@/lib/utils";

export const metadata = { title: "Dashboard — Collab:Build" };

async function loadDashboard(userId: string, isAdmin: boolean) {
  // ADMIN vê tudo; demais veem só projetos onde é membro/tem access.
  const projectWhere = isAdmin
    ? {}
    : {
        OR: [
          { ownerId: userId },
          { members: { some: { userId } } },
          { accesses: { some: { userId } } },
        ],
      };

  const [projects, openIssues, openRisks, openCRs] = await Promise.all([
    prisma.project.findMany({
      where: projectWhere,
      include: {
        unidade: true,
        owner: { select: { name: true, email: true } },
        _count: {
          select: { tasks: true, issues: true, risks: true, changeRequests: true },
        },
      },
      orderBy: [{ priority: "asc" }, { name: "asc" }],
    }),
    prisma.issue.count({
      where: {
        deletedAt: null,
        status: { in: ["OPEN", "IN_PROGRESS"] },
        ...(isAdmin ? {} : { project: projectWhere }),
      },
    }),
    prisma.risk.count({
      where: {
        deletedAt: null,
        status: { in: ["IDENTIFIED", "ANALYZING", "MITIGATING"] },
        ...(isAdmin ? {} : { project: projectWhere }),
      },
    }),
    prisma.changeRequest.count({
      where: {
        deletedAt: null,
        status: { in: ["OPEN", "UNDER_REVIEW", "APPROVED", "IN_IMPLEMENTATION"] },
        ...(isAdmin ? {} : { project: projectWhere }),
      },
    }),
  ]);
  return { projects, openIssues, openRisks, openCRs };
}

export default async function DashboardPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null; // layout já redireciona

  const { projects, openIssues, openRisks, openCRs } = await loadDashboard(
    session.user.id,
    session.user.role === "ADMIN",
  );
  const scorecards = await Promise.all(projects.map((p) => projectScorecard(p.id)));

  const stats = [
    { label: "Projetos", value: projects.length, icon: FolderKanban, href: "/projects", tone: "bg-brand-50 text-brand-700" },
    { label: "Issues abertas", value: openIssues, icon: Bug, href: "/issues", tone: "bg-amber-50 text-amber-700" },
    { label: "Riscos ativos", value: openRisks, icon: AlertTriangle, href: "/risks", tone: "bg-rose-50 text-rose-700" },
    { label: "CRs em curso", value: openCRs, icon: GitPullRequestArrow, href: "/change-requests", tone: "bg-emerald-50 text-emerald-700" },
  ];

  return (
    <div className="space-y-6">
      <header>
        <h1 className="text-2xl font-semibold text-slate-900">
          Olá, {session.user.name?.split(" ")[0] ?? "usuário"} 👋
        </h1>
        <p className="text-sm text-slate-500">
          Visão geral dos projetos aos quais você tem acesso.
        </p>
      </header>

      <section className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        {stats.map((s) => {
          const Icon = s.icon;
          return (
            <Link
              key={s.label}
              href={s.href}
              className="rounded-lg border bg-white p-4 shadow-sm transition hover:shadow"
            >
              <div className="flex items-center justify-between">
                <span className="text-sm font-medium text-slate-500">{s.label}</span>
                <span className={`grid h-8 w-8 place-items-center rounded ${s.tone}`}>
                  <Icon className="h-4 w-4" />
                </span>
              </div>
              <div className="mt-3 text-3xl font-semibold text-slate-900">{s.value}</div>
            </Link>
          );
        })}
      </section>

      <section>
        <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
          Seus projetos
        </h2>
        {projects.length === 0 ? (
          <div className="rounded-lg border bg-white p-8 text-center text-sm text-slate-500">
            Você ainda não tem acesso a nenhum projeto. Fale com o administrador.
          </div>
        ) : (
          <div className="overflow-hidden rounded-lg border bg-white">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                <tr>
                  <th className="px-4 py-2 text-left">Código</th>
                  <th className="px-4 py-2 text-left">Nome</th>
                  <th className="px-4 py-2 text-left">Módulo</th>
                  <th className="px-4 py-2 text-left">Unidade</th>
                  <th className="px-4 py-2 text-left">Owner</th>
                  <th className="px-4 py-2 text-right">Tarefas</th>
                  <th className="px-4 py-2 text-left">Andamento</th>
                  <th className="px-4 py-2 text-right">Atrasadas</th>
                  <th className="px-4 py-2 text-right">Issues</th>
                  <th className="px-4 py-2 text-right">Riscos</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {projects.map((p, i) => {
                  const sc = scorecards[i];
                  const late = sc.delayed0to10 + sc.delayedMore10;
                  return (
                  <tr key={p.id} className="hover:bg-slate-50">
                    <td className="px-4 py-2">
                      <Link href={`/projects/${p.id}`} className="font-medium text-brand-700 hover:underline">
                        {p.code}
                      </Link>
                    </td>
                    <td className="px-4 py-2">{p.name}</td>
                    <td className="px-4 py-2">
                      <span className="rounded bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-700">
                        {p.module}
                      </span>
                    </td>
                    <td className="px-4 py-2 text-slate-600">{p.unidade?.code ?? "—"}</td>
                    <td className="px-4 py-2 text-slate-600">{p.owner.name}</td>
                    <td className="px-4 py-2 text-right tabular-nums">{p._count.tasks}</td>
                    <td className="px-4 py-2">
                      <div className="flex items-center gap-2">
                        <div className="h-1.5 w-20 overflow-hidden rounded-full bg-slate-100">
                          <div className="h-full bg-brand-500" style={{ width: `${sc.andamentoPct}%` }} />
                        </div>
                        <span className="text-xs tabular-nums text-slate-600">{sc.andamentoPct}%</span>
                      </div>
                    </td>
                    <td className={cn("px-4 py-2 text-right tabular-nums", late > 0 ? "text-rose-700" : "text-slate-400")}>
                      {late}
                    </td>
                    <td className="px-4 py-2 text-right tabular-nums">{p._count.issues}</td>
                    <td className="px-4 py-2 text-right tabular-nums">{p._count.risks}</td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </div>
  );
}
