import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import Link from "next/link";
import { StatusBadge } from "@/components/work-items/StatusBadge";

export const metadata = { title: "Issues — CollabZ" };

export default async function GlobalIssuesPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const isAdmin = session.user.role === "ADMIN";

  const projectWhere = isAdmin
    ? {}
    : { OR: [{ ownerId: session.user.id }, { members: { some: { userId: session.user.id } } }, { accesses: { some: { userId: session.user.id } } }] };

  const issues = await prisma.issue.findMany({
    where: { deletedAt: null, project: projectWhere },
    include: {
      project: { select: { id: true, code: true, name: true } },
      owner: { select: { name: true } },
    },
    orderBy: [{ status: "asc" }, { openedAt: "desc" }],
    take: 200,
  });

  return (
    <div className="space-y-4">
      <header>
        <h1 className="text-2xl font-semibold">Issues</h1>
        <p className="text-sm text-slate-500">Visão consolidada — todas as issues dos projetos a que você tem acesso ({issues.length}).</p>
      </header>

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Projeto</th>
              <th className="px-3 py-2 text-left">Título</th>
              <th className="px-3 py-2 text-left">Severidade</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-left">Aberta</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {issues.length === 0 && <tr><td colSpan={6} className="px-4 py-6 text-center text-slate-500">Sem issues.</td></tr>}
            {issues.map((i) => (
              <tr key={i.id} className="hover:bg-slate-50">
                <td className="px-3 py-2">
                  <Link href={`/projects/${i.project.id}/issues`} className="font-mono text-xs text-brand-700 hover:underline">{i.project.code}</Link>
                </td>
                <td className="px-3 py-2">
                  <Link href={`/projects/${i.project.id}/issues`} className="font-medium hover:underline">{i.title}</Link>
                </td>
                <td className="px-3 py-2"><StatusBadge value={i.severity} /></td>
                <td className="px-3 py-2"><StatusBadge value={i.status} /></td>
                <td className="px-3 py-2 text-xs text-slate-600">{i.owner?.name ?? "—"}</td>
                <td className="px-3 py-2 text-xs whitespace-nowrap">{i.openedAt.toLocaleDateString("pt-BR")}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
