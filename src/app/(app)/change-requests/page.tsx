import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import Link from "next/link";
import { StatusBadge } from "@/components/work-items/StatusBadge";

export const metadata = { title: "Change Requests" };

export default async function GlobalCRsPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const isAdmin = session.user.role === "ADMIN";
  const projectWhere = isAdmin ? {} : {
    OR: [{ ownerId: session.user.id }, { members: { some: { userId: session.user.id } } }, { accesses: { some: { userId: session.user.id } } }],
  };

  const crs = await prisma.changeRequest.findMany({
    where: { deletedAt: null, project: projectWhere },
    include: {
      project: { select: { id: true, code: true, name: true } },
      owner: { select: { name: true } },
      comite: { select: { code: true } },
      _count: { select: { increments: true, comments: true } },
    },
    orderBy: [{ status: "asc" }, { dueDate: "asc" }],
    take: 200,
  });

  return (
    <div className="space-y-4">
      <header>
        <h1 className="text-2xl font-semibold">Change Requests</h1>
        <p className="text-sm text-slate-500">CRs consolidados ({crs.length}).</p>
      </header>

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Projeto</th>
              <th className="px-3 py-2 text-left">Nome</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Prioridade</th>
              <th className="px-3 py-2 text-left">Comitê</th>
              <th className="px-3 py-2 text-left">Owner</th>
              <th className="px-3 py-2 text-right">Inc.</th>
              <th className="px-3 py-2 text-right">Coments</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {crs.length === 0 && <tr><td colSpan={8} className="px-4 py-6 text-center text-slate-500">Sem CRs.</td></tr>}
            {crs.map((c) => (
              <tr key={c.id} className="hover:bg-slate-50">
                <td className="px-3 py-2">
                  <Link href={`/projects/${c.project.id}/change-requests`} className="font-mono text-xs text-brand-700 hover:underline">{c.project.code}</Link>
                </td>
                <td className="px-3 py-2">
                  <Link href={`/projects/${c.project.id}/change-requests`} className="font-medium hover:underline">{c.name}</Link>
                </td>
                <td className="px-3 py-2"><StatusBadge value={c.status} /></td>
                <td className="px-3 py-2"><StatusBadge value={c.priority} /></td>
                <td className="px-3 py-2 text-xs text-slate-600">{c.comite?.code ?? "—"}</td>
                <td className="px-3 py-2 text-xs text-slate-600">{c.owner?.name ?? "—"}</td>
                <td className="px-3 py-2 text-right tabular-nums text-xs">{c._count.increments}</td>
                <td className="px-3 py-2 text-right tabular-nums text-xs">{c._count.comments}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
