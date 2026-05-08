import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import Link from "next/link";
import { StatusBadge } from "@/components/work-items/StatusBadge";

export const metadata = { title: "Riscos" };

export default async function GlobalRisksPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const isAdmin = session.user.role === "ADMIN";
  const projectWhere = isAdmin ? {} : {
    OR: [{ ownerId: session.user.id }, { members: { some: { userId: session.user.id } } }, { accesses: { some: { userId: session.user.id } } }],
  };

  const risks = await prisma.risk.findMany({
    where: { deletedAt: null, project: projectWhere },
    include: {
      project: { select: { id: true, code: true, name: true } },
      owner: { select: { name: true } },
    },
    orderBy: [{ exposure: "desc" }],
    take: 200,
  });

  return (
    <div className="space-y-4">
      <header>
        <h1 className="text-2xl font-semibold">Riscos</h1>
        <p className="text-sm text-slate-500">Riscos consolidados, ordenados por exposição decrescente ({risks.length}).</p>
      </header>

      <div className="overflow-x-auto rounded-lg border bg-white">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-500">
            <tr>
              <th className="px-3 py-2 text-left">Projeto</th>
              <th className="px-3 py-2 text-left">Título</th>
              <th className="px-3 py-2 text-right">Prob.</th>
              <th className="px-3 py-2 text-right">Imp.</th>
              <th className="px-3 py-2 text-right">Expo.</th>
              <th className="px-3 py-2 text-left">Status</th>
              <th className="px-3 py-2 text-left">Owner</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {risks.length === 0 && <tr><td colSpan={7} className="px-4 py-6 text-center text-slate-500">Sem riscos.</td></tr>}
            {risks.map((r) => (
              <tr key={r.id} className="hover:bg-slate-50">
                <td className="px-3 py-2">
                  <Link href={`/projects/${r.project.id}/risks`} className="font-mono text-xs text-brand-700 hover:underline">{r.project.code}</Link>
                </td>
                <td className="px-3 py-2">
                  <Link href={`/projects/${r.project.id}/risks`} className="font-medium hover:underline">{r.title}</Link>
                </td>
                <td className="px-3 py-2 text-right tabular-nums">{r.probability.toFixed(2)}</td>
                <td className="px-3 py-2 text-right tabular-nums">{r.impact.toFixed(2)}</td>
                <td className="px-3 py-2 text-right tabular-nums font-medium">{r.exposure.toFixed(2)}</td>
                <td className="px-3 py-2"><StatusBadge value={r.status} /></td>
                <td className="px-3 py-2 text-xs text-slate-600">{r.owner?.name ?? "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
