import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import Link from "next/link";

export const metadata = { title: "Projetos — PMO CollabZ" };

export default async function ProjectsPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const isAdmin = session.user.role === "ADMIN";

  const projects = await prisma.project.findMany({
    where: isAdmin
      ? {}
      : {
          OR: [
            { ownerId: session.user.id },
            { members: { some: { userId: session.user.id } } },
            { accesses: { some: { userId: session.user.id } } },
          ],
        },
    include: {
      unidade: true,
      owner: { select: { name: true } },
      _count: { select: { tasks: true, issues: true, risks: true } },
    },
    orderBy: [{ priority: "asc" }, { name: "asc" }],
  });

  return (
    <div className="space-y-4">
      <header className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-semibold">Projetos</h1>
          <p className="text-sm text-slate-500">{projects.length} projeto(s) acessível(is)</p>
        </div>
      </header>

      {projects.length === 0 ? (
        <div className="rounded-lg border bg-white p-8 text-center text-sm text-slate-500">
          Sem projetos para exibir.
        </div>
      ) : (
        <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
          {projects.map((p) => (
            <Link
              key={p.id}
              href={`/projects/${p.id}`}
              className="rounded-lg border bg-white p-4 shadow-sm transition hover:shadow"
            >
              <div className="flex items-start justify-between">
                <div>
                  <div className="text-xs font-mono text-slate-500">{p.code}</div>
                  <h2 className="mt-1 font-semibold text-slate-900">{p.name}</h2>
                </div>
                <span className="rounded bg-brand-50 px-2 py-0.5 text-xs font-medium text-brand-700">
                  {p.module}
                </span>
              </div>
              <p className="mt-2 line-clamp-2 text-sm text-slate-600">{p.description ?? "—"}</p>
              <div className="mt-3 flex gap-4 text-xs text-slate-500">
                <span>{p._count.tasks} tarefas</span>
                <span>{p._count.issues} issues</span>
                <span>{p._count.risks} riscos</span>
              </div>
              <div className="mt-3 flex items-center justify-between text-xs text-slate-500">
                <span>{p.unidade?.code ?? "—"}</span>
                <span>{p.owner.name}</span>
              </div>
            </Link>
          ))}
        </div>
      )}
    </div>
  );
}
