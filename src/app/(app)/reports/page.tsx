import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import Link from "next/link";
import { FileBarChart } from "lucide-react";

export const metadata = { title: "Relatórios — Collab:Build" };

export default async function GlobalReportsPage() {
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
    select: { id: true, code: true, name: true, module: true },
    orderBy: [{ priority: "asc" }, { code: "asc" }],
  });

  return (
    <div className="space-y-4">
      <header>
        <h1 className="text-2xl font-semibold">Relatórios</h1>
        <p className="text-sm text-slate-500">
          Os 6 relatórios padrão estão disponíveis em cada projeto. Selecione abaixo.
        </p>
      </header>

      {projects.length === 0 ? (
        <div className="rounded border bg-white p-8 text-center text-sm text-slate-500">
          Sem projetos acessíveis.
        </div>
      ) : (
        <ul className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
          {projects.map((p) => (
            <li key={p.id}>
              <Link
                href={`/projects/${p.id}/reports`}
                className="flex items-center gap-3 rounded-lg border bg-white p-3 hover:shadow-sm"
              >
                <span className="grid h-9 w-9 place-items-center rounded bg-brand-50 text-brand-700">
                  <FileBarChart className="h-4 w-4" />
                </span>
                <div className="min-w-0">
                  <div className="font-mono text-xs text-slate-500">{p.code}</div>
                  <div className="truncate font-medium">{p.name}</div>
                </div>
                <span className="ml-auto rounded bg-slate-100 px-2 py-0.5 text-xs">{p.module}</span>
              </Link>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
