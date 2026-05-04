import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { AccessMatrix } from "@/components/admin/AccessMatrix";

export const metadata = { title: "Matriz de acesso — Admin" };

export default async function AccessAdminPage() {
  const [projects, users, accesses] = await Promise.all([
    prisma.project.findMany({
      select: { id: true, code: true, name: true, module: true },
      orderBy: { code: "asc" },
    }),
    prisma.user.findMany({
      select: { id: true, name: true, email: true, role: true, active: true },
      orderBy: { name: "asc" },
    }),
    prisma.access.findMany({
      select: { userId: true, projectId: true, module: true, canRead: true, canWrite: true, canAdmin: true },
    }),
  ]);

  return (
    <div className="space-y-4">
      <AdminHeader
        title="Matriz de acesso"
        description="Quem pode ler/escrever/administrar cada projeto + módulo. ADMIN tem bypass total."
      />
      {projects.length === 0 ? (
        <div className="rounded border bg-white p-8 text-center text-sm text-slate-500">
          Cadastre ao menos um projeto para configurar acessos.
        </div>
      ) : (
        <AccessMatrix projects={projects} users={users} accesses={accesses} />
      )}
    </div>
  );
}
