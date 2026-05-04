import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { ProjectCRUD } from "@/components/admin/ProjectCRUD";

export const metadata = { title: "Projetos — Admin" };

export default async function ProjectsAdminPage() {
  const [rows, users, unidades, governancas] = await Promise.all([
    prisma.project.findMany({
      include: {
        owner: { select: { name: true } },
        unidade: { select: { code: true } },
      },
      orderBy: [{ priority: "asc" }, { code: "asc" }],
    }),
    prisma.user.findMany({
      where: { active: true, role: { in: ["ADMIN", "MANAGER"] } },
      select: { id: true, name: true, email: true },
      orderBy: { name: "asc" },
    }),
    prisma.unidade.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
    prisma.governanca.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
  ]);
  return (
    <div className="space-y-4">
      <AdminHeader
        title="Projetos"
        description="Cadastro de projetos com módulo (CUT/GVI/GRF/TCP/TIN/PMO/LIGHT), status, datas, owner e prioridade."
      />
      <ProjectCRUD rows={rows} users={users} unidades={unidades} governancas={governancas} />
    </div>
  );
}
