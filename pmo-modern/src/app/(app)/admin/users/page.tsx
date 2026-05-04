import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { UserCRUD } from "@/components/admin/UserCRUD";

export const metadata = { title: "Usuários — Admin" };

export default async function UsersAdminPage() {
  const [rows, unidades, areas] = await Promise.all([
    prisma.user.findMany({
      include: { unidade: { select: { code: true } }, area: { select: { code: true } } },
      orderBy: [{ active: "desc" }, { name: "asc" }],
    }),
    prisma.unidade.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
    prisma.area.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
  ]);
  return (
    <div className="space-y-4">
      <AdminHeader title="Usuários" description="Cadastro, papel, unidade/área, reset de senha e ativar/desativar." />
      <UserCRUD rows={rows} unidades={unidades} areas={areas} />
    </div>
  );
}
