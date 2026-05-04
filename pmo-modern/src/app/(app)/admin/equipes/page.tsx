import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { SimpleEntityCRUD } from "@/components/admin/SimpleEntityCRUD";

export const metadata = { title: "Equipes — Admin" };

export default async function EquipesAdminPage() {
  const rows = await prisma.equipe.findMany({ orderBy: { code: "asc" } });
  return (
    <div className="space-y-4">
      <AdminHeader title="Equipes" description="Equipes de execução (associadas às tarefas do cronograma)." />
      <SimpleEntityCRUD kind="equipe" labelSingular="equipe" rows={rows} />
    </div>
  );
}
