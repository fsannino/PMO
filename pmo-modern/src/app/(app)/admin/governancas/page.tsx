import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { SimpleEntityCRUD } from "@/components/admin/SimpleEntityCRUD";

export const metadata = { title: "Governanças — Admin" };

export default async function GovernancasAdminPage() {
  const rows = await prisma.governanca.findMany({ orderBy: { code: "asc" } });
  return (
    <div className="space-y-4">
      <AdminHeader title="Governanças" description="Estruturas de governança (associadas a projetos)." />
      <SimpleEntityCRUD kind="governanca" labelSingular="governança" rows={rows} />
    </div>
  );
}
