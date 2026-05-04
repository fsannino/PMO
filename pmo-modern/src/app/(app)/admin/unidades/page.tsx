import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { SimpleEntityCRUD } from "@/components/admin/SimpleEntityCRUD";

export const metadata = { title: "Unidades — Admin" };

export default async function UnidadesAdminPage() {
  const rows = await prisma.unidade.findMany({ orderBy: { code: "asc" } });
  return (
    <div className="space-y-4">
      <AdminHeader title="Unidades" description="Unidades operacionais (REFAP, RECAP, etc.)." />
      <SimpleEntityCRUD kind="unidade" labelSingular="unidade" rows={rows} />
    </div>
  );
}
