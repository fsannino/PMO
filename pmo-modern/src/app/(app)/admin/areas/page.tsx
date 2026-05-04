import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { SimpleEntityCRUD } from "@/components/admin/SimpleEntityCRUD";

export const metadata = { title: "Áreas — Admin" };

export default async function AreasAdminPage() {
  const rows = await prisma.area.findMany({ orderBy: { code: "asc" } });
  return (
    <div className="space-y-4">
      <AdminHeader title="Áreas" description="Cadastros agrupados por área (TI, Manutenção, etc.)." />
      <SimpleEntityCRUD kind="area" labelSingular="área" rows={rows} />
    </div>
  );
}
