import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { SimpleEntityCRUD } from "@/components/admin/SimpleEntityCRUD";

export const metadata = { title: "Comitês — Admin" };

export default async function ComitesAdminPage() {
  const rows = await prisma.comite.findMany({ orderBy: { code: "asc" } });
  return (
    <div className="space-y-4">
      <AdminHeader title="Comitês" description="Comitês para aprovação de Change Requests." />
      <SimpleEntityCRUD kind="comite" labelSingular="comitê" rows={rows} />
    </div>
  );
}
