import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { SimpleEntityCRUD } from "@/components/admin/SimpleEntityCRUD";

export const metadata = { title: "Frentes — Admin" };

export default async function FrentesAdminPage() {
  const rows = await prisma.frente.findMany({ orderBy: { code: "asc" } });
  return (
    <div className="space-y-4">
      <AdminHeader title="Frentes" description="Frentes de projeto: Implantação, Testes, Operação Assistida etc." />
      <SimpleEntityCRUD kind="frente" labelSingular="frente" rows={rows} />
    </div>
  );
}
