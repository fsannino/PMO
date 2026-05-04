import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { redirect } from "next/navigation";
import { ComingSoon } from "@/components/ComingSoon";

export const metadata = { title: "Administração — PMO Sinergia" };

export default async function AdminPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user || session.user.role !== "ADMIN") redirect("/forbidden");
  return (
    <ComingSoon
      title="Administração"
      session={3}
      description="CRUDs admin: usuários, projetos, áreas/frentes/governanças/equipes, matriz de acesso, janelas HB, módulos."
    />
  );
}
