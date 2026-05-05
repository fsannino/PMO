import { prisma } from "@/lib/db";
import { AdminHeader } from "@/components/admin/AdminHeader";
import { ConfigFechamentoEditor } from "@/components/admin/ConfigFechamentoEditor";

export const metadata = { title: "Janelas HB — Admin" };

export default async function ConfigFechamentoPage() {
  const [projects, configs] = await Promise.all([
    prisma.project.findMany({
      select: { id: true, code: true, name: true },
      orderBy: { code: "asc" },
    }),
    prisma.configFechamento.findMany({
      select: { projectId: true, dayOfWeek: true, enabled: true, startTime: true, endTime: true },
    }),
  ]);

  return (
    <div className="space-y-4">
      <AdminHeader
        title="Janelas HB (CONFIG_FECHAMENTO)"
        description="Horário em que medições podem ser registradas/confirmadas. Configuração global é o fallback; projetos podem sobrescrever."
      />
      <ConfigFechamentoEditor projects={projects} configs={configs} />
    </div>
  );
}
