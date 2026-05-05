import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import Link from "next/link";
import { ImportPanel } from "@/components/import/ImportPanel";

export const metadata = { title: "Importar — CollabZ" };

export default async function ImportPage({ params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;

  const project = await prisma.project.findUnique({
    where: { id: params.id },
    select: { id: true, module: true },
  });
  if (!project) notFound();
  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");

  const canWrite = await hasAccess(session.user.id, project.id, moduleParsed.data, "write");
  if (!canWrite) redirect("/forbidden");

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <p className="text-sm text-slate-600">
          Carregue planilhas Excel (6 templates: IS, RK, AC, CR, TK, TC) ou XML do MS Project.
        </p>
        <Link
          href={`/projects/${project.id}/import/history`}
          className="text-sm text-brand-700 hover:underline"
        >
          Histórico de importações →
        </Link>
      </div>
      <ImportPanel projectId={project.id} />
    </div>
  );
}
