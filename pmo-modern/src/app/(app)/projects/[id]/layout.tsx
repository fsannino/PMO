import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { ProjectTabs } from "@/components/ProjectTabs";
import { ModuleBanner } from "@/components/ModuleBanner";
import { getModuleConfig } from "@/lib/modules";

export default async function ProjectLayout({
  children,
  params,
}: {
  children: React.ReactNode;
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;

  const project = await prisma.project.findUnique({
    where: { id },
    select: { id: true, code: true, name: true, module: true },
  });
  if (!project) notFound();

  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");

  const ok = await hasAccess(session.user.id, project.id, moduleParsed.data, "read");
  if (!ok) redirect("/forbidden");

  const moduleConfig = getModuleConfig(project.module);

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-baseline justify-between gap-2">
        <div>
          <span className="font-mono text-xs text-slate-500">{project.code}</span>
          <h1 className="text-xl font-semibold">{project.name}</h1>
        </div>
        <span className={`rounded px-2 py-0.5 text-xs font-medium ${moduleConfig.color}`}>
          {moduleConfig.emoji} {moduleConfig.label}
        </span>
      </header>
      <ModuleBanner module={project.module} />
      <ProjectTabs projectId={project.id} hiddenLabels={moduleConfig.tabsHidden ?? []} />
      <div className="pt-2">{children}</div>
    </div>
  );
}
