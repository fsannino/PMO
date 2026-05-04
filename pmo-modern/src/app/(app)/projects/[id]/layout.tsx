import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { ProjectTabs } from "@/components/ProjectTabs";

export default async function ProjectLayout({
  children,
  params,
}: {
  children: React.ReactNode;
  params: { id: string };
}) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;

  const project = await prisma.project.findUnique({
    where: { id: params.id },
    select: { id: true, code: true, name: true, module: true },
  });
  if (!project) notFound();

  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");

  const ok = await hasAccess(session.user.id, project.id, moduleParsed.data, "read");
  if (!ok) redirect("/forbidden");

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-baseline justify-between gap-2">
        <div>
          <span className="font-mono text-xs text-slate-500">{project.code}</span>
          <h1 className="text-xl font-semibold">{project.name}</h1>
        </div>
        <span className="rounded bg-brand-50 px-2 py-0.5 text-xs font-medium text-brand-700">
          {project.module}
        </span>
      </header>
      <ProjectTabs projectId={project.id} />
      <div className="pt-2">{children}</div>
    </div>
  );
}
