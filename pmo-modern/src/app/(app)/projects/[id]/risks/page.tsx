import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { RisksList, type RiskRow } from "@/components/work-items/RisksList";

export const metadata = { title: "Riscos — CollabZ" };

export default async function ProjectRisksPage({ params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const project = await prisma.project.findUnique({ where: { id: params.id }, select: { id: true, module: true } });
  if (!project) notFound();
  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");
  const canWrite = await hasAccess(session.user.id, project.id, moduleParsed.data, "write");

  const [rows, users] = await Promise.all([
    prisma.risk.findMany({
      where: { projectId: project.id, deletedAt: null },
      include: {
        owner: { select: { name: true } },
        comments: { include: { author: { select: { name: true } } }, orderBy: { createdAt: "asc" } },
      },
      orderBy: [{ exposure: "desc" }],
    }),
    prisma.user.findMany({ where: { active: true }, select: { id: true, name: true }, orderBy: { name: "asc" } }),
  ]);

  return (
    <RisksList
      projectId={project.id}
      rows={rows as unknown as RiskRow[]}
      users={users}
      canWrite={canWrite}
      currentUserId={session.user.id}
      isAdmin={session.user.role === "ADMIN"}
    />
  );
}
