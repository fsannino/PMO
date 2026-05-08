import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { CRsList, type CRRow } from "@/components/work-items/CRsList";

export const metadata = { title: "Change Requests — Collab:Build" };

export default async function ProjectCRsPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const project = await prisma.project.findUnique({ where: { id }, select: { id: true, module: true } });
  if (!project) notFound();
  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");
  const canWrite = await hasAccess(session.user.id, project.id, moduleParsed.data, "write");

  const [rows, users, comites] = await Promise.all([
    prisma.changeRequest.findMany({
      where: { projectId: project.id, deletedAt: null },
      include: {
        owner: { select: { name: true } },
        comite: { select: { code: true, name: true } },
        comments: { include: { author: { select: { name: true } } }, orderBy: { createdAt: "asc" } },
        increments: { orderBy: { sequence: "asc" } },
      },
      orderBy: [{ status: "asc" }, { dueDate: "asc" }],
    }),
    prisma.user.findMany({ where: { active: true }, select: { id: true, name: true }, orderBy: { name: "asc" } }),
    prisma.comite.findMany({ where: { active: true }, select: { id: true, name: true, code: true }, orderBy: { code: "asc" } }),
  ]);

  return (
    <CRsList
      projectId={project.id}
      rows={rows as unknown as CRRow[]}
      users={users}
      comites={comites}
      canWrite={canWrite}
      currentUserId={session.user.id}
      isAdmin={session.user.role === "ADMIN"}
    />
  );
}
