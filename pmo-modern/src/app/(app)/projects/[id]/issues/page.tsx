import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { IssuesList, type IssueRow } from "@/components/work-items/IssuesList";

export const metadata = { title: "Issues — CollabZ" };

export default async function ProjectIssuesPage({ params }: { params: { id: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;
  const project = await prisma.project.findUnique({ where: { id: params.id }, select: { id: true, module: true } });
  if (!project) notFound();
  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");
  const canWrite = await hasAccess(session.user.id, project.id, moduleParsed.data, "write");

  const [rows, users, areas, tasks] = await Promise.all([
    prisma.issue.findMany({
      where: { projectId: project.id, deletedAt: null },
      include: {
        owner: { select: { name: true } },
        area: { select: { code: true } },
        task: { select: { name: true, wbs: true } },
        comments: {
          include: { author: { select: { name: true } } },
          orderBy: { createdAt: "asc" },
        },
      },
      orderBy: [{ status: "asc" }, { openedAt: "desc" }],
    }),
    prisma.user.findMany({ where: { active: true }, select: { id: true, name: true }, orderBy: { name: "asc" } }),
    prisma.area.findMany({ where: { active: true }, select: { id: true, name: true, code: true }, orderBy: { code: "asc" } }),
    prisma.task.findMany({
      where: { projectId: project.id, deletedAt: null, isSummary: false },
      select: { id: true, name: true, wbs: true },
      orderBy: [{ wbs: "asc" }],
    }),
  ]);

  return (
    <IssuesList
      projectId={project.id}
      rows={rows as unknown as IssueRow[]}
      users={users}
      areas={areas}
      tasks={tasks}
      canWrite={canWrite}
      currentUserId={session.user.id}
      isAdmin={session.user.role === "ADMIN"}
    />
  );
}
