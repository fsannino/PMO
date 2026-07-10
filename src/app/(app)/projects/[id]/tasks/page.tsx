import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { TaskTreeView, type TaskRow } from "@/components/tasks/TaskTreeView";

export const metadata = { title: "Tarefas" };

export default async function TasksPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const session = await getServerSession(authOptions);
  if (!session?.user) return null;

  const project = await prisma.project.findUnique({
    where: { id },
    select: { id: true, module: true },
  });
  if (!project) notFound();
  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) redirect("/forbidden");

  const canWrite = await hasAccess(session.user.id, project.id, moduleParsed.data, "write");

  const [tasks, users, areas, frentes, equipes] = await Promise.all([
    prisma.task.findMany({
      where: { projectId: project.id },
      include: {
        assignee: { select: { name: true } },
        equipe: { select: { code: true } },
        area: { select: { code: true } },
        predecessors: {
          select: {
            id: true,
            predecessorId: true,
            predecessor: { select: { name: true, wbs: true } },
          },
        },
      },
      orderBy: [{ wbs: "asc" }, { startDate: "asc" }],
    }),
    prisma.user.findMany({
      where: { active: true },
      select: { id: true, name: true, email: true },
      orderBy: { name: "asc" },
    }),
    prisma.area.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
    prisma.frente.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
    prisma.equipe.findMany({ where: { active: true }, orderBy: { code: "asc" } }),
  ]);

  return (
    <TaskTreeView
      projectId={project.id}
      rows={tasks as TaskRow[]}
      users={users}
      areas={areas}
      frentes={frentes}
      equipes={equipes}
      canWrite={canWrite}
    />
  );
}
