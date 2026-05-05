import { prisma } from "@/lib/db";
import { notFound } from "next/navigation";
import { GanttChart, type GanttTask } from "@/components/gantt/GanttChart";
import { buildTree, flattenTree } from "@/lib/task-tree";

export const metadata = { title: "Gantt — CollabZ" };

export default async function GanttPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const project = await prisma.project.findUnique({
    where: { id },
    select: { id: true },
  });
  if (!project) notFound();

  const tasks = await prisma.task.findMany({
    where: { projectId: project.id, deletedAt: null },
    select: {
      id: true,
      name: true,
      wbs: true,
      startDate: true,
      endDate: true,
      baselineStart: true,
      baselineEnd: true,
      percentDone: true,
      status: true,
      isMilestone: true,
      isSummary: true,
      parentId: true,
      durationDays: true,
      deletedAt: true,
      predecessors: { select: { predecessorId: true } },
    },
    orderBy: [{ wbs: "asc" }, { startDate: "asc" }],
  });

  // ordenar respeitando hierarquia (pai vem antes dos filhos)
  const tree = buildTree(tasks);
  const ordered = flattenTree(tree);

  const ganttTasks: GanttTask[] = ordered.map((t) => ({
    id: t.id,
    name: t.name,
    wbs: t.wbs,
    startDate: t.startDate,
    endDate: t.endDate,
    baselineStart: t.baselineStart,
    baselineEnd: t.baselineEnd,
    percentDone: t.percentDone,
    status: t.status,
    isMilestone: t.isMilestone,
    isSummary: t.isSummary,
    parentId: t.parentId,
    predecessorIds: (t as unknown as { predecessors: { predecessorId: string }[] }).predecessors.map((p) => p.predecessorId),
  }));

  if (ganttTasks.length === 0) {
    return (
      <div className="rounded-lg border bg-white p-8 text-center">
        <p className="text-sm text-slate-500">
          Sem tarefas para exibir. Adicione tarefas em <strong>Tarefas</strong>.
        </p>
      </div>
    );
  }

  return <GanttChart tasks={ganttTasks} />;
}
