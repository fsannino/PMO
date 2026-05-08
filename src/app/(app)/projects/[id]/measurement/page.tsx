import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";
import { notFound, redirect } from "next/navigation";
import { MeasurementGrid, type MeasurementRow } from "@/components/measurement/MeasurementGrid";
import { getMeasurementWindow, periodOf } from "@/lib/closing-window";

export const metadata = { title: "Medição — CollabZ" };

export default async function MeasurementPage({ params }: { params: Promise<{ id: string }> }) {
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
  if (!canWrite) redirect("/forbidden");

  const period = periodOf();
  const win = await getMeasurementWindow(project.id);

  const tasks = await prisma.task.findMany({
    where: {
      projectId: project.id,
      deletedAt: null,
      isSummary: false,
      OR: [{ assigneeId: session.user.id }, { assigneeId: null }],
    },
    select: {
      id: true,
      wbs: true,
      name: true,
      startDate: true,
      endDate: true,
      percentDone: true,
      measurements: {
        where: { period, userId: session.user.id, confirmed: false },
        select: { percentDone: true, hoursWorked: true, comment: true },
        take: 1,
      },
      measurementLocks: {
        where: { period },
        select: { id: true },
        take: 1,
      },
    },
    orderBy: [{ wbs: "asc" }, { startDate: "asc" }],
  });

  const today = new Date();
  const rows: MeasurementRow[] = tasks.map((t) => {
    const draft = t.measurements[0];
    const daysLate = t.percentDone >= 100 ? 0 : Math.floor((today.getTime() - t.endDate.getTime()) / 86400000);
    return {
      id: t.id,
      wbs: t.wbs,
      name: t.name,
      startDate: t.startDate.toISOString(),
      endDate: t.endDate.toISOString(),
      taskPercent: t.percentDone,
      draftPercent: draft?.percentDone ?? null,
      draftHours: draft?.hoursWorked ?? null,
      draftComment: draft?.comment ?? null,
      locked: t.measurementLocks.length > 0,
      daysLate,
    };
  });

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-baseline justify-between gap-2 text-sm text-slate-600">
        <span>
          Período: <strong>{period}</strong>
          {win.config && (
            <span className="ml-3 text-xs text-slate-500">
              Janela {win.config.scope === "project" ? "(projeto)" : "(global)"}{" "}
              {win.config.startTime}–{win.config.endTime}
            </span>
          )}
        </span>
        <span className="text-xs">
          {rows.length} tarefa(s) · {rows.filter((r) => r.locked).length} confirmada(s)
        </span>
      </div>
      <MeasurementGrid
        projectId={project.id}
        rows={rows}
        windowOpen={win.open}
        windowReason={win.reason}
      />
    </div>
  );
}
