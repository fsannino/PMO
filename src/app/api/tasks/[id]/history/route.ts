import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

export async function GET(_req: Request, { params }: { params: Promise<{ id: string }> }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "unauthorized" }, { status: 401 });

  const { id } = await params;
  const task = await prisma.task.findUnique({
    where: { id },
    select: { projectId: true, project: { select: { module: true } } },
  });
  if (!task) return NextResponse.json({ error: "not_found" }, { status: 404 });

  const moduleParsed = ModuleSchema.safeParse(task.project.module);
  if (!moduleParsed.success) return NextResponse.json({ error: "bad_module" }, { status: 400 });

  const ok = await hasAccess(session.user.id, task.projectId, moduleParsed.data, "read");
  if (!ok) return NextResponse.json({ error: "forbidden" }, { status: 403 });

  const entries = await prisma.taskHistory.findMany({
    where: { taskId: id },
    include: { user: { select: { name: true } } },
    orderBy: { createdAt: "desc" },
    take: 200,
  });
  return NextResponse.json({ entries });
}
