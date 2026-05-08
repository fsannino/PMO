import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { hasAccess } from "@/lib/access";
import { ModuleSchema } from "@/lib/enums";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";
import {
  REPORTS,
  reportHeader,
  reportDetalhado,
  reportConsolidado,
  reportCriticidade,
  reportComentarios,
  reportIssuesKPI,
  reportAtrasadas,
  type ReportId,
} from "@/lib/reports";
import { buildReportXlsx } from "@/lib/report-xlsx";
import { buildReportPdf } from "@/lib/report-pdf";

const VALID_IDS = new Set(Object.keys(REPORTS) as ReportId[]);

export async function GET(
  req: Request,
  { params }: { params: Promise<{ projectId: string; reportId: string }> },
) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "unauthorized" }, { status: 401 });

  const { projectId, reportId: reportIdParam } = await params;
  const project = await prisma.project.findUnique({
    where: { id: projectId },
    select: { id: true, code: true, module: true },
  });
  if (!project) return NextResponse.json({ error: "project not found" }, { status: 404 });
  const moduleParsed = ModuleSchema.safeParse(project.module);
  if (!moduleParsed.success) return NextResponse.json({ error: "bad module" }, { status: 400 });
  const ok = await hasAccess(session.user.id, project.id, moduleParsed.data, "read");
  if (!ok) return NextResponse.json({ error: "forbidden" }, { status: 403 });

  if (!VALID_IDS.has(reportIdParam as ReportId)) {
    return NextResponse.json({ error: "invalid report" }, { status: 400 });
  }
  const reportId = reportIdParam as ReportId;

  const url = new URL(req.url);
  const format = (url.searchParams.get("format") ?? "xlsx").toLowerCase();
  if (format !== "xlsx" && format !== "pdf") {
    return NextResponse.json({ error: "format must be xlsx or pdf" }, { status: 400 });
  }

  const header = await reportHeader(project.id, session.user.name ?? session.user.email ?? "—");
  if (!header) return NextResponse.json({ error: "could not build header" }, { status: 500 });

  let data: unknown;
  switch (reportId) {
    case "detalhado":   data = await reportDetalhado(project.id); break;
    case "consolidado": data = await reportConsolidado(project.id); break;
    case "criticidade": data = await reportCriticidade(project.id); break;
    case "comentarios": data = await reportComentarios(project.id); break;
    case "issues-kpi":  data = await reportIssuesKPI(project.id); break;
    case "atrasadas":   data = await reportAtrasadas(project.id); break;
  }

  const filenameBase = `${project.code}_${reportId}_${new Date().toISOString().slice(0, 10)}`;

  if (format === "xlsx") {
    const buffer = buildReportXlsx(reportId, header, data);
    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filenameBase}.xlsx"`,
      },
    });
  }
  const pdfBuffer = await buildReportPdf(reportId, header, data);
  return new NextResponse(new Uint8Array(pdfBuffer), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${filenameBase}.pdf"`,
    },
  });
}
