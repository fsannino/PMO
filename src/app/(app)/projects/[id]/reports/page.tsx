import { prisma } from "@/lib/db";
import { notFound } from "next/navigation";
import { ReportsGallery } from "@/components/reports/ReportsGallery";

export const metadata = { title: "Relatórios — CollabZ" };

export default async function ProjectReportsPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const project = await prisma.project.findUnique({ where: { id }, select: { id: true } });
  if (!project) notFound();
  return (
    <div className="space-y-4">
      <p className="text-sm text-slate-600">
        6 relatórios padrão. Clique em <strong>Preview</strong> para ver no browser ou nos botões
        XLSX / PDF para baixar direto.
      </p>
      <ReportsGallery projectId={project.id} />
    </div>
  );
}
