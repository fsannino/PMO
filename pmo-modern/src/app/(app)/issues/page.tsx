import { ComingSoon } from "@/components/ComingSoon";

export const metadata = { title: "Issues — PMO CollabZ" };

export default function IssuesPage() {
  return (
    <ComingSoon
      title="Issues"
      session={8}
      description="Gestão de problemas/defeitos: cadastro, severidade, workflow, threads de comentários, KPI."
    />
  );
}
