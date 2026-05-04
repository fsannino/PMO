import { ComingSoon } from "@/components/ComingSoon";

export const metadata = { title: "Change Requests — PMO Sinergia" };

export default function CRsPage() {
  return (
    <ComingSoon
      title="Change Requests"
      session={8}
      description="Solicitações de mudança com workflow, incrementos e alterações."
    />
  );
}
