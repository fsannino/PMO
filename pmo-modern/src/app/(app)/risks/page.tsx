import { ComingSoon } from "@/components/ComingSoon";

export const metadata = { title: "Riscos — PMO Sinergia" };

export default function RisksPage() {
  return (
    <ComingSoon
      title="Riscos"
      session={8}
      description="Gestão de riscos: probabilidade × impacto, mitigação, status, comentários."
    />
  );
}
