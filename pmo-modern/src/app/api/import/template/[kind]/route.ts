import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { ImportTemplateSchema } from "@/lib/enums";
import { buildTemplateXlsx } from "@/lib/import-excel";

export async function GET(_req: Request, { params }: { params: { kind: string } }) {
  const session = await getServerSession(authOptions);
  if (!session?.user) return NextResponse.json({ error: "unauthorized" }, { status: 401 });

  const parsed = ImportTemplateSchema.safeParse(params.kind.toUpperCase());
  if (!parsed.success) return NextResponse.json({ error: "invalid template" }, { status: 400 });

  const buffer = buildTemplateXlsx(parsed.data);
  return new NextResponse(new Uint8Array(buffer), {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="template_${parsed.data}.xlsx"`,
    },
  });
}
