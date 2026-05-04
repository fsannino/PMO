import { redirect } from "next/navigation";
import Link from "next/link";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import {
  Users,
  FolderKanban,
  Building2,
  Network,
  ShieldCheck,
  CalendarClock,
  Layers3,
  GitBranch,
  Briefcase,
} from "lucide-react";

export const metadata = { title: "Administração — PMO CollabZ" };

export default async function AdminPage() {
  const session = await getServerSession(authOptions);
  if (!session?.user || session.user.role !== "ADMIN") redirect("/forbidden");

  const [users, projects, areas, frentes, governancas, equipes, comites, unidades, accesses, configs] =
    await Promise.all([
      prisma.user.count(),
      prisma.project.count(),
      prisma.area.count(),
      prisma.frente.count(),
      prisma.governanca.count(),
      prisma.equipe.count(),
      prisma.comite.count(),
      prisma.unidade.count(),
      prisma.access.count(),
      prisma.configFechamento.count(),
    ]);

  const cards = [
    { href: "/admin/users", title: "Usuários", icon: Users, count: users, desc: "Cadastro, papel, ativar/desativar, reset de senha." },
    { href: "/admin/projects", title: "Projetos", icon: FolderKanban, count: projects, desc: "Cadastro, módulo, datas, owner, status." },
    { href: "/admin/access", title: "Matriz de acesso", icon: ShieldCheck, count: accesses, desc: "Quem acessa qual projeto/módulo (Access)." },
    { href: "/admin/config-fechamento", title: "Janelas HB", icon: CalendarClock, count: configs, desc: "Horários de medição por dia da semana." },
    { href: "/admin/unidades", title: "Unidades", icon: Building2, count: unidades, desc: "REFAP, RECAP, etc." },
    { href: "/admin/areas", title: "Áreas", icon: Layers3, count: areas, desc: "TI, Manutenção, Operações, Engenharia, etc." },
    { href: "/admin/frentes", title: "Frentes", icon: GitBranch, count: frentes, desc: "Implantação, Testes, Operação Assistida." },
    { href: "/admin/governancas", title: "Governanças", icon: Network, count: governancas, desc: "Estruturas de governança." },
    { href: "/admin/equipes", title: "Equipes", icon: Users, count: equipes, desc: "Equipes de execução do cronograma." },
    { href: "/admin/comites", title: "Comitês", icon: Briefcase, count: comites, desc: "Comitês para Change Requests." },
  ];

  return (
    <div className="space-y-6">
      <header>
        <h1 className="text-2xl font-semibold">Administração</h1>
        <p className="text-sm text-slate-500">
          Substitui o aplicativo legado <code>Manutencao_SMR.exe</code>.
        </p>
      </header>

      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
        {cards.map((c) => {
          const Icon = c.icon;
          return (
            <Link
              key={c.href}
              href={c.href}
              className="flex items-start gap-3 rounded-lg border bg-white p-4 shadow-sm transition hover:shadow"
            >
              <span className="grid h-10 w-10 shrink-0 place-items-center rounded bg-brand-50 text-brand-700">
                <Icon className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <div className="flex items-center justify-between gap-2">
                  <h2 className="font-medium text-slate-900">{c.title}</h2>
                  <span className="rounded bg-slate-100 px-2 py-0.5 text-xs font-mono text-slate-700">
                    {c.count}
                  </span>
                </div>
                <p className="mt-1 text-xs text-slate-500">{c.desc}</p>
              </div>
            </Link>
          );
        })}
      </div>
    </div>
  );
}
