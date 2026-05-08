"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import {
  LayoutDashboard,
  FolderKanban,
  AlertTriangle,
  Bug,
  GitPullRequestArrow,
  FileBarChart,
  Settings,
  Power,
} from "lucide-react";
import { cn } from "@/lib/utils";
import type { Role } from "@/lib/enums";
import { signOut } from "next-auth/react";

type NavItem = {
  href: string;
  label: string;
  icon: React.ComponentType<{ className?: string }>;
  adminOnly?: boolean;
};

const ITEMS: NavItem[] = [
  { href: "/", label: "Dashboard", icon: LayoutDashboard },
  { href: "/projects", label: "Projetos", icon: FolderKanban },
  { href: "/issues", label: "Issues", icon: Bug },
  { href: "/risks", label: "Riscos", icon: AlertTriangle },
  { href: "/change-requests", label: "Change Requests", icon: GitPullRequestArrow },
  { href: "/reports", label: "Relatórios", icon: FileBarChart },
  { href: "/admin", label: "Administração", icon: Settings, adminOnly: true },
];

function PmoLogo() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect x="1" y="9" width="17" height="17" rx="3" fill="#3b82f6" opacity="0.85" />
      <rect x="12" y="1" width="17" height="17" rx="3" fill="#06b6d4" />
    </svg>
  );
}

export function Sidebar({ role }: { role: Role }) {
  const pathname = usePathname();
  const items = ITEMS.filter((i) => !i.adminOnly || role === "ADMIN");

  return (
    <aside className="hidden w-64 shrink-0 flex-col bg-slate-900 text-white md:flex">
      {/* Logo */}
      <div className="flex items-center gap-3 border-b border-white/10 px-5 py-4">
        <PmoLogo />
        <div>
          <p className="text-base font-bold leading-tight tracking-wide">PMO CollabZ</p>
          <p className="text-[10px] uppercase tracking-widest text-slate-400">
            Gestão de Projetos
          </p>
        </div>
      </div>

      {/* Nav */}
      <nav className="flex-1 space-y-0.5 px-3 py-4">
        {items.map((item) => {
          const Icon = item.icon;
          const active =
            item.href === "/"
              ? pathname === "/"
              : pathname === item.href || pathname.startsWith(item.href + "/");
          return (
            <Link
              key={item.href}
              href={item.href}
              className={cn(
                "flex items-center gap-3 rounded-r-lg py-2.5 pl-3 pr-3 text-sm font-medium transition-colors",
                "border-l-2",
                active
                  ? "border-yellow-400 bg-white/10 text-white"
                  : "border-transparent text-slate-400 hover:bg-white/5 hover:text-white",
              )}
            >
              <Icon className={cn("h-4 w-4 shrink-0", active ? "text-yellow-400" : "")} />
              {item.label}
            </Link>
          );
        })}
      </nav>

      {/* Sair */}
      <div className="border-t border-white/10 px-3 py-3">
        <button
          onClick={() => signOut({ callbackUrl: "/login" })}
          className="flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-400 transition-colors hover:bg-white/5 hover:text-white"
        >
          <Power className="h-4 w-4 shrink-0" />
          Sair
        </button>
        <p className="mt-3 text-center text-[10px] text-slate-600">© 2026 PMO CollabZ</p>
      </div>
    </aside>
  );
}
