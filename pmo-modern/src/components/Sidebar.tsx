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
} from "lucide-react";
import { cn } from "@/lib/utils";
import type { Role } from "@/lib/enums";

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

export function Sidebar({ role }: { role: Role }) {
  const pathname = usePathname();
  const items = ITEMS.filter((i) => !i.adminOnly || role === "ADMIN");

  return (
    <aside className="hidden w-60 shrink-0 border-r bg-white md:block">
      <div className="flex h-14 items-center border-b px-4">
        <span className="text-lg font-semibold text-brand-700">PMO Sinergia</span>
      </div>
      <nav className="space-y-1 p-3">
        {items.map((item) => {
          const Icon = item.icon;
          const active = item.href === "/"
            ? pathname === "/"
            : pathname === item.href || pathname.startsWith(item.href + "/");
          return (
            <Link
              key={item.href}
              href={item.href}
              className={cn(
                "flex items-center gap-3 rounded-md px-3 py-2 text-sm",
                active
                  ? "bg-brand-50 font-medium text-brand-700"
                  : "text-slate-700 hover:bg-slate-100",
              )}
            >
              <Icon className="h-4 w-4" />
              {item.label}
            </Link>
          );
        })}
      </nav>
    </aside>
  );
}
