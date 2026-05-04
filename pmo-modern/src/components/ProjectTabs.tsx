"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { cn } from "@/lib/utils";

export function ProjectTabs({ projectId }: { projectId: string }) {
  const pathname = usePathname();
  const tabs = [
    { href: `/projects/${projectId}`, label: "Detalhes", exact: true },
    { href: `/projects/${projectId}/tasks`, label: "Tarefas" },
    { href: `/projects/${projectId}/gantt`, label: "Gantt" },
    { href: `/projects/${projectId}/import`, label: "Importar" },
  ];
  return (
    <nav className="border-b">
      <ul className="-mb-px flex gap-1">
        {tabs.map((t) => {
          const active = t.exact ? pathname === t.href : pathname.startsWith(t.href);
          return (
            <li key={t.href}>
              <Link
                href={t.href}
                className={cn(
                  "inline-block border-b-2 px-3 py-2 text-sm",
                  active
                    ? "border-brand-600 font-medium text-brand-700"
                    : "border-transparent text-slate-600 hover:text-brand-700",
                )}
              >
                {t.label}
              </Link>
            </li>
          );
        })}
      </ul>
    </nav>
  );
}
