import type { ReactNode } from "react";
import { Sidebar } from "./Sidebar";
import { UserMenu } from "./UserMenu";
import type { Role } from "@/lib/enums";

export function AppShell({
  children,
  role,
}: {
  children: ReactNode;
  role: Role;
}) {
  return (
    <div className="flex h-screen w-full">
      <Sidebar role={role} />
      <div className="flex min-w-0 flex-1 flex-col">
        <header className="flex h-14 items-center justify-between border-b bg-white px-4">
          <span className="text-sm text-slate-500 md:hidden">PMO CollabZ</span>
          <span className="hidden text-sm text-slate-500 md:block">
            Sistema integrado de gestão de projetos
          </span>
          <UserMenu />
        </header>
        <main className="min-h-0 flex-1 overflow-auto p-6">{children}</main>
      </div>
    </div>
  );
}
