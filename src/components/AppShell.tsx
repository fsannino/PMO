import type { ReactNode } from "react";
import { Sidebar } from "./Sidebar";
import { ToastProvider } from "./ui/toast";
import type { Role } from "@/lib/enums";

type User = { name?: string | null; email?: string | null; role: Role };

export function AppShell({ children, user }: { children: ReactNode; user: User }) {
  return (
    <ToastProvider>
      <div className="flex min-h-screen">
        <Sidebar role={user.role} />
        <div className="flex flex-1 flex-col overflow-hidden">
          <header className="flex h-14 shrink-0 items-center justify-between border-b border-gray-200 bg-white px-6">
            <div />
            <span className="text-sm text-gray-600">
              Olá,{" "}
              <span className="font-semibold text-gray-900">{user.name ?? user.email}</span>
            </span>
          </header>
          <main className="flex-1 overflow-auto p-6">{children}</main>
        </div>
      </div>
    </ToastProvider>
  );
}
