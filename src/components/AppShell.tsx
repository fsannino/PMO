import type { ReactNode } from "react";
import { Sidebar } from "./Sidebar";
import { ToastProvider } from "./ui/toast";
import type { Role } from "@/lib/enums";

type User = { name?: string | null; email?: string | null; role: Role };

function Avatar({ name }: { name?: string | null }) {
  const initial = name?.charAt(0).toUpperCase() ?? "U";
  return (
    <span className="grid h-8 w-8 shrink-0 place-items-center rounded-full bg-blue-700 text-xs font-bold text-white">
      {initial}
    </span>
  );
}

export function AppShell({
  children,
  user,
  breadcrumb,
}: {
  children: ReactNode;
  user: User;
  breadcrumb?: { label: string; description?: string };
}) {
  return (
    <ToastProvider>
      {/* Thin accent bar at top */}
      <div className="h-0.5 w-full bg-gradient-to-r from-blue-700 via-cyan-500 to-yellow-400" />

      <div className="flex min-h-[calc(100vh-2px)]">
        <Sidebar role={user.role} />
        <div className="flex flex-1 flex-col overflow-hidden">
          <header className="flex h-14 shrink-0 items-center justify-between border-b border-gray-200 bg-white px-6">
            {/* Breadcrumb */}
            <div className="flex items-center gap-2 text-sm">
              {breadcrumb ? (
                <>
                  <span className="font-semibold uppercase tracking-widest text-slate-500 text-xs">
                    {breadcrumb.label}
                  </span>
                  {breadcrumb.description && (
                    <>
                      <span className="text-slate-300">•</span>
                      <span className="text-slate-500">{breadcrumb.description}</span>
                    </>
                  )}
                </>
              ) : (
                <span className="text-xs font-semibold uppercase tracking-widest text-slate-400">
                  PMO CollabZ
                </span>
              )}
            </div>

            {/* User */}
            <div className="flex items-center gap-2.5">
              <span className="text-sm text-gray-600">
                Olá,{" "}
                <span className="font-semibold text-gray-900">{user.name ?? user.email}</span>
              </span>
              <Avatar name={user.name ?? user.email} />
            </div>
          </header>

          <main className="flex-1 overflow-auto p-6">{children}</main>
        </div>
      </div>
    </ToastProvider>
  );
}
