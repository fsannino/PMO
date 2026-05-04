"use client";

import { signOut, useSession } from "next-auth/react";
import { LogOut, User } from "lucide-react";
import { useState, useRef, useEffect } from "react";

export function UserMenu() {
  const { data: session } = useSession();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function onDocClick(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  if (!session?.user) return null;

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-2 rounded-md px-3 py-2 text-sm hover:bg-slate-100"
      >
        <span className="grid h-8 w-8 place-items-center rounded-full bg-brand-100 text-brand-700">
          <User className="h-4 w-4" />
        </span>
        <span className="hidden text-left md:block">
          <span className="block font-medium">{session.user.name}</span>
          <span className="block text-xs text-slate-500">{session.user.role}</span>
        </span>
      </button>
      {open && (
        <div className="absolute right-0 mt-2 w-56 rounded-md border bg-white py-1 shadow-lg">
          <div className="px-3 py-2 text-xs text-slate-500">
            {session.user.email}
          </div>
          <button
            onClick={() => signOut({ callbackUrl: "/login" })}
            className="flex w-full items-center gap-2 px-3 py-2 text-sm text-slate-700 hover:bg-slate-100"
          >
            <LogOut className="h-4 w-4" />
            Sair
          </button>
        </div>
      )}
    </div>
  );
}
