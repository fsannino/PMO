import { LoginForm } from "@/components/LoginForm";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { redirect } from "next/navigation";
import { Suspense } from "react";

export const metadata = { title: "Entrar — PMO CollabZ" };

export default async function LoginPage() {
  const session = await getServerSession(authOptions);
  if (session) redirect("/");

  return (
    <div className="grid min-h-screen place-items-center bg-slate-100 p-4">
      <div className="w-full max-w-sm rounded-lg border bg-white p-6 shadow">
        <div className="mb-6 text-center">
          <h1 className="text-xl font-semibold text-brand-700">PMO CollabZ</h1>
          <p className="text-sm text-slate-500">Acesso ao sistema</p>
        </div>
        <Suspense fallback={null}>
          <LoginForm />
        </Suspense>
        <details className="mt-6 text-xs text-slate-500">
          <summary className="cursor-pointer">Logins de demonstração</summary>
          <ul className="mt-2 space-y-1">
            <li>admin@pmo.local / admin123 (ADMIN)</li>
            <li>gerente@pmo.local / gerente123 (MANAGER)</li>
            <li>alpha@pmo.local / membro123 (MEMBER)</li>
            <li>viewer@pmo.local / viewer123 (VIEWER)</li>
          </ul>
        </details>
      </div>
    </div>
  );
}
