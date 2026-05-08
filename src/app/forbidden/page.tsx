import Link from "next/link";

export const metadata = { title: "Acesso negado — Collab:Build" };

export default function ForbiddenPage() {
  return (
    <div className="grid min-h-screen place-items-center bg-slate-50 p-4">
      <div className="max-w-md rounded-lg border bg-white p-8 text-center shadow-sm">
        <div className="text-6xl">🚫</div>
        <h1 className="mt-4 text-xl font-semibold text-slate-900">Acesso negado</h1>
        <p className="mt-2 text-sm text-slate-600">
          Você não tem permissão para acessar este recurso. Se acredita que isso
          é um engano, entre em contato com o administrador do PMO.
        </p>
        <Link
          href="/"
          className="mt-6 inline-block rounded-md bg-brand-600 px-4 py-2 text-sm font-medium text-white hover:bg-brand-700"
        >
          Voltar ao início
        </Link>
      </div>
    </div>
  );
}
