import Link from "next/link";
import { ChevronLeft } from "lucide-react";

export function AdminHeader({
  title,
  description,
  back = "/admin",
  action,
}: {
  title: string;
  description?: string;
  back?: string | null;
  action?: React.ReactNode;
}) {
  return (
    <header className="flex flex-wrap items-end justify-between gap-3 border-b pb-4">
      <div>
        {back && (
          <Link
            href={back}
            className="mb-1 inline-flex items-center gap-1 text-xs text-slate-500 hover:text-brand-700"
          >
            <ChevronLeft className="h-3 w-3" /> voltar
          </Link>
        )}
        <h1 className="text-2xl font-semibold text-slate-900">{title}</h1>
        {description && <p className="text-sm text-slate-500">{description}</p>}
      </div>
      {action && <div>{action}</div>}
    </header>
  );
}
