import { cn } from "@/lib/utils";
import type { ReactNode } from "react";

type Accent = "blue" | "yellow" | "cyan" | "red" | "green" | "slate";

const ACCENT_CLASSES: Record<Accent, string> = {
  blue:   "border-t-blue-700",
  yellow: "border-t-yellow-400",
  cyan:   "border-t-cyan-500",
  red:    "border-t-red-500",
  green:  "border-t-green-500",
  slate:  "border-t-slate-400",
};

interface CardProps {
  children: ReactNode;
  accent?: Accent;
  className?: string;
  padding?: boolean;
}

export function Card({ children, accent = "blue", className, padding = true }: CardProps) {
  return (
    <div
      className={cn(
        "rounded-xl border border-gray-100 bg-white shadow-sm",
        "border-t-4",
        ACCENT_CLASSES[accent],
        padding ? "p-5" : "overflow-hidden",
        className,
      )}
    >
      {children}
    </div>
  );
}
