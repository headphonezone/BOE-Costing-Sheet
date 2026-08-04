import type { FieldStatus } from "@/lib/types";
import { StatusBadge } from "./StatusBadge";

export function SummaryTile({
  label,
  value,
  status,
  active,
  onClick,
}: {
  label: string;
  value: string;
  status?: FieldStatus | null;
  active?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`flex flex-col items-start gap-1 rounded-xl border px-4 py-3 text-left transition ${
        active
          ? "border-blue-500 bg-blue-50 dark:border-blue-500 dark:bg-blue-950"
          : "border-slate-200 bg-white hover:border-blue-300 dark:border-slate-800 dark:bg-slate-900"
      }`}
    >
      <span className="text-xs font-medium uppercase tracking-wide text-slate-400">
        {label}
      </span>
      <span className="text-lg font-semibold">{value}</span>
      {status && <StatusBadge status={status} />}
    </button>
  );
}
