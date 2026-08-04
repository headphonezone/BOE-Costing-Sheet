import type { FieldStatus } from "@/lib/types";

export function StatusBadge({ status }: { status: FieldStatus | null | undefined }) {
  const isFixed = status === "fixed";
  return (
    <span
      className={`inline-flex items-center rounded-full px-2 py-0.5 text-[11px] font-semibold uppercase tracking-wide ${
        isFixed
          ? "bg-emerald-100 text-emerald-800 dark:bg-emerald-900 dark:text-emerald-200"
          : "bg-amber-100 text-amber-800 dark:bg-amber-900 dark:text-amber-200"
      }`}
    >
      {isFixed ? "Fixed" : "Provisional"}
    </span>
  );
}
