"use client";

import { useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { deleteBoe } from "@/lib/db";
import type { Boe } from "@/lib/types";

function fmtMoney(n: number | null) {
  if (n === null || n === undefined) return "—";
  return n.toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

export function BoeCard({ boe }: { boe: Boe }) {
  const router = useRouter();
  const [deleting, setDeleting] = useState(false);

  async function handleDelete(e: React.MouseEvent) {
    e.preventDefault();
    e.stopPropagation();
    const ok = window.confirm(
      `Permanently delete BE No ${boe.be_no}? This removes all its items, licences, field history, and uploaded documents. This cannot be undone.`
    );
    if (!ok) return;
    setDeleting(true);
    try {
      await deleteBoe(boe.be_no);
      router.refresh();
    } catch (err) {
      window.alert(err instanceof Error ? err.message : "Failed to delete BOE");
      setDeleting(false);
    }
  }

  return (
    <Link
      href={`/boe/${boe.be_no}`}
      className="group relative rounded-xl border border-slate-200 bg-white p-5 shadow-sm transition hover:border-blue-400 hover:shadow-md dark:border-slate-800 dark:bg-slate-900"
    >
      <button
        type="button"
        onClick={handleDelete}
        disabled={deleting}
        title="Delete BOE"
        className="absolute top-3 right-3 rounded-md p-1 text-slate-300 opacity-0 transition hover:bg-red-50 hover:text-red-600 group-hover:opacity-100 disabled:opacity-50 dark:hover:bg-red-950 dark:hover:text-red-400"
      >
        {deleting ? (
          <span className="text-xs">…</span>
        ) : (
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" className="h-4 w-4">
            <path
              fillRule="evenodd"
              d="M8.75 1A2.75 2.75 0 0 0 6 3.75v.25H3.5a.75.75 0 0 0 0 1.5h.55l.72 10.09A2.75 2.75 0 0 0 7.51 18h4.98a2.75 2.75 0 0 0 2.74-2.41L15.95 5.5h.55a.75.75 0 0 0 0-1.5H14v-.25A2.75 2.75 0 0 0 11.25 1h-2.5ZM10 4.5h1.5v-.25a1.25 1.25 0 0 0-1.25-1.25h-2.5a1.25 1.25 0 0 0-1.25 1.25v.25H10Zm-3.5 2a.75.75 0 0 1 .75.75v7a.75.75 0 0 1-1.5 0v-7a.75.75 0 0 1 .75-.75Zm3.5 0a.75.75 0 0 1 .75.75v7a.75.75 0 0 1-1.5 0v-7a.75.75 0 0 1 .75-.75Zm3.5 0a.75.75 0 0 1 .75.75v7a.75.75 0 0 1-1.5 0v-7a.75.75 0 0 1 .75-.75Z"
              clipRule="evenodd"
            />
          </svg>
        )}
      </button>

      <div className="flex items-start justify-between pr-6">
        <div>
          <div className="text-xs font-medium uppercase tracking-wide text-slate-400">
            BE No
          </div>
          <div className="text-lg font-semibold group-hover:text-blue-600">
            {boe.be_no}
          </div>
        </div>
        <div className="text-right text-xs text-slate-400">
          {boe.be_date ?? "—"}
        </div>
      </div>

      <div className="mt-3 truncate text-sm text-slate-600 dark:text-slate-300">
        {boe.supplier_name ?? "Unknown supplier"}
      </div>
      <div className="mt-0.5 text-xs text-slate-400">
        Invoice {boe.inv_no ?? "—"}
      </div>

      <div className="mt-4 grid grid-cols-2 gap-3 border-t border-slate-100 pt-3 text-sm dark:border-slate-800">
        <div>
          <div className="text-xs text-slate-400">Assess Value</div>
          <div className="font-medium">₹{fmtMoney(boe.total_assess_value)}</div>
        </div>
        <div>
          <div className="text-xs text-slate-400">Total Duty</div>
          <div className="font-medium">₹{fmtMoney(boe.total_duty)}</div>
        </div>
      </div>
    </Link>
  );
}
