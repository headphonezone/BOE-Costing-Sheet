import Link from "next/link";
import { supabase } from "@/lib/supabase";
import type { Boe } from "@/lib/types";

export const dynamic = "force-dynamic";

function fmtMoney(n: number | null) {
  if (n === null || n === undefined) return "—";
  return n.toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

export default async function HomePage() {
  const { data, error } = await supabase
    .from("boes")
    .select("*")
    .order("updated_at", { ascending: false });

  const boes = (data ?? []) as Boe[];

  return (
    <main className="mx-auto max-w-6xl px-6 py-10">
      <div className="mb-8 flex items-baseline justify-between">
        <div>
          <h1 className="text-2xl font-bold">BOE Costing Dashboard</h1>
          <p className="mt-1 text-sm text-slate-500">
            {boes.length} Bill{boes.length === 1 ? "" : "s"} of Entry on file
          </p>
        </div>
        <Link
          href="/upload"
          className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white transition hover:bg-blue-700"
        >
          + Upload BOE
        </Link>
      </div>

      {error && (
        <div className="mb-6 rounded-lg border border-red-300 bg-red-50 px-4 py-3 text-sm text-red-700 dark:border-red-900 dark:bg-red-950 dark:text-red-300">
          Could not load BOEs: {error.message}
        </div>
      )}

      {boes.length === 0 && !error && (
        <div className="rounded-lg border border-dashed border-slate-300 px-6 py-12 text-center text-slate-500 dark:border-slate-700">
          No BOEs uploaded yet. Upload one through the backend to see it here.
        </div>
      )}

      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
        {boes.map((boe) => (
          <Link
            key={boe.be_no}
            href={`/boe/${boe.be_no}`}
            className="group rounded-xl border border-slate-200 bg-white p-5 shadow-sm transition hover:border-blue-400 hover:shadow-md dark:border-slate-800 dark:bg-slate-900"
          >
            <div className="flex items-start justify-between">
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
        ))}
      </div>
    </main>
  );
}
