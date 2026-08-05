import Link from "next/link";
import { supabase } from "@/lib/supabase";
import { BoeCard } from "@/components/BoeCard";
import type { Boe } from "@/lib/types";

export const dynamic = "force-dynamic";

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
          <BoeCard key={boe.be_no} boe={boe} />
        ))}
      </div>
    </main>
  );
}
