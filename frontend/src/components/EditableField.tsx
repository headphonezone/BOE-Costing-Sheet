"use client";

import { useState } from "react";
import type { FieldStatus } from "@/lib/types";

export function EditableField({
  label,
  value,
  status,
  onChange,
}: {
  label: string;
  value: number;
  status: FieldStatus;
  onChange: (value: number, status: FieldStatus) => void | Promise<void>;
}) {
  const [draft, setDraft] = useState(String(value));
  const [saving, setSaving] = useState(false);

  const isFixed = status === "fixed";

  async function commitValue() {
    const parsed = parseFloat(draft);
    if (Number.isNaN(parsed) || parsed === value) {
      setDraft(String(value));
      return;
    }
    setSaving(true);
    try {
      await onChange(parsed, status);
    } finally {
      setSaving(false);
    }
  }

  async function toggleStatus() {
    setSaving(true);
    try {
      await onChange(value, isFixed ? "provisional" : "fixed");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div
      className={`rounded-xl border p-4 transition ${
        isFixed
          ? "border-emerald-300 bg-emerald-50 dark:border-emerald-800 dark:bg-emerald-950"
          : "border-amber-300 bg-amber-50 dark:border-amber-800 dark:bg-amber-950"
      }`}
    >
      <div className="mb-2 flex items-center justify-between">
        <span className="text-xs font-semibold uppercase tracking-wide text-slate-600 dark:text-slate-300">
          {label}
        </span>
        <button
          onClick={toggleStatus}
          disabled={saving}
          className={`rounded-full px-2 py-0.5 text-[11px] font-semibold uppercase tracking-wide transition disabled:opacity-50 ${
            isFixed
              ? "bg-emerald-600 text-white hover:bg-emerald-700"
              : "bg-amber-500 text-white hover:bg-amber-600"
          }`}
          title="Click to toggle Provisional / Fixed"
        >
          {isFixed ? "Fixed" : "Provisional"}
        </button>
      </div>
      <div className="flex items-center gap-2">
        <span className="text-slate-500">₹</span>
        <input
          type="number"
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commitValue}
          onKeyDown={(e) => {
            if (e.key === "Enter") (e.target as HTMLInputElement).blur();
          }}
          disabled={saving}
          className="w-full rounded-md border border-slate-300 bg-white px-2 py-1 text-right font-medium tabular-nums outline-none focus:border-blue-500 disabled:opacity-50 dark:border-slate-700 dark:bg-slate-900"
        />
      </div>
      {saving && <div className="mt-1 text-[11px] text-slate-400">Saving…</div>}
    </div>
  );
}
