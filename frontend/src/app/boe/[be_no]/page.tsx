"use client";

import { use, useEffect, useRef, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { getBoeDetail, updateVariableField, deleteBoe, type BoeDetail } from "@/lib/db";
import { API_BASE_URL } from "@/lib/supabase";
import {
  VARIABLE_FIELDS,
  VARIABLE_FIELD_LABELS,
  type VariableFieldKey,
  type FieldStatus,
} from "@/lib/types";
import { SummaryTile } from "@/components/SummaryTile";
import { EditableField } from "@/components/EditableField";
import { StatusBadge } from "@/components/StatusBadge";
import { DocumentsPanel } from "@/components/DocumentsPanel";

function fmtMoney(n: number | null | undefined) {
  if (n === null || n === undefined) return "—";
  return n.toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

type Section = "overview" | "duty" | "licences" | "history" | "documents";

export default function BoeDashboardPage(props: PageProps<"/boe/[be_no]">) {
  const { be_no } = use(props.params);
  const router = useRouter();

  const [detail, setDetail] = useState<BoeDetail | null | undefined>(undefined);
  const [deleting, setDeleting] = useState(false);
  const [activeSection, setActiveSection] = useState<Section>("overview");
  const sectionRefs = {
    overview: useRef<HTMLDivElement>(null),
    duty: useRef<HTMLDivElement>(null),
    licences: useRef<HTMLDivElement>(null),
    history: useRef<HTMLDivElement>(null),
    documents: useRef<HTMLDivElement>(null),
  };

  async function reload() {
    const d = await getBoeDetail(be_no);
    setDetail(d);
  }

  useEffect(() => {
    reload();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [be_no]);

  function goTo(section: Section) {
    setActiveSection(section);
    sectionRefs[section].current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function handleFieldChange(field: VariableFieldKey, value: number, status: FieldStatus) {
    await updateVariableField(be_no, field, value, status, detail?.variableFields ?? null);
    await reload();
  }

  async function handleDelete() {
    const ok = window.confirm(
      `Permanently delete BE No ${be_no}? This removes all its items, licences, field history, and uploaded documents. This cannot be undone.`
    );
    if (!ok) return;
    setDeleting(true);
    try {
      await deleteBoe(be_no);
      router.push("/");
    } catch (e) {
      window.alert(e instanceof Error ? e.message : "Failed to delete BOE");
      setDeleting(false);
    }
  }

  if (detail === undefined) {
    return <main className="mx-auto max-w-6xl px-6 py-10 text-slate-500">Loading…</main>;
  }

  if (detail === null) {
    return (
      <main className="mx-auto max-w-6xl px-6 py-10">
        <Link href="/" className="text-sm text-blue-600 hover:underline">
          ← Back
        </Link>
        <p className="mt-6 text-slate-500">No BOE found for {be_no}.</p>
      </main>
    );
  }

  const { boe, items, licences, variableFields, fieldHistory, documents } = detail;

  const vfDefaults: Record<VariableFieldKey, { value: number; status: FieldStatus }> = {
    exchange_rate: { value: variableFields?.exchange_rate ?? boe.exchange_rate ?? 0, status: variableFields?.exchange_rate_status ?? "provisional" },
    freight_charges: { value: variableFields?.freight_charges ?? boe.freight_inr ?? 0, status: variableFields?.freight_charges_status ?? "provisional" },
    clearing_charges: { value: variableFields?.clearing_charges ?? 0, status: variableFields?.clearing_charges_status ?? "provisional" },
    supplier_freight: { value: variableFields?.supplier_freight ?? 0, status: variableFields?.supplier_freight_status ?? "provisional" },
    bank_charges: { value: variableFields?.bank_charges ?? 0, status: variableFields?.bank_charges_status ?? "provisional" },
    own_bank_charges: { value: variableFields?.own_bank_charges ?? 0, status: variableFields?.own_bank_charges_status ?? "provisional" },
  };

  const totalLicenceDebit = licences.reduce((s, l) => s + (l.debit_duty ?? 0), 0);
  const totalDuty = items.reduce((s, it) => s + (it.bcd ?? 0) + (it.sws ?? 0) + (it.igst ?? 0), 0);

  return (
    <main className="mx-auto max-w-6xl px-6 py-10">
      <Link href="/" className="text-sm text-blue-600 hover:underline">
        ← All BOEs
      </Link>

      <div className="mt-3 mb-6 flex flex-wrap items-start justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold">BE No {boe.be_no}</h1>
          <p className="mt-1 text-sm text-slate-500">
            {boe.supplier_name ?? "Unknown supplier"} · Invoice {boe.inv_no ?? "—"} · {boe.be_date ?? "—"}
          </p>
        </div>
        <div className="flex gap-2">
          <a
            href={`${API_BASE_URL}/boe/${be_no}/excel`}
            className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white transition hover:bg-blue-700"
          >
            Download Excel
          </a>
          <button
            type="button"
            onClick={handleDelete}
            disabled={deleting}
            className="rounded-lg border border-red-300 px-4 py-2 text-sm font-medium text-red-600 transition hover:bg-red-50 disabled:opacity-50 dark:border-red-900 dark:text-red-400 dark:hover:bg-red-950"
          >
            {deleting ? "Deleting…" : "Delete BOE"}
          </button>
        </div>
      </div>

      {/* Summary tiles */}
      <div className="mb-8 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
        <SummaryTile
          label="Exchange Rate"
          value={fmtMoney(vfDefaults.exchange_rate.value)}
          status={vfDefaults.exchange_rate.status}
          active={activeSection === "overview"}
          onClick={() => goTo("overview")}
        />
        <SummaryTile
          label="Freight Charges"
          value={`₹${fmtMoney(vfDefaults.freight_charges.value)}`}
          status={vfDefaults.freight_charges.status}
          active={activeSection === "overview"}
          onClick={() => goTo("overview")}
        />
        <SummaryTile
          label="Clearing Charges"
          value={`₹${fmtMoney(vfDefaults.clearing_charges.value)}`}
          status={vfDefaults.clearing_charges.status}
          active={activeSection === "overview"}
          onClick={() => goTo("overview")}
        />
        <SummaryTile
          label="Insurance"
          value={`₹${fmtMoney(boe.insurance_inr)}`}
          active={activeSection === "overview"}
          onClick={() => goTo("overview")}
        />
        <SummaryTile
          label="Total Duty"
          value={`₹${fmtMoney(totalDuty)}`}
          active={activeSection === "duty"}
          onClick={() => goTo("duty")}
        />
        <SummaryTile
          label="Licences"
          value={`${licences.length} · ₹${fmtMoney(totalLicenceDebit)}`}
          active={activeSection === "licences"}
          onClick={() => goTo("licences")}
        />
        <SummaryTile
          label="Assess Value"
          value={`₹${fmtMoney(boe.total_assess_value)}`}
          active={activeSection === "duty"}
          onClick={() => goTo("duty")}
        />
        <SummaryTile
          label="Field History"
          value={`${fieldHistory.length} change${fieldHistory.length === 1 ? "" : "s"}`}
          active={activeSection === "history"}
          onClick={() => goTo("history")}
        />
        <SummaryTile
          label="Documents"
          value={`${documents.length} file${documents.length === 1 ? "" : "s"}`}
          active={activeSection === "documents"}
          onClick={() => goTo("documents")}
        />
      </div>

      {/* Overview / variable fields */}
      <section ref={sectionRefs.overview} className="mb-10 scroll-mt-6">
        <h2 className="mb-3 text-lg font-semibold">Variable / Provisional Cost Items</h2>
        <p className="mb-4 text-sm text-slate-500">
          Yellow = Provisional, Green = Fixed. Editing here writes straight to the database.
        </p>
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {VARIABLE_FIELDS.map((field) => (
            <EditableField
              key={field}
              label={VARIABLE_FIELD_LABELS[field]}
              value={vfDefaults[field].value}
              status={vfDefaults[field].status}
              onChange={(value, status) => handleFieldChange(field, value, status)}
            />
          ))}
        </div>
      </section>

      {/* Duty / items */}
      <section ref={sectionRefs.duty} className="mb-10 scroll-mt-6">
        <h2 className="mb-3 text-lg font-semibold">Items &amp; Duty ({items.length})</h2>
        <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800">
          <table className="w-full text-sm">
            <thead className="bg-slate-100 text-left text-xs uppercase tracking-wide text-slate-500 dark:bg-slate-800">
              <tr>
                <th className="px-3 py-2">#</th>
                <th className="px-3 py-2">Description</th>
                <th className="px-3 py-2 text-right">Qty</th>
                <th className="px-3 py-2 text-right">Assess Value</th>
                <th className="px-3 py-2 text-right">BCD</th>
                <th className="px-3 py-2 text-right">SWS</th>
                <th className="px-3 py-2 text-right">IGST</th>
              </tr>
            </thead>
            <tbody>
              {items.map((it) => (
                <tr key={it.id} className="border-t border-slate-100 dark:border-slate-800">
                  <td className="px-3 py-2 text-slate-400">{it.global_sno}</td>
                  <td className="max-w-xs truncate px-3 py-2">{it.description}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{it.qty}</td>
                  <td className="px-3 py-2 text-right tabular-nums">₹{fmtMoney(it.assess_value)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">₹{fmtMoney(it.bcd)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">₹{fmtMoney(it.sws)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">₹{fmtMoney(it.igst)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Licences */}
      <section ref={sectionRefs.licences} className="mb-10 scroll-mt-6">
        <h2 className="mb-3 text-lg font-semibold">Licences ({licences.length})</h2>
        <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800">
          <table className="w-full text-sm">
            <thead className="bg-slate-100 text-left text-xs uppercase tracking-wide text-slate-500 dark:bg-slate-800">
              <tr>
                <th className="px-3 py-2">Item</th>
                <th className="px-3 py-2">Licence No</th>
                <th className="px-3 py-2 text-right">Debit Duty</th>
              </tr>
            </thead>
            <tbody>
              {licences.map((lic) => (
                <tr key={lic.id} className="border-t border-slate-100 dark:border-slate-800">
                  <td className="px-3 py-2 text-slate-400">{lic.itemsn}</td>
                  <td className="px-3 py-2 font-mono">{lic.lic_no}</td>
                  <td className="px-3 py-2 text-right tabular-nums">₹{fmtMoney(lic.debit_duty)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Field history / variance */}
      <section ref={sectionRefs.history} className="mb-10 scroll-mt-6">
        <h2 className="mb-3 text-lg font-semibold">Provisional → Fixed History</h2>
        {fieldHistory.length === 0 ? (
          <p className="text-sm text-slate-500">No changes recorded yet.</p>
        ) : (
          <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800">
            <table className="w-full text-sm">
              <thead className="bg-slate-100 text-left text-xs uppercase tracking-wide text-slate-500 dark:bg-slate-800">
                <tr>
                  <th className="px-3 py-2">Field</th>
                  <th className="px-3 py-2">Was</th>
                  <th className="px-3 py-2">Now</th>
                  <th className="px-3 py-2 text-right">Difference</th>
                  <th className="px-3 py-2">When</th>
                </tr>
              </thead>
              <tbody>
                {fieldHistory.map((h) => {
                  const diff = (h.new_value ?? 0) - (h.old_value ?? 0);
                  return (
                    <tr key={h.id} className="border-t border-slate-100 dark:border-slate-800">
                      <td className="px-3 py-2">{VARIABLE_FIELD_LABELS[h.field_name]}</td>
                      <td className="px-3 py-2">
                        <span className="tabular-nums">₹{fmtMoney(h.old_value)}</span>{" "}
                        {h.old_status && <StatusBadge status={h.old_status} />}
                      </td>
                      <td className="px-3 py-2">
                        <span className="tabular-nums">₹{fmtMoney(h.new_value)}</span>{" "}
                        {h.new_status && <StatusBadge status={h.new_status} />}
                      </td>
                      <td
                        className={`px-3 py-2 text-right tabular-nums ${
                          diff > 0 ? "text-red-600" : diff < 0 ? "text-emerald-600" : "text-slate-400"
                        }`}
                      >
                        {diff > 0 ? "+" : ""}
                        {fmtMoney(diff)}
                      </td>
                      <td className="px-3 py-2 text-xs text-slate-400">
                        {new Date(h.changed_at).toLocaleString()}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {/* Supporting documents */}
      <section ref={sectionRefs.documents} className="mb-10 scroll-mt-6">
        <h2 className="mb-3 text-lg font-semibold">Supporting Documents ({documents.length})</h2>
        <DocumentsPanel be_no={be_no} documents={documents} onUploaded={reload} />
      </section>
    </main>
  );
}
