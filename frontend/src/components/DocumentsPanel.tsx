"use client";

import { Fragment, useEffect, useRef, useState } from "react";
import { API_BASE_URL, supabase } from "@/lib/supabase";
import type { BoeDocument } from "@/lib/types";

function fmtNum(n: number | null | undefined) {
  if (n === null || n === undefined) return "—";
  return n.toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

function extractionSummary(doc: BoeDocument): string {
  const e = doc.extraction;
  if (!e) return doc.doc_type === "BOE" ? "" : "No data extracted";
  const parts: string[] = [];
  if (e.doc_number) parts.push(e.doc_number);
  if (e.total_value !== null) parts.push(`${e.currency ?? ""} ${fmtNum(e.total_value)}`.trim());
  if (e.total_packages !== null) parts.push(`${e.total_packages} pkgs`);
  if (e.gross_weight_kg !== null) parts.push(`${fmtNum(e.gross_weight_kg)} kg gross`);
  if (e.certificate_no) parts.push(e.certificate_no);
  if (e.origin_country) parts.push(e.origin_country);
  return parts.length > 0 ? parts.join(" · ") : "No structured fields found";
}

const DOC_TYPES = ["BOE", "INVOICE", "PACKING_LIST", "COO", "OTHER"] as const;
const DOCS_BUCKET = "boe-documents";

export function DocumentsPanel({
  be_no,
  documents,
  onUploaded,
}: {
  be_no: string;
  documents: BoeDocument[];
  onUploaded: () => void;
}) {
  const [docType, setDocType] = useState<(typeof DOC_TYPES)[number]>("INVOICE");
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [expanded, setExpanded] = useState<number | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // The boe-documents bucket is private, so links need signed URLs (not
  // getPublicUrl, which only works for public buckets and would 403 here).
  const [signedUrls, setSignedUrls] = useState<Record<string, string>>({});

  useEffect(() => {
    let cancelled = false;
    async function loadSignedUrls() {
      const entries = await Promise.all(
        documents.map(async (doc) => {
          const { data } = await supabase.storage
            .from(DOCS_BUCKET)
            .createSignedUrl(doc.storage_path, 60 * 60); // 1 hour
          return [doc.storage_path, data?.signedUrl ?? ""] as const;
        })
      );
      if (!cancelled) setSignedUrls(Object.fromEntries(entries));
    }
    if (documents.length > 0) loadSignedUrls();
    return () => {
      cancelled = true;
    };
  }, [documents]);

  async function handleFile(file: File) {
    setUploading(true);
    setError(null);
    try {
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch(`${API_BASE_URL}/boe/${be_no}/documents?doc_type=${docType}`, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.detail || `Upload failed (${res.status})`);
      }
      onUploaded();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not reach the backend. Is it running?");
    } finally {
      setUploading(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  function fileUrl(storage_path: string) {
    return signedUrls[storage_path];
  }

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center gap-2">
        <select
          value={docType}
          onChange={(e) => setDocType(e.target.value as (typeof DOC_TYPES)[number])}
          className="rounded-md border border-slate-300 bg-white px-2 py-1.5 text-sm dark:border-slate-700 dark:bg-slate-900"
        >
          {DOC_TYPES.map((t) => (
            <option key={t} value={t}>
              {t.replace("_", " ")}
            </option>
          ))}
        </select>
        <input
          ref={inputRef}
          type="file"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) handleFile(file);
          }}
          disabled={uploading}
          className="text-sm file:mr-3 file:rounded-md file:border-0 file:bg-blue-600 file:px-3 file:py-1.5 file:text-sm file:font-medium file:text-white hover:file:bg-blue-700 disabled:opacity-50"
        />
        {uploading && <span className="text-xs text-slate-400">Uploading…</span>}
      </div>

      {error && (
        <div className="mb-4 rounded-lg border border-red-300 bg-red-50 px-3 py-2 text-xs text-red-700 dark:border-red-900 dark:bg-red-950 dark:text-red-300">
          {error}
        </div>
      )}

      {documents.length === 0 ? (
        <p className="text-sm text-slate-500">No supporting documents uploaded yet.</p>
      ) : (
        <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800">
          <table className="w-full text-sm">
            <thead className="bg-slate-100 text-left text-xs uppercase tracking-wide text-slate-500 dark:bg-slate-800">
              <tr>
                <th className="px-3 py-2">Type</th>
                <th className="px-3 py-2">File</th>
                <th className="px-3 py-2">Extracted Data</th>
                <th className="px-3 py-2">Uploaded</th>
              </tr>
            </thead>
            <tbody>
              {documents.map((doc) => {
                const hasDetail = Boolean(
                  doc.extraction && (doc.extraction.line_items?.length || doc.extraction.raw_text)
                );
                const isOpen = expanded === doc.id;
                return (
                  <Fragment key={doc.id}>
                    <tr
                      className={`border-t border-slate-100 dark:border-slate-800 ${
                        hasDetail ? "cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/50" : ""
                      }`}
                      onClick={() => hasDetail && setExpanded(isOpen ? null : doc.id)}
                    >
                      <td className="px-3 py-2">
                        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600 dark:bg-slate-800 dark:text-slate-300">
                          {doc.doc_type}
                        </span>
                      </td>
                      <td className="px-3 py-2">
                        <a
                          href={fileUrl(doc.storage_path) || undefined}
                          target="_blank"
                          rel="noopener noreferrer"
                          onClick={(e) => e.stopPropagation()}
                          className={
                            fileUrl(doc.storage_path)
                              ? "text-blue-600 hover:underline"
                              : "cursor-default text-slate-400"
                          }
                        >
                          {doc.file_name}
                        </a>
                      </td>
                      <td className="px-3 py-2 text-xs text-slate-500">
                        {extractionSummary(doc)}
                        {hasDetail && (
                          <span className="ml-1 text-slate-400">{isOpen ? "▲" : "▼"}</span>
                        )}
                      </td>
                      <td className="px-3 py-2 text-xs text-slate-400">
                        {new Date(doc.uploaded_at).toLocaleString()}
                      </td>
                    </tr>
                    {isOpen && doc.extraction && (
                      <tr className="border-t border-slate-100 bg-slate-50 dark:border-slate-800 dark:bg-slate-800/30">
                        <td colSpan={4} className="px-3 py-3">
                          {doc.extraction.line_items && doc.extraction.line_items.length > 0 && (
                            <div className="mb-3 overflow-x-auto">
                              <table className="w-full text-xs">
                                <thead className="text-left text-slate-400">
                                  <tr>
                                    <th className="pr-3 py-1">Description</th>
                                    <th className="pr-3 py-1 text-right">Qty</th>
                                    <th className="pr-3 py-1 text-right">Unit Price</th>
                                    <th className="pr-3 py-1 text-right">Amount</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {doc.extraction.line_items.map((li, i) => (
                                    <tr key={i} className="border-t border-slate-200 dark:border-slate-700">
                                      <td className="pr-3 py-1">{li.description}</td>
                                      <td className="pr-3 py-1 text-right tabular-nums">{fmtNum(li.qty)}</td>
                                      <td className="pr-3 py-1 text-right tabular-nums">{fmtNum(li.unit_price)}</td>
                                      <td className="pr-3 py-1 text-right tabular-nums">{fmtNum(li.amount)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                          {doc.extraction.raw_text && (
                            <details>
                              <summary className="cursor-pointer text-xs text-slate-500">
                                Raw extracted text
                              </summary>
                              <pre className="mt-2 max-h-64 overflow-auto whitespace-pre-wrap text-xs text-slate-500">
                                {doc.extraction.raw_text}
                              </pre>
                            </details>
                          )}
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
