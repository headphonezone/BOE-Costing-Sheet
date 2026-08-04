"use client";

import { useRef, useState } from "react";
import { useRouter } from "next/navigation";
import Link from "next/link";
import { API_BASE_URL } from "@/lib/supabase";

type UploadState =
  | { status: "idle" }
  | { status: "uploading"; fileName: string }
  | { status: "error"; message: string }
  | { status: "done"; be_no: string; items_saved: number; licences_saved: number };

export default function UploadPage() {
  const router = useRouter();
  const [state, setState] = useState<UploadState>({ status: "idle" });
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  async function handleFile(file: File) {
    if (file.type !== "application/pdf") {
      setState({ status: "error", message: "Please upload a PDF file." });
      return;
    }
    setState({ status: "uploading", fileName: file.name });

    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch(`${API_BASE_URL}/boe/upload`, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.detail || `Upload failed (${res.status})`);
      }
      const data = await res.json();
      setState({ status: "done", be_no: data.be_no, items_saved: data.items_saved, licences_saved: data.licences_saved });
    } catch (err) {
      setState({
        status: "error",
        message: err instanceof Error ? err.message : "Could not reach the backend. Is it running?",
      });
    }
  }

  return (
    <main className="mx-auto max-w-2xl px-6 py-10">
      <Link href="/" className="text-sm text-blue-600 hover:underline">
        ← All BOEs
      </Link>

      <h1 className="mt-3 mb-2 text-2xl font-bold">Upload a BOE</h1>
      <p className="mb-8 text-sm text-slate-500">
        Upload a Bill of Entry PDF. It&apos;ll be parsed and every item, duty, and licence saved to the
        database automatically.
      </p>

      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragOver(false);
          const file = e.dataTransfer.files?.[0];
          if (file) handleFile(file);
        }}
        onClick={() => inputRef.current?.click()}
        className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-6 py-16 text-center transition ${
          dragOver
            ? "border-blue-500 bg-blue-50 dark:bg-blue-950"
            : "border-slate-300 bg-white hover:border-blue-400 dark:border-slate-700 dark:bg-slate-900"
        }`}
      >
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) handleFile(file);
          }}
        />

        {state.status === "uploading" ? (
          <>
            <div className="mb-2 h-8 w-8 animate-spin rounded-full border-2 border-blue-500 border-t-transparent" />
            <p className="text-sm font-medium">Parsing {state.fileName}…</p>
            <p className="mt-1 text-xs text-slate-400">This can take a few seconds for large BOEs</p>
          </>
        ) : (
          <>
            <p className="text-sm font-medium">Drag &amp; drop a BOE PDF here, or click to browse</p>
            <p className="mt-1 text-xs text-slate-400">Only .pdf files are accepted</p>
          </>
        )}
      </div>

      {state.status === "error" && (
        <div className="mt-6 rounded-lg border border-red-300 bg-red-50 px-4 py-3 text-sm text-red-700 dark:border-red-900 dark:bg-red-950 dark:text-red-300">
          {state.message}
          <button
            onClick={() => setState({ status: "idle" })}
            className="ml-3 font-medium underline underline-offset-2"
          >
            Try again
          </button>
        </div>
      )}

      {state.status === "done" && (
        <div className="mt-6 rounded-lg border border-emerald-300 bg-emerald-50 px-4 py-4 text-sm text-emerald-800 dark:border-emerald-800 dark:bg-emerald-950 dark:text-emerald-200">
          <p className="font-medium">
            Saved BE No {state.be_no} — {state.items_saved} items, {state.licences_saved} licence rows.
          </p>
          <div className="mt-3 flex gap-3">
            <button
              onClick={() => router.push(`/boe/${state.be_no}`)}
              className="rounded-lg bg-emerald-600 px-3 py-1.5 font-medium text-white hover:bg-emerald-700"
            >
              Open Dashboard
            </button>
            <button
              onClick={() => setState({ status: "idle" })}
              className="rounded-lg border border-emerald-300 px-3 py-1.5 font-medium text-emerald-700 hover:bg-emerald-100 dark:border-emerald-700 dark:text-emerald-300"
            >
              Upload Another
            </button>
          </div>
        </div>
      )}
    </main>
  );
}
