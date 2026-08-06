import { supabase, API_BASE_URL } from "./supabase";
import type {
  Boe,
  BoeItem,
  BoeLicence,
  BoeDocument,
  BoeDocumentExtraction,
  BoeVariableFields,
  BoeFieldHistory,
  VariableFieldKey,
  FieldStatus,
} from "./types";

export type BoeDetail = {
  boe: Boe;
  items: BoeItem[];
  licences: BoeLicence[];
  documents: BoeDocument[];
  variableFields: BoeVariableFields | null;
  fieldHistory: BoeFieldHistory[];
};

export async function getBoeDetail(be_no: string): Promise<BoeDetail | null> {
  const [
    { data: boe },
    { data: items },
    { data: licences },
    { data: documents },
    { data: extractions },
    { data: vf },
    { data: history },
  ] = await Promise.all([
    supabase.from("boes").select("*").eq("be_no", be_no).maybeSingle(),
    supabase.from("boe_items").select("*").eq("be_no", be_no).order("global_sno"),
    supabase.from("boe_licences").select("*").eq("be_no", be_no),
    supabase.from("boe_documents").select("*").eq("be_no", be_no),
    supabase.from("boe_document_extractions").select("*").eq("be_no", be_no),
    supabase.from("boe_variable_fields").select("*").eq("be_no", be_no).maybeSingle(),
    supabase
      .from("boe_field_history")
      .select("*")
      .eq("be_no", be_no)
      .order("changed_at", { ascending: false }),
  ]);

  if (!boe) return null;

  const extractionsByPath = new Map(
    ((extractions ?? []) as BoeDocumentExtraction[]).map((e) => [e.storage_path, e])
  );
  const documentsWithExtraction = ((documents ?? []) as BoeDocument[]).map((d) => ({
    ...d,
    extraction: extractionsByPath.get(d.storage_path) ?? null,
  }));

  return {
    boe: boe as Boe,
    items: (items ?? []) as BoeItem[],
    licences: (licences ?? []) as BoeLicence[],
    documents: documentsWithExtraction,
    variableFields: (vf ?? null) as BoeVariableFields | null,
    fieldHistory: (history ?? []) as BoeFieldHistory[],
  };
}

/**
 * Updates one provisional/fixed field and logs the transition to
 * boe_field_history -- this is what lets the dashboard always show what a
 * field *was* before it became fixed, not just the latest number.
 */
export async function updateVariableField(
  be_no: string,
  field: VariableFieldKey,
  value: number,
  status: FieldStatus,
  current: BoeVariableFields | null
): Promise<void> {
  const oldValue = current?.[field] ?? null;
  const oldStatus = current?.[`${field}_status`] ?? null;

  const row: Record<string, unknown> = {
    be_no,
    [field]: value,
    [`${field}_status`]: status,
  };

  const { error: upsertError } = await supabase
    .from("boe_variable_fields")
    .upsert(row, { onConflict: "be_no" });
  if (upsertError) throw upsertError;

  const { error: historyError } = await supabase.from("boe_field_history").insert({
    be_no,
    field_name: field,
    old_value: oldValue,
    old_status: oldStatus,
    new_value: value,
    new_status: status,
  });
  if (historyError) throw historyError;
}

/**
 * Permanently deletes a BOE (rows across every table + its Storage files).
 * Goes through the backend rather than Supabase directly so the Storage
 * cleanup runs under the backend's service-role key.
 */
export async function deleteBoe(be_no: string): Promise<void> {
  const res = await fetch(`${API_BASE_URL}/boe/${be_no}`, { method: "DELETE" });
  if (!res.ok) {
    const body = await res.json().catch(() => null);
    throw new Error(body?.detail || `Failed to delete ${be_no} (${res.status})`);
  }
}
