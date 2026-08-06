export type Boe = {
  be_no: string;
  be_date: string | null;
  port_code: string | null;
  importer_name: string | null;
  supplier_name: string | null;
  inv_no: string | null;
  inv_date: string | null;
  inv_value_usd: number | null;
  freight_inr: number | null;
  insurance_inr: number | null;
  misc_charges_inr: number | null;
  exchange_rate: number | null;
  hawb_no: string | null;
  total_assess_value: number | null;
  total_duty: number | null;
  created_at: string;
  updated_at: string;
};

export type BoeItem = {
  id: number;
  be_no: string;
  global_sno: number | null;
  invsno: number;
  itemsn: number;
  cth: string | null;
  description: string | null;
  unit_price_usd: number | null;
  qty: number | null;
  uqc: string | null;
  assess_value: number | null;
  bcd: number | null;
  bcd_forgone: number | null;
  sws: number | null;
  igst: number | null;
  total_duty: number | null;
};

export type BoeLicence = {
  id: number;
  be_no: string;
  invsno: number;
  itemsn: number;
  lic_no: string | null;
  lic_date: string | null;
  code: string | null;
  port: string | null;
  debit_value: number | null;
  debit_duty: number | null;
};

export type BoeDocument = {
  id: number;
  be_no: string;
  doc_type: string | null;
  file_name: string | null;
  storage_path: string;
  uploaded_at: string;
  extraction?: BoeDocumentExtraction | null;
};

export type BoeDocumentLineItem = {
  description: string;
  qty: number;
  unit_price: number;
  amount: number;
};

export type BoeDocumentExtraction = {
  id: number;
  be_no: string;
  storage_path: string;
  doc_type: string;
  doc_number: string | null;
  doc_date: string | null;
  total_value: number | null;
  currency: string | null;
  total_packages: number | null;
  gross_weight_kg: number | null;
  net_weight_kg: number | null;
  certificate_no: string | null;
  origin_country: string | null;
  line_items: BoeDocumentLineItem[] | null;
  raw_text: string | null;
  extracted_at: string;
};

export const VARIABLE_FIELDS = [
  "exchange_rate",
  "freight_charges",
  "clearing_charges",
  "supplier_freight",
  "bank_charges",
  "own_bank_charges",
] as const;

export type VariableFieldKey = (typeof VARIABLE_FIELDS)[number];
export type FieldStatus = "provisional" | "fixed";

export type BoeVariableFields = {
  be_no: string;
  updated_at: string;
} & {
  [K in VariableFieldKey]: number | null;
} & {
  [K in `${VariableFieldKey}_status`]: FieldStatus | null;
};

export type BoeFieldHistory = {
  id: number;
  be_no: string;
  field_name: VariableFieldKey;
  old_value: number | null;
  old_status: FieldStatus | null;
  new_value: number | null;
  new_status: FieldStatus | null;
  changed_at: string;
};

export const VARIABLE_FIELD_LABELS: Record<VariableFieldKey, string> = {
  exchange_rate: "Exchange Rate",
  freight_charges: "Freight Charges",
  clearing_charges: "Clearing Charges",
  supplier_freight: "Supplier Freight",
  bank_charges: "Bank Charges",
  own_bank_charges: "Own Bank Charges",
};
