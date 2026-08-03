"""
Supabase persistence for the per-BOE variable/provisional cost fields
(Exchange Rate, Freight Charges, Clearing Charges, Supplier Freight,
Bank Charges, Own Bank Charges).

Each field has a value and a status ('provisional' or 'fixed'). One row
per BE No. Credentials come from Streamlit secrets, never hardcoded here.
"""
import streamlit as st
from supabase import create_client

TABLE = "boe_variable_fields"

FIELDS = [
    "exchange_rate",
    "freight_charges",
    "clearing_charges",
    "supplier_freight",
    "bank_charges",
    "own_bank_charges",
]


@st.cache_resource
def _client():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


def get_boe_record(be_no: str) -> dict | None:
    """Returns the saved row for this BE No, or None if never saved."""
    if not be_no:
        return None
    resp = _client().table(TABLE).select("*").eq("be_no", be_no).limit(1).execute()
    rows = resp.data or []
    return rows[0] if rows else None


def upsert_boe_record(be_no: str, values: dict) -> None:
    """
    values: {field_name: {'value': float, 'status': 'provisional'|'fixed'}}
    for any subset of FIELDS.
    """
    if not be_no:
        return
    row = {"be_no": be_no}
    for field in FIELDS:
        if field in values:
            row[field] = values[field]["value"]
            row[f"{field}_status"] = values[field]["status"]
    _client().table(TABLE).upsert(row, on_conflict="be_no").execute()
