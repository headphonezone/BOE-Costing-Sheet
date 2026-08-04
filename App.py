"""
Streamlit UI for the BOE -> Excel filler. All parsing/Excel-generation
logic lives in boe_parser.py (shared with the FastAPI backend); this file
is just the interactive front-end.
"""
import io

import pdfplumber
import streamlit as st

import db
from boe_parser import (
    parse_header, parse_hawb, parse_all_items, parse_invoice_summary_multi,
    parse_licences_from_page, extract_assess_values_from_pages,
    parse_be_no_from_pages, extract_duties_from_page, parse_scheme_g,
    fill_excel,
)

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="BOE → Excel", page_icon="📦", layout="wide")

st.markdown("""
<style>
.main-title{font-size:2rem;font-weight:700;color:#4A90D9;margin-bottom:0}
.sub-title{color:#aaa;margin-top:2px;margin-bottom:1rem}
.info-box{
    background:#1e2d3d !important;
    border-left:4px solid #4A90D9 !important;
    padding:14px 18px !important;
    border-radius:6px !important;
    margin:8px 0 !important;
    color:#e8f0fe !important;
}
.info-box b{color:#7ec8e3 !important;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📦 BOE → Excel Auto-Filler</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Ferrari Video | Focal Shipments — Upload BOE PDF → Download Filled Excel</div>', unsafe_allow_html=True)

st.markdown("""<div class="info-box">
<b>How it works:</b> Upload any Focal BOE PDF → extracts all item details (across all
invoices on the BOE), duties, and Part IV Scheme/Licence details → fills Excel
template → download ready.<br><br>
<b>BCD Rule:</b> BCD comes from Part IV Section G (Scheme Notification & Duty
Foregone Details) when the item's duty was paid via licence, which is already
correctly summed even when an item draws from multiple licences.
</div>""", unsafe_allow_html=True)

st.divider()

pdf_file = st.file_uploader("📄 Upload Bill of Entry PDF", type=["pdf"],
                             help="Upload the BOE PDF from ICEGATE")

if pdf_file:
    pdf_bytes = pdf_file.read()

    with st.spinner("🔍 Parsing PDF..."):
        try:
            pages_text = []
            all_duties = {}
            all_bcd_forgone = {}

            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                total_pages = len(pdf.pages)

                for pg in pdf.pages:
                    d = extract_duties_from_page(pg)
                    if d:
                        all_duties.update(d)

                for pg in pdf.pages:
                    all_bcd_forgone.update(parse_scheme_g(pg))

                for p in pdf.pages:
                    pages_text.append(p.extract_text() or "")

            header = parse_header(pages_text[0])
            ex_rate = header.get('exchange_rate', 1.0)
            header['hawb_no'] = parse_hawb(pages_text[0])

            # multi-invoice-aware item + meta parsing (Bug 1 & 2 fixes)
            meta, items = parse_all_items(pages_text[1:], ex_rate)

            # cross-check header invoice no/value against Part I summary table
            inv_summary_list = parse_invoice_summary_multi(pages_text[0])
            if inv_summary_list and not meta.get('inv_no'):
                meta['inv_no'] = inv_summary_list[0]['inv_no']
                meta['inv_value'] = inv_summary_list[0]['inv_value']

            licences = []
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for pg in pdf.pages:
                    lics = parse_licences_from_page(pg)
                    if lics:
                        licences.extend(lics)

            seen = set()
            unique_lics = []
            for lic in licences:
                key = (lic['invsno'], lic['itmsno'], lic['lic_no'], lic['debit_duty'])
                if key not in seen:
                    seen.add(key)
                    unique_lics.append(lic)
            licences = unique_lics

            assess_values = {}
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                part3_pages = []
                in_p3 = False
                for pg in pdf.pages:
                    pg_text = pg.extract_text() or ''
                    if 'PART - III' in pg_text:
                        in_p3 = True
                    if in_p3:
                        part3_pages.append(pg)
                    if in_p3 and 'PART - IV' in pg_text:
                        break
                assess_values = extract_assess_values_from_pages(part3_pages)

            be_no_all = parse_be_no_from_pages(pages_text)
            all_duties['_be_no'] = be_no_all or header.get('be_no', '')
            all_duties['_be_date'] = header.get('be_date', '')

        except Exception as e:
            st.error(f"❌ PDF parsing error: {e}")
            st.exception(e)
            st.stop()

    st.success(f"✅ Parsed — {total_pages} pages · {len(items)} items · {len(licences)} licence rows")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Exchange Rate", f"₹ {header.get('exchange_rate', '—')}")
    c2.metric("Invoice No", meta.get('inv_no', '—'))
    c3.metric("Invoice Date", meta.get('inv_date', '—'))
    c4.metric("Freight (₹)", f"{meta.get('freight', '—'):,}" if meta.get('freight') else '—')
    c5.metric("Insurance (₹)", f"{meta.get('insurance', '—'):,}" if meta.get('insurance') else '—')

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("Supplier", meta.get('supplier', '—')[:25] if meta.get('supplier') else '—')
    c7.metric("AWB / HAWB No", header.get('hawb_no', '—'))
    c8.metric("BE No", header.get('be_no', '—'))
    c9.metric("Inv Value (USD)", f"{meta.get('inv_value', '—'):,}" if meta.get('inv_value') else '—')
    c10.metric("BE Date", header.get('be_date', '—'))

    # ── Variable / Provisional Cost Items ───────────────────────────────────
    # These six fields are editable in the interface, persisted per BE No in
    # Supabase, and re-loaded automatically the next time this same BOE is
    # opened. Each carries a Provisional (yellow) / Fixed (green) status that
    # is mirrored as the cell fill color in the generated Excel.
    be_no_key = all_duties.get('_be_no', '') or header.get('be_no', '')

    st.divider()
    st.subheader("🛠️ Variable / Provisional Cost Items")
    st.caption(
        "Editable per BOE (keyed by BE No) and saved to Supabase. "
        "Yellow = Provisional, Green = Fixed — same colors are applied in the Excel."
    )

    if not be_no_key:
        st.warning("⚠️ No BE No found on this PDF — variable items can't be saved/reloaded without it.")

    try:
        saved_record = db.get_boe_record(be_no_key) if be_no_key else None
    except Exception as e:
        saved_record = None
        st.error(f"Could not reach Supabase: {e}")
    saved_record = saved_record or {}

    field_defs = [
        ("exchange_rate", "Exchange Rate", header.get('exchange_rate', 0.0)),
        ("freight_charges", "Freight Charges", meta.get('freight', 0.0)),
        ("clearing_charges", "Clearing Charges", 0.0),
        ("supplier_freight", "Supplier Freight", 0.0),
        ("bank_charges", "Bank Charges", 0.0),
        ("own_bank_charges", "Own Bank Charges", 0.0),
    ]

    variable_fields = {}
    field_cols = st.columns(3)
    for i, (fkey, flabel, fdefault) in enumerate(field_defs):
        saved_val = saved_record.get(fkey)
        saved_status = saved_record.get(f"{fkey}_status") or "provisional"
        starting_value = float(saved_val) if saved_val is not None else float(fdefault or 0.0)

        status_widget_key = f"status_{fkey}_{be_no_key}"
        value_widget_key = f"val_{fkey}_{be_no_key}"
        # reflects the *current* radio selection (post-rerun), not just the
        # saved one, so the badge updates the moment you flip the toggle
        live_status = st.session_state.get(status_widget_key, saved_status)
        badge_color = "#92D050" if live_status == "fixed" else "#FFFF00"

        with field_cols[i % 3]:
            st.markdown(
                f"<div style='background-color:{badge_color}; padding:3px 10px; "
                f"border-radius:4px; font-weight:700; display:inline-block; "
                f"margin-bottom:2px; color:#000;'>{flabel.upper()}</div>",
                unsafe_allow_html=True,
            )
            value = st.number_input(
                flabel, value=starting_value, key=value_widget_key,
                label_visibility="collapsed",
            )
            status = st.radio(
                f"{flabel} status", ["provisional", "fixed"],
                index=0 if saved_status == "provisional" else 1,
                key=status_widget_key, horizontal=True,
                label_visibility="collapsed",
            )
            variable_fields[fkey] = {"value": value, "status": status}

    if be_no_key:
        try:
            db.upsert_boe_record(be_no_key, variable_fields)
        except Exception as e:
            st.error(f"⚠️ Could not save variable items to Supabase: {e}")

    import pandas as pd

    tab1, tab2, tab3, tab4 = st.tabs([
        f"📦 Items ({len(items)})",
        f"🏦 Duties ({len(all_duties) - 2})",
        f"📜 Licences ({len(licences)})",
        "🔍 BCD Source",
    ])

    with tab1:
        if items:
            df = pd.DataFrame(items)
            df = df[['global_sno', 'invsno', 'itemsn', 'desc', 'price', 'qty']]
            df.columns = ['Row #', 'Invoice', 'Item # (in invoice)', 'Description', 'Unit Price (USD)', 'Qty']
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No items extracted")

    with tab2:
        duty_rows = [
            {'Invoice': k[0], 'Item': k[1], 'BCD Paid': v.get('bcd'), 'SWS': v.get('sws'), 'IGST': v.get('igst')}
            for k, v in sorted((i for i in all_duties.items() if isinstance(i[0], tuple)), key=lambda x: (x[0][0], x[0][1]))
        ]
        if duty_rows:
            st.dataframe(pd.DataFrame(duty_rows), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No duties extracted")

    with tab3:
        if licences:
            st.caption("Reference licence numbers from F. LICENCE DETAILS.")
            df_lic = pd.DataFrame(licences)
            st.dataframe(df_lic, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ No licence entries found")

    with tab4:
        bcd_rows = []
        for it in items:
            key = (it['invsno'], it['itemsn'])
            d = all_duties.get(key, {})
            bcd_paid = d.get('bcd', 0)
            bcd_fg = all_bcd_forgone.get(key)

            if bcd_paid == 0 and bcd_fg:
                bcd_written = round(bcd_fg, 2)
                source = '📜 SCHEME G (duty forgone)'
            elif bcd_paid > 0:
                bcd_written = bcd_paid
                source = '💵 CASH / CYBER'
            else:
                bcd_written = 0
                source = '⚪ NIL'

            bcd_rows.append({
                'Row #': it['global_sno'],
                'Invoice-Item': f"{it['invsno']}-{it['itemsn']}",
                'Description': it['desc'][:45],
                'Source': source,
                'BCD Written': bcd_written,
            })

        st.caption("Preview of what will be written into BCD column in Excel")
        st.dataframe(pd.DataFrame(bcd_rows), use_container_width=True, hide_index=True)

    st.divider()

    with st.spinner("✍️ Filling Excel..."):
        try:
            filled_bytes = fill_excel(header, meta, items, all_duties, all_bcd_forgone, licences,
                                       assess_values, variable_fields)
            filename = f"BOE_{meta.get('inv_no', 'filled')}.xlsx"
            if filled_bytes is None:
                st.error("❌ fill_excel returned None — check _fill_c_sheet for errors")
                st.stop()
        except Exception as e:
            st.error(f"❌ Excel filling error: {e}")
            st.exception(e)
            st.stop()

    st.download_button(
        label="📥 Download Filled Excel",
        data=filled_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    st.markdown("""<div class="info-box">
    <b>What was filled:</b><br>
    ✅ <b>C-SHEET</b>: Supplier · Exchange Rate · Invoice No · Invoice Value (USD) ·
    Invoice Date · AWB/HAWB No · Freight · Insurance · All item/duty/cost formulas<br>
    ✅ <b>D-DETAILS</b>: BOE No · BCD from Part IV Scheme/Duty-Forgone details ·
    SWS · IGST · Paid/Cyber Receipt · Licence reference table · All totals<br><br>
    <b>Note:</b> This BOE may contain multiple invoices. All items across every
    invoice are numbered sequentially in the Excel (Row # column); the header
    fields (supplier / invoice no / freight / insurance) reflect the invoice
    with the most line items. If you need separate costing per invoice,
    let me know and I can split this into one sheet per invoice instead.
    </div>""", unsafe_allow_html=True)

else:
    st.markdown("""<div class="info-box">
    👆 Upload the BOE PDF above. Excel template is embedded.
    </div>""", unsafe_allow_html=True)

st.divider()
st.caption("Ferrari Video | BOE Automation | Streamlit · pdfplumber · openpyxl")
