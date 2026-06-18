# STREAMING_CHUNK: Initializing Streamlit application view and workspace interfaces...
import streamlit as st
import pandas as pd
import os
import base64
import time
import json
import threading

# Import local custom feature modules
from config import QUEUE_FILE, APPROVED_FILE
from database import init_excel_files, safe_read_excel, safe_approve_po, read_worker_status
from mail_worker import continuous_mail_sync_loop

init_excel_files()

# Initialize background pipeline worker single instance
if "background_sync_started" not in st.session_state:
    st.session_state["background_sync_started"] = True
    bg_thread = threading.Thread(target=continuous_mail_sync_loop, daemon=True)
    bg_thread.start()

# --- SIDEBAR REGION ---
with st.sidebar:
    st.title("MANE KANCOR")
    st.caption("Procurement Automation Desk")
    st.divider()

    workspace = st.radio(
        "Choose Active Workspace Desk:", 
        ["📊 Operational Dashboard", "📁 PO Section (Pending)", "✅ Approved PO Registry", "🔍 Verification Desk"]
    )
    st.divider()
    
    @st.fragment(run_every=4)
    def render_sidebar_status_widget():
        st.subheader("📻 Sync Engine Status")
        st.markdown(f"**Status:** {read_worker_status()}")
        
        df_live_q = safe_read_excel(QUEUE_FILE)
        df_live_a = safe_read_excel(APPROVED_FILE)
        st.write(f"Pending Inbox Queue Rows: **{len(df_live_q)}**")
        st.write(f"Cleared Master Ledger Rows: **{len(df_live_a)}**")
        
    render_sidebar_status_widget()

    if st.button("🔄 Refresh Application UI Manually", use_container_width=True):
        st.rerun()

# --- WORKSPACE A: OPERATIONAL DASHBOARD ---
if workspace == "📊 Operational Dashboard":
    st.title("📊 Operational Management Insights")
    st.write("Real-time summary tracking structural supply document ingestion workflows.")
    st.divider()
    
    @st.fragment(run_every=4)
    def render_live_dashboard_view():
        df_q = safe_read_excel(QUEUE_FILE)
        df_a = safe_read_excel(APPROVED_FILE)
        
        col1, col2 = st.columns(2)
        with col1: st.metric(label="Awaiting Audit (in purchase_orders.xlsx)", value=len(df_q))
        with col2: st.metric(label="Successfully Verified (in approved_orders.xlsx)", value=len(df_a))
            
        st.subheader("🕒 Live Incoming Mail Feed (Unapproved Rows Only)")
        if not df_q.empty:
            refined_df = df_q[["PO ID", "Company Name", "System Timestamp", "Approval Status"]].copy()
            refined_df.columns = ["Order ID", "Company Name", "Ingestion Inbound Time", "Status"]
            st.dataframe(refined_df.sort_values(by="Ingestion Inbound Time", ascending=False), use_container_width=True, hide_index=True)
        else:
            st.info("No temporary records are active inside your incoming email workspace.")
            
    render_live_dashboard_view()

# --- WORKSPACE B: PO SECTION (PENDING ONLY) ---
elif workspace == "📁 PO Section (Pending)":
    st.title("📁 Unverified Purchase Order Registry (Queue File)")
    st.write("Lightweight view of outstanding pipeline extractions waiting on verification desks.")
    st.divider()
    
    @st.fragment(run_every=4)
    def render_live_pending_queue():
        df_q = safe_read_excel(QUEUE_FILE)
        if not df_q.empty:
            clean_view_df = df_q[["PO ID", "Company Name", "System Timestamp"]].copy()
            clean_view_df.columns = ["Order ID", "Company Name", "Logged Time"]
            st.dataframe(clean_view_df, use_container_width=True, hide_index=True)
        else:
            st.success("🎉 Ingestion queue clear! purchase_orders.xlsx is empty.")
            
    render_live_pending_queue()

# --- WORKSPACE C: APPROVED LOG ARCHIVE ---
elif workspace == "✅ Approved PO Registry":
    st.title("✅ Permanent Cleared Production Ledger (approved_orders.xlsx)")
    st.write("Historical lookup workstation for purchase orders successfully validated and committed.")
    st.divider()
    
    df_approved = safe_read_excel(APPROVED_FILE)
    if df_approved.empty:
        st.info("No records have cleared historic confirmation gates yet.")
        st.session_state["selected_approved_po"] = None
    else:
        unique_regions = ["All Operational Regions"] + list(df_approved["Region"].dropna().unique())
        selected_region = st.selectbox("🎯 Filter Ledger via Destination Region:", unique_regions)
        filtered_df = df_approved if selected_region == "All Operational Regions" else df_approved[df_approved["Region"] == selected_region]
        
        if filtered_df.empty:
            st.warning("No cleared purchase orders map to this regional partition.")
        else:
            st.write("### 📂 Active Cleared Ledgers")
            cols = st.columns(3)
            for idx, (_, row) in enumerate(filtered_df.sort_values(by="System Timestamp", ascending=False).iterrows()):
                with cols[idx % 3]:
                    with st.container(border=True):
                        st.subheader(f"✅ {row['PO ID']}")
                        st.write(f"**🏢 Company:** {row['Company Name']}")
                        st.write(f"**💰 Total:** {row['Total Amount']}")
                        st.write(f"**📍 Destination:** {row['Region']}")
                        
                        if st.button(f"📄 Open File Archives", key=f"view_{row['PO ID']}", use_container_width=True):
                            st.session_state["selected_approved_po"] = row["PO ID"]
                            st.rerun()

            if st.session_state.get("selected_approved_po"):
                target_po_id = st.session_state["selected_approved_po"]
                matched_rows = df_approved[df_approved["PO ID"] == target_po_id]
                
                if not matched_rows.empty:
                    selected_row = matched_rows.iloc[0]
                    st.divider()
                    st.subheader(f"🛠️ Immutable Audit Panel Summary: {selected_row['PO ID']}")
                    
                    panel_left, panel_right = st.columns([1, 1])
                    with panel_left:
                        st.write("##### 📄 Original Source Attachment Payload")
                        pdf_target_path = selected_row["PDF Path"]
                        if os.path.exists(pdf_target_path):
                            try:
                                with open(pdf_target_path, "rb") as f:
                                    base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700" type="application/pdf"></iframe>'
                                st.markdown(pdf_display, unsafe_allow_html=True)
                            except Exception as render_ex:
                                st.error(f"Render Error Engine: {render_ex}")
                        else:
                            st.error("⚠️ Attachment missing from system storage pathways.")
                            
                    with panel_right:
                        st.write("##### 🔒 Committed Core Parameter Records")
                        st.text_input("Company Reference", value=str(selected_row["Company Name"]), disabled=True, key="view_comp")
                        st.text_input("Order Date Doc-String", value=str(selected_row["Order Date"]), disabled=True, key="view_dt")
                        st.text_input("Total Valued Amount", value=str(selected_row["Total Amount"]), disabled=True, key="view_amt")
                        st.text_area("Delivery Address Target Block", value=str(selected_row["Address to be Delivered"]), height=70, disabled=True, key="view_addr")
                        st.text_input("Inferred Region Cluster Zone", value=str(selected_row["Region"]), disabled=True, key="view_reg")
                        
                        st.write("##### 📦 Closed Item Breakdown")
                        try: parsed_products = json.loads(selected_row["Products JSON"])
                        except: parsed_products = {"Unknown Item Summary": "None"}
                            
                        for prod_item, prod_qty in parsed_products.items():
                            col_p1, col_p2 = st.columns([2, 1])
                            with col_p1: st.text_input("Item Description Label", value=str(prod_item), disabled=True, key=f"view_pn_{prod_item}_{selected_row['PO ID']}")
                            with col_p2: st.text_input("Committed Quantities", value=str(prod_qty), disabled=True, key=f"view_pq_{prod_item}_{selected_row['PO ID']}")

# --- WORKSPACE D: INTERACTIVE AUDIT DESK ---
elif workspace == "🔍 Verification Desk":
    st.title("🔍 Selection Matrix & Split-Screen Verification Desk")
    st.write("Review unverified extraction models side-by-side with original file attachment feeds.")
    st.divider()
    
    df_q = safe_read_excel(QUEUE_FILE)
    
    if df_q.empty:
        st.success("🎉 Verification desk cleared! No pending items found.")
        st.session_state["selected_verification_po"] = None
    else:
        if "selected_verification_po" in st.session_state and st.session_state["selected_verification_po"] is not None:
            target_id = st.session_state["selected_verification_po"]
            matching_rows = df_q[df_q["PO ID"] == target_id]
            
            if matching_rows.empty:
                st.warning("Selected purchase order was not found. It might have been approved or modified.")
                st.session_state["selected_verification_po"] = None
                st.rerun()
            else:
                po_data = matching_rows.iloc[0]
                
                col_back_btn, col_title_lbl = st.columns([1, 4])
                with col_back_btn:
                    if st.button("⬅️ Back to Selection list", use_container_width=True):
                        st.session_state["selected_verification_po"] = None
                        st.rerun()
                with col_title_lbl:
                    st.subheader(f"📝 Active Verification Workspace: {target_id}")
                
                st.divider()
                col_pdf, col_form = st.columns([1, 1])
                
                with col_pdf:
                    st.markdown("##### 📄 Original Purchase Order PDF Preview")
                    pdf_target_path = po_data["PDF Path"]
                    if os.path.exists(pdf_target_path):
                        try:
                            with open(pdf_target_path, "rb") as f:
                                base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="750" type="application/pdf"></iframe>'
                            st.markdown(pdf_display, unsafe_allow_html=True)
                        except Exception as render_ex:
                            st.error(f"Render Error Engine: {render_ex}")
                    else:
                        st.error("⚠️ Attachment payload missing from server storage pathways.")
                
                with col_form:
                    st.markdown("##### ✏️ Modifiable Parameters Form")
                    edit_company = st.text_input("Company Name", value=str(po_data["Company Name"]), key="form_company")
                    edit_date = st.text_input("Order Date Reference", value=str(po_data["Order Date"]), key="form_date")
                    edit_amount = st.text_input("Total Order Amount", value=str(po_data["Total Amount"]), key="form_amount")
                    edit_region = st.text_input("Assigned Region", value=str(po_data["Region"]), key="form_region")
                    edit_address = st.text_area("Delivery Address Target", value=str(po_data["Address to be Delivered"]), height=70, key="form_address")
                    
                    st.write(f"**📬 Email Source Sender:** `{po_data['Email Sender']}`")
                    st.write(f"**📧 Subject Line:** {po_data['Email Subject']}")
                    
                    st.write("##### 📦 Itemized Full Manifest Catalog")
                    try: parsed_products = json.loads(po_data["Products JSON"])
                    except: parsed_products = {"Unknown Item": "None"}
                        
                    updated_products_dict = {}
                    for prod_item, prod_qty in parsed_products.items():
                        col_p1, col_p2 = st.columns([2, 1])
                        with col_p1: new_name = st.text_input(f"Product Detail", value=str(prod_item), key=f"name_{prod_item}")
                        with col_p2: new_qty = st.text_input(f"Qty Metrics", value=str(prod_qty), key=f"qty_{prod_item}")
                        if new_name: updated_products_dict[new_name] = new_qty
                            
                    if "extra_items_count" not in st.session_state:
                        st.session_state["extra_items_count"] = 0
                        
                    if st.button("➕ Append New Custom Item Row", use_container_width=True):
                        st.session_state["extra_items_count"] += 1
                        st.rerun()
                        
                    for i in range(st.session_state["extra_items_count"]):
                        col_e1, col_e2 = st.columns([2, 1])
                        with col_e1: e_name = st.text_input("New Item Description", key=f"extra_name_{i}")
                        with col_e2: e_qty = st.text_input("New Item Quantity", key=f"extra_qty_{i}")
                        if e_name: updated_products_dict[e_name] = e_qty
                    
                    st.divider()
                    if st.button("✅ Complete Audit & Move to approved_orders.xlsx", type="primary", use_container_width=True):
                        updated_fields = {
                            "Company Name": edit_company, "Order Date": edit_date, "Total Amount": edit_amount,
                            "Region": edit_region, "Address to be Delivered": edit_address, "Products JSON": json.dumps(updated_products_dict)
                        }
                        success = safe_approve_po(target_id, updated_fields)
                        if success:
                            st.toast(f"🎉 Success! {target_id} moved safely to approved_orders.xlsx!", icon="✅")
                            if "selected_verification_po" in st.session_state:
                                del st.session_state["selected_verification_po"]
                            st.session_state["extra_items_count"] = 0
                            time.sleep(1)
                            st.rerun()
        else:
            st.subheader("📥 Outstanding Pending Orders")
            st.write("Click on any order card below to open the dedicated split-screen verification interface.")
            
            cols = st.columns(3)
            for idx, row in df_q.iterrows():
                with cols[idx % 3]:
                    with st.container(border=True):
                        st.subheader(f"📦 {row['PO ID']}")
                        st.markdown(f"**🏢 Company:** `{row['Company Name']}`")
                        st.markdown(f"**📍 Destination Region:** `{row['Region']}`")
                        st.markdown(f"**💰 Total Value:** `{row['Total Amount']}`")
                        st.caption(f"📅 Logged: {row['System Timestamp']}")
                        
                        if st.button(f"⚡ Evaluate {row['PO ID']}", key=f"eval_{row['PO ID']}", use_container_width=True):
                            st.session_state["selected_verification_po"] = row["PO ID"]
                            st.rerun()