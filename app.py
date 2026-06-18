import streamlit as st
import pandas as pd
import imaplib
import email
from email.header import decode_header
import os
import time
import base64
from datetime import datetime
import json
import threading
from google import genai
from google.genai import types

# --- STEP 1: UI CONFIGURATION ---
st.set_page_config(
    page_title="Mane Kancor | Procurement Desk", 
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILE = "purchase_orders.xlsx"
DOWNLOADS_DIR = os.path.join(os.getcwd(), "downloads")
STATUS_FILE = "worker_status.txt"

if not os.path.exists(DOWNLOADS_DIR):
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Added 'Message-ID' column to strictly track server uniqueness
DTYPE_MAPPING = {
    "PO ID": str, "Message-ID": str, "System Timestamp": str, "Email Sender": str, "Email Subject": str,
    "Company Name": str, "Order Date": str, "Total Amount": str,
    "Address to be Delivered": str, "Region": str, "Approval Status": str, "PDF Path": str,
    "Products JSON": str  
}

def init_excel_file():
    headers = list(DTYPE_MAPPING.keys())
    if not os.path.exists(EXCEL_FILE):
        df_empty = pd.DataFrame(columns=headers).astype(DTYPE_MAPPING)
        df_empty.to_excel(EXCEL_FILE, index=False)
    else:
        try:
            df = pd.read_excel(EXCEL_FILE)
            updated = False
            if "Products JSON" not in df.columns:
                df["Products JSON"] = "{}"
                updated = True
            if "Message-ID" not in df.columns:
                df["Message-ID"] = "Unknown"
                updated = True
            if updated:
                df.astype(DTYPE_MAPPING).to_excel(EXCEL_FILE, index=False)
        except:
            pass

init_excel_file()

# Thread-safe read/write lock for file integrity
excel_lock = threading.Lock()

def safe_read_excel():
    with excel_lock:
        try:
            if os.path.exists(EXCEL_FILE):
                df = pd.read_excel(EXCEL_FILE, dtype=DTYPE_MAPPING, engine='openpyxl')
                return df.fillna("None")
        except Exception as e:
            pass
        return pd.DataFrame(columns=list(DTYPE_MAPPING.keys())).astype(DTYPE_MAPPING)

def write_worker_status(message):
    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            f.write(message)
    except:
        pass

def read_worker_status():
    if not os.path.exists(STATUS_FILE):
        return "Idle"
    try:
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            return f.read()
    except:
        return "Syncing logs..."


# --- STEP 2: SAFE CONFIGURATION LOADING ---
try:
    EMAIL_USER = st.secrets["EMAIL_USER"]
    EMAIL_PASS = st.secrets["EMAIL_PASS"]
    IMAP_SERVER = st.secrets["IMAP_SERVER"]
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    
    client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as config_err:
    st.error(f"⚠️ Secrets Configuration Error: {config_err}")
    st.stop()


# --- STEP 3: NATIVE PRODUCT EXTRACTION ENGINE ---
def clean_text(text):
    if isinstance(text, bytes):
        return text.decode(errors="ignore")
    return text

def extract_po_details_native(pdf_path):
    try:
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        prompt = """
        Analyze this raw input PDF file directly. Extract the required parameters.
        Identify EVERY SINGLE product or line item ordered in this document without truncation.
        Extract them structurally as a flat JSON dictionary where the Key is the product description/name, 
        and the Value is its corresponding ordered quantity/unit metric.
        
        Return your response strictly as a single JSON object with these exact keys:
        {
            "company_name": "Name of issuer/buyer",
            "order_date": "Date of the order document",
            "products_dict": {"Item Description 1": "Quantity"},
            "total_amount": "Grand total value with its currency symbol",
            "address_to_be_delivered": "Full shipping address destination block",
            "region": "Inferred geographical region based on address analysis"
        }
        Do not add any Markdown formatting or wrap it in triple backtick json blocks.
        """

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
                prompt
            ],
            config=types.GenerateContentConfig(temperature=0.0),
        )

        raw_json_text = response.text.strip()
        
        lines = [line.strip() for line in raw_json_text.splitlines() if line.strip()]
        if lines and lines[0][:3] == "```":
            lines.pop(0)
        if lines and lines[-1][:3] == "```":
            lines.pop(-1)
            
        raw_json_text = "\n".join(lines).strip()

        ml_result = json.loads(raw_json_text)
        products_map = ml_result.get("products_dict", {})
        
        if not isinstance(products_map, dict):
            products_map = {"Extracted Line Items Summary": str(products_map)}

        return {
            "Company Name": ml_result.get("company_name", "Unknown"),
            "Order Date": ml_result.get("order_date", "Unknown"),
            "Total Amount": ml_result.get("total_amount", "Unknown"),
            "Address to be Delivered": ml_result.get("address_to_be_delivered", "Unknown"),
            "Region": ml_result.get("region", "Unknown"),
            "Products JSON": json.dumps(products_map)
        }

    except Exception as e:
        error_msg = f"{str(e)}"
        return {
            "Company Name": f"Extraction Error: {error_msg}", "Order Date": "Unknown", "Total Amount": "Unknown",
            "Address to be Delivered": "Unknown", "Region": "Missing from Document",
            "Products JSON": json.dumps({"Error": error_msg})
        }


# --- STEP 4: SEAMLESS BACKGROUND EXTRACTION ENGINE ---
def continuous_mail_sync_loop():
    while True:
        mail = None
        try:
            write_worker_status(f"🔄 Checking mail servers at {datetime.now().strftime('%H:%M:%S')}...")
            
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
            mail.login(EMAIL_USER, EMAIL_PASS)
            mail.select("INBOX")
            
            status, messages = mail.search(None, 'UNSEEN')
            mail_ids = messages[0].split()

            write_worker_status(f"🟢 Active. Unread POs Found: {len(mail_ids)} ({datetime.now().strftime('%H:%M:%S')})")

            if mail_ids:
                for mail_id in mail_ids:
                    is_processed_successfully = False
                    
                    try:
                        res, msg_data = mail.fetch(mail_id, '(RFC822)')
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                
                                subject = clean_text(decode_header(msg["Subject"])[0][0])
                                sender = str(msg.get("From"))
                                msg_id = str(msg.get("Message-ID", f"FALLBACK-{time.time()}"))
                                
                                # FIXED LOGIC: Validate against server Message-ID to process identical filenames safely
                                df_current = safe_read_excel()
                                if not df_current.empty and msg_id in df_current["Message-ID"].astype(str).values:
                                    is_processed_successfully = True
                                    continue
                                
                                for part in msg.walk():
                                    if part.get_content_maintype() == 'multipart':
                                        continue
                                    
                                    filename = part.get_filename()
                                    if filename and filename.lower().endswith('.pdf'):
                                        
                                        po_uid = f"PO-{int(time.time())}"
                                        unique_filename = f"{po_uid}_{filename}"
                                        filepath = os.path.join(DOWNLOADS_DIR, unique_filename)
                                        
                                        with open(filepath, "wb") as f:
                                            f.write(part.get_payload(decode=True))
                                        
                                        write_worker_status(f"🧠 Extracting data fields for {po_uid}...")
                                        extracted = extract_po_details_native(filepath)
                                        
                                        new_row = {
                                            "PO ID": po_uid,
                                            "Message-ID": msg_id,
                                            "System Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            "Email Sender": sender, 
                                            "Email Subject": str(subject),
                                            "Company Name": str(extracted["Company Name"]),
                                            "Order Date": str(extracted["Order Date"]),
                                            "Total Amount": str(extracted["Total Amount"]), 
                                            "Address to be Delivered": str(extracted["Address to be Delivered"]),
                                            "Region": str(extracted["Region"]),
                                            "Approval Status": "Pending",
                                            "PDF Path": str(filepath),
                                            "Products JSON": extracted["Products JSON"]
                                        }
                                        
                                        with excel_lock:
                                            df_latest = pd.read_excel(EXCEL_FILE, dtype=DTYPE_MAPPING, engine='openpyxl').fillna("None")
                                            df_new_record = pd.DataFrame([new_row]).astype(DTYPE_MAPPING)
                                            pd.concat([df_latest, df_new_record], ignore_index=True).to_excel(EXCEL_FILE, index=False)
                                        
                                        is_processed_successfully = True
                                        time.sleep(0.5)
                                        
                        if not is_processed_successfully:
                            is_processed_successfully = True
                            
                    except Exception as inner_ex:
                        pass
                    finally:
                        if is_processed_successfully:
                            mail.store(mail_id, '+FLAGS', '\\Seen')
                        
            try:
                mail.close()
                mail.logout()
            except:
                pass
                
        except Exception as e:
            write_worker_status(f"🔴 Pipeline Sync Error: {str(e)}")
            
        time.sleep(10)

# --- STEP 5: INITIALIZE BACKGROUND PIPELINE SINGLETON ---
if "background_sync_started" not in st.session_state:
    st.session_state["background_sync_started"] = True
    bg_thread = threading.Thread(target=continuous_mail_sync_loop, daemon=True)
    bg_thread.start()


# --- STEP 6: SIDEBAR NAVIGATION & ISOLATED LIVE STATUS FRAGMENT ---
with st.sidebar:
    st.title("MANE KANCOR")
    st.caption("Procurement Automation Desk")
    st.divider()

    workspace = st.radio(
        "Choose Active Workspace Desk:", 
        ["📊 Operational Dashboard", "📁 PO Section (Pending)", "✅ Approved PO Registry", "🔍 Verification Desk"]
    )

    st.divider()
    
    # Auto-refreshes every 4 seconds dynamically without resetting page state or selections
    @st.fragment(run_every=4)
    def render_sidebar_status_widget():
        st.subheader("📻 Sync Engine Status")
        st.markdown(f"**Status:** {read_worker_status()}")
        
        df_live = safe_read_excel()
        pending_count = len(df_live[df_live["Approval Status"] == "Pending"]) if not df_live.empty else 0
        st.write(f"Total PO Records in Excel: **{len(df_live)}**")
        st.write(f"Pending Audits: **{pending_count}**")
        
    render_sidebar_status_widget()

    if st.button("🔄 Refresh Application UI Manually", use_container_width=True):
        st.rerun()


# --- WORKSPACE A: OPERATIONAL DASHBOARD (LIVE ISOLATED STREAM) ---
if workspace == "📊 Operational Dashboard":
    st.title("📊 Operational Management Insights")
    st.write("Real-time summary tracking structural supply document ingestion workflows.")
    st.divider()
    
    # Live Isolated Stream Fragment: Updates UI tables automatically every 4 seconds
    @st.fragment(run_every=4)
    def render_live_dashboard_view():
        df_live_dash = safe_read_excel()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(label="Total Logged Pipeline Items", value=len(df_live_dash))
        with col2:
            st.metric(label="Awaiting Audit Verification", value=len(df_live_dash[df_live_dash["Approval Status"] == "Pending"]) if not df_live_dash.empty else 0)
        with col3:
            st.metric(label="Successfully Verified Registry", value=len(df_live_dash[df_live_dash["Approval Status"] == "Approved"]) if not df_live_dash.empty else 0)
            
        st.subheader("🕒 Production Processing Feed (Auto-Updating Live)")
        
        if not df_live_dash.empty:
            refined_dashboard_df = df_live_dash[["PO ID", "Company Name", "System Timestamp", "Approval Status"]].copy()
            refined_dashboard_df.columns = ["Order ID", "Company Name", "Ingestion Timestamp", "Audit Status"]
            
            st.dataframe(
                refined_dashboard_df.sort_values(by="Ingestion Timestamp", ascending=False), 
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No records are currently stored inside the local pipeline registry.")
            
    render_live_dashboard_view()


# --- WORKSPACE B: PO SECTION (PENDING ONLY) ---
elif workspace == "📁 PO Section (Pending)":
    st.title("📁 Unverified Purchase Order Registry")
    st.write("Lightweight view of outstanding pipeline extractions waiting on verification desks.")
    st.divider()
    
    @st.fragment(run_every=4)
    def render_live_pending_queue():
        df_live_data = safe_read_excel()
        df_pending = df_live_data[df_live_data["Approval Status"] == "Pending"] if not df_live_data.empty else pd.DataFrame()
        
        if not df_pending.empty:
            clean_view_df = df_pending[["PO ID", "Company Name", "System Timestamp"]].copy()
            clean_view_df.columns = ["Order ID", "Company Name", "Logged Time"]
            st.dataframe(clean_view_df, use_container_width=True, hide_index=True)
        else:
            st.success("🎉 Ingestion queue clear! No outstanding unverified rows found.")
            
    render_live_pending_queue()


# --- WORKSPACE C: APPROVED LOG ARCHIVE ---
elif workspace == "✅ Approved PO Registry":
    st.title("✅ Permanent Cleared Production Ledger")
    st.write("Historical lookup workstation for purchase orders successfully validated.")
    st.divider()
    
    df_data = safe_read_excel()
    df_approved = df_data[df_data["Approval Status"] == "Approved"] if not df_data.empty else pd.DataFrame()
    
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
                col_target = cols[idx % 3]
                
                with col_target:
                    with st.container(border=True):
                        st.subheader(f"📦 {row['PO ID']}")
                        st.write(f"**🏢 Company:** {row['Company Name']}")
                        st.write(f"**🌍 Region:** {row['Region']} | **💰 Total:** {row['Total Amount']}")
                        
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
                        
                        st.write("##### 📦 Closed Item Manifest Breakdown")
                        try:
                            parsed_products = json.loads(selected_row["Products JSON"])
                        except:
                            parsed_products = {"Unknown Item Summary": "None"}
                            
                        for prod_item, prod_qty in parsed_products.items():
                            col_p1, col_p2 = st.columns([2, 1])
                            with col_p1:
                                st.text_input("Item Description Label", value=str(prod_item), disabled=True, key=f"view_pname_{prod_item}_{selected_row['PO ID']}")
                            with col_p2:
                                st.text_input("Committed Quantities", value=str(prod_qty), disabled=True, key=f"view_pqty_{prod_item}_{selected_row['PO ID']}")


# --- WORKSPACE D: INTERACTIVE AUDIT DESK (STABLE ENTRY FORMS) ---
elif workspace == "🔍 Verification Desk":
    st.title("🔍 Selection Matrix & Split-Screen Verification Desk")
    st.write("Review unverified extraction models side-by-side with original file attachment feeds.")
    st.divider()
    
    df_data = safe_read_excel()
    df_pending = df_data[df_data["Approval Status"] == "Pending"] if not df_data.empty else pd.DataFrame()
    
    if df_pending.empty:
        st.success("🎉 Verification desk cleared! No items require manual balancing cycles.")
        st.session_state["selected_verification_po"] = None
    else:
        st.write("### 📂 Choose an Outstanding Box to Launch Workspace Desk:")
        
        cols = st.columns(3)
        for idx, row in df_pending.iterrows():
            col_target = cols[idx % 3]
            
            with col_target:
                with st.container(border=True):
                    st.subheader(f"📦 {row['PO ID']}")
                    st.write(f"**🏢 Company:** {row['Company Name']}")
                    st.write(f"**🌍 Region:** {row['Region']} | **💰 Total:** {row['Total Amount']}")
                    
                    if st.button(f"⚡ Start Audit Evaluation", key=f"btn_{row['PO ID']}", use_container_width=True):
                        st.session_state["selected_verification_po"] = row["PO ID"]
                        st.rerun()

        if st.session_state.get("selected_verification_po"):
            target_po_id = st.session_state["selected_verification_po"]
            matched_rows = df_pending[df_pending["PO ID"] == target_po_id]
            
            if matched_rows.empty:
                st.session_state["selected_verification_po"] = None
                st.rerun()
            else:
                selected_row = matched_rows.iloc[0]
                
                st.divider()
                st.subheader(f"🛠️ Live Interactive Workspace Core: {selected_row['PO ID']}")
                
                panel_left, panel_right = st.columns([1, 1])
                
                with panel_left:
                    st.write("##### 📄 Original Source PDF Preview")
                    pdf_target_path = selected_row["PDF Path"]
                    
                    if os.path.exists(pdf_target_path):
                        try:
                            with open(pdf_target_path, "rb") as f:
                                base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700" type="application/pdf"></iframe>'
                            st.markdown(pdf_display, unsafe_allow_html=True)
                        except Exception as render_ex:
                            st.error(f"Render Panel Error Engine: {render_ex}")
                    else:
                        st.error("⚠️ Local source attachment payload missing on storage paths.")
                        
                with panel_right:
                    st.markdown("##### ✏️ Modifiable Form Parameters")
                    
                    edit_company = st.text_input("Company Name Reference", value=str(selected_row["Company Name"]), key="edit_comp")
                    edit_date = st.text_input("Order Date Reference", value=str(selected_row["Order Date"]), key="edit_dt")
                    edit_amount = st.text_input("Total Amount Invoiced Summary", value=str(selected_row["Total Amount"]), key="edit_amt")
                    edit_address = st.text_area("Address Target Deliver Destination", value=str(selected_row["Address to be Delivered"]), height=70, key="edit_addr")
                    edit_region = st.text_input("Region Mapping Destination Zone", value=str(selected_row["Region"]), key="edit_reg")
                    
                    st.write("##### 📦 Itemized Full Manifest Catalog (Dynamic Input Fields)")
                    
                    try:
                        parsed_products = json.loads(selected_row["Products JSON"])
                    except:
                        parsed_products = {"Unknown Item": "None"}
                        
                    updated_products_dict = {}
                    for prod_item, prod_qty in parsed_products.items():
                        col_p1, col_p2 = st.columns([2, 1])
                        with col_p1:
                            new_name = st.text_input(f"Product Detail", value=str(prod_item), key=f"name_{prod_item}")
                        with col_p2:
                            new_qty = st.text_input(f"Qty Metrics", value=str(prod_qty), key=f"qty_{prod_item}")
                        if new_name:
                            updated_products_dict[new_name] = new_qty
                            
                    if "extra_items_count" not in st.session_state:
                        st.session_state["extra_items_count"] = 0
                        
                    if st.button("➕ Append New Custom Item Row", use_container_width=True):
                        st.session_state["extra_items_count"] += 1
                        st.rerun()
                        
                    for i in range(st.session_state["extra_items_count"]):
                        col_e1, col_e2 = st.columns([2, 1])
                        with col_e1:
                            e_name = st.text_input("New Item Description", key=f"extra_name_{i}")
                        with col_e2:
                            e_qty = st.text_input("New Item Quantity", key=f"extra_qty_{i}")
                        if e_name:
                            updated_products_dict[e_name] = e_qty
                    
                    st.divider()
                    
                    if st.button("🚀 Authorize Validation Entry and Commit", use_container_width=True, key="approve_action_btn"):
                        with excel_lock:
                            df_master_latest = pd.read_excel(EXCEL_FILE, dtype=DTYPE_MAPPING, engine='openpyxl').fillna("None")
                            row_idx = df_master_latest[df_master_latest["PO ID"] == selected_row["PO ID"]].index
                            
                            if not row_idx.empty:
                                df_master_latest.at[row_idx[0], "Company Name"] = str(edit_company)
                                df_master_latest.at[row_idx[0], "Order Date"] = str(edit_date)
                                df_master_latest.at[row_idx[0], "Total Amount"] = str(edit_amount)
                                df_master_latest.at[row_idx[0], "Address to be Delivered"] = str(edit_address)
                                df_master_latest.at[row_idx[0], "Region"] = str(edit_region)
                                df_master_latest.at[row_idx[0], "Approval Status"] = "Approved"
                                df_master_latest.at[row_idx[0], "Products JSON"] = json.dumps(updated_products_dict)
                                
                                df_master_latest = df_master_latest.astype(DTYPE_MAPPING)
                                df_master_latest.to_excel(EXCEL_FILE, index=False)
                        
                        st.session_state["selected_verification_po"] = None
                        st.session_state["extra_items_count"] = 0
                        
                        st.toast(f"🎉 Approved! Record successfully committed.", icon="✅")
                        time.sleep(0.5)
                        st.rerun()