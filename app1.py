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

# --- STEP 1: UI & DATABASE CONFIGURATION ---
st.set_page_config(
    page_title="Mane Kancor | Procurement Desk", 
    layout="wide",
    initial_sidebar_state="expanded"
)

QUEUE_FILE = "purchase_orders.xlsx"
APPROVED_FILE = "approved_orders.xlsx"
DOWNLOADS_DIR = os.path.join(os.getcwd(), "downloads")
STATUS_FILE = "worker_status.txt"

if not os.path.exists(DOWNLOADS_DIR):
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Shared master schema for both Excel files
DTYPE_MAPPING = {
    "PO ID": str, "Purchase Order Number": str, "Message-ID": str, "System Timestamp": str, 
    "Email Sender": str, "Email Subject": str, "Company Name": str, "Order Date": str, 
    "Total Amount": str, "Address to be Delivered": str, "Region": str, 
    "Approval Status": str, "PDF Path": str, "Products JSON": str  
}

def robust_read_excel(file_path):
    """Guarantees all schema columns exist on load and converts columns to clean strings safely."""
    try:
        if os.path.exists(file_path):
            df = pd.read_excel(file_path, engine='openpyxl')
            schema_columns = list(DTYPE_MAPPING.keys())
            
            # Ensure all schema columns exist
            for col in schema_columns:
                if col not in df.columns:
                    df[col] = "None"
            
            # Reorder to match schema exactly
            df = df[schema_columns]
            
            # Safely cast each column to string and strip spaces
            for col in df.columns:
                df[col] = df[col].astype(str).fillna("None").str.strip()
                
            return df
    except Exception as e:
        print(f"Error reading Excel {file_path}: {e}")
    return pd.DataFrame(columns=list(DTYPE_MAPPING.keys())).astype(DTYPE_MAPPING)

def init_excel_files():
    """Initializes Excel files if they do not exist, or migrates them if schema is outdated."""
    headers = list(DTYPE_MAPPING.keys())
    for file_path in [QUEUE_FILE, APPROVED_FILE]:
        if not os.path.exists(file_path):
            pd.DataFrame(columns=headers).astype(DTYPE_MAPPING).to_excel(file_path, index=False)
        else:
            try:
                df = robust_read_excel(file_path)
                df.to_excel(file_path, index=False)
            except:
                pass

init_excel_files()

excel_lock = threading.Lock()

def safe_read_excel(file_path):
    """Safely reads Excel files with thread locks and normalized columns."""
    with excel_lock:
        return robust_read_excel(file_path)

def safe_approve_po(po_uid, updated_row_dict=None):
    with excel_lock:
        try:
            df_queue = robust_read_excel(QUEUE_FILE)
            df_approved = robust_read_excel(APPROVED_FILE)
            
            row_to_move = df_queue[df_queue["PO ID"] == po_uid]
            if not row_to_move.empty:
                row_copy = row_to_move.copy()
                
                if updated_row_dict:
                    for key, val in updated_row_dict.items():
                        if key in row_copy.columns:
                            row_copy[key] = str(val)
                
                row_copy["Approval Status"] = "Approved"
                row_copy["System Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                df_approved = pd.concat([df_approved, row_copy], ignore_index=True)
                df_approved.to_excel(APPROVED_FILE, index=False)
                
                df_queue = df_queue[df_queue["PO ID"] != po_uid]
                df_queue.to_excel(QUEUE_FILE, index=False)
                return True
        except Exception as e:
            st.error(f"Error executing database transfer: {e}")
        return False

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
        Analyze this raw input PDF file directly and extract the required parameters.
        
        Locate the specific Purchase Order (PO) Number, Ref Code, Order Number, or Contract Number. 
        It is typically labeled with identifiers such as "PO No.", "PO Number", "Purchase Order No.", "PO#", or "Our Ref".
        Extract this exact alphanumeric string value. If absolutely no order identifier is present, default to "Unknown".
        
        Identify EVERY SINGLE product or line item ordered in this document without truncation.
        Extract them structurally as a flat JSON dictionary where the Key is the product description/name, 
        and the Value is its corresponding ordered quantity/unit metric (e.g., {"Product A": "100 Kg"}).
        """

        # Corrected structured Schema definition using clean production API specifications
        response_schema = types.Schema(
            type=types.Type.OBJECT,
            properties={
                "purchase_order_number": types.Schema(
                    type=types.Type.STRING,
                    description="The extracted PO Number, Ref No, P.O. No., or Order ID from the document."
                ),
                "company_name": types.Schema(type=types.Type.STRING),
                "order_date": types.Schema(type=types.Type.STRING),
                "products_dict": types.Schema(
                    type=types.Type.OBJECT,
                    description="A key-value dictionary representing product items mapped to their quantitative measurements.",
                ),
                "total_amount": types.Schema(type=types.Type.STRING),
                "address_to_be_delivered": types.Schema(type=types.Type.STRING),
                "region": types.Schema(type=types.Type.STRING),
            },
            required=["purchase_order_number", "company_name", "order_date", "products_dict", "total_amount", "address_to_be_delivered", "region"]
        )

        # Targeting stable flagship gemini-2.5-flash for accurate pipeline execution
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
                prompt
            ],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                response_schema=response_schema
            ),
        )

        raw_json_text = response.text.strip()
        ml_result = json.loads(raw_json_text)
        products_map = ml_result.get("products_dict", {})
        
        if not isinstance(products_map, dict):
            products_map = {"Extracted Line Items Summary": str(products_map)}

        return {
            "Purchase Order Number": str(ml_result.get("purchase_order_number", "Unknown")).strip(),
            "Company Name": str(ml_result.get("company_name", "Unknown")).strip(),
            "Order Date": str(ml_result.get("order_date", "Unknown")).strip(),
            "Total Amount": str(ml_result.get("total_amount", "Unknown")).strip(),
            "Address to be Delivered": str(ml_result.get("address_to_be_delivered", "Unknown")).strip(),
            "Region": str(ml_result.get("region", "Unknown")).strip(),
            "Products JSON": json.dumps(products_map)
        }

    except Exception as e:
        error_msg = f"{str(e)}"
        return {
            "Purchase Order Number": "Unknown", "Company Name": "Unknown", "Order Date": "Unknown", 
            "Total Amount": "Unknown", "Address to be Delivered": "Unknown", "Region": "Missing from Document",
            "Products JSON": json.dumps({"Error": error_msg})
        }


# --- STEP 4: SEAMLESS BACKGROUND EXTRACTION ENGINE WITH PRE-FILTER DEDUPLICATION ---
def continuous_mail_sync_loop():
    while True:
        try:
            write_worker_status(f"🔄 Checking mail servers...")
            
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
            mail.login(EMAIL_USER, EMAIL_PASS)
            mail.select("INBOX")
            
            status, messages = mail.search(None, 'ALL')
            mail_ids = messages[0].split()
            mail_ids = mail_ids[-30:] # Scan the latest 30 inbox emails for processing

            write_worker_status(f"🟢 Active. Scanning {len(mail_ids)} latest inbox items...")

            if mail_ids:
                for mail_id in reversed(mail_ids):
                    try:
                        res, msg_data = mail.fetch(mail_id, '(RFC822)')
                        
                        msg = None
                        subject = ""
                        sender = ""
                        msg_id = ""
                        
                        # Extract basic headers first to check duplication before downloading files
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                subject = clean_text(decode_header(msg["Subject"])[0][0])
                                sender = str(msg.get("From"))
                                msg_id = str(msg.get("Message-ID", "")).strip()
                                break
                        
                        if not msg:
                            continue
                            
                        if not msg_id:
                            msg_id = f"FALLBACK-{mail_id}"
                        
                        # PRE-FILTER DEDUPLICATION: Load active databases and check Message-ID first
                        df_current_q = safe_read_excel(QUEUE_FILE)
                        df_current_a = safe_read_excel(APPROVED_FILE)
                        
                        msg_ids_q = df_current_q["Message-ID"].astype(str).str.strip() if not df_current_q.empty else pd.Series(dtype=str)
                        msg_ids_a = df_current_a["Message-ID"].astype(str).str.strip() if not df_current_a.empty else pd.Series(dtype=str)
                        
                        if (not msg_ids_q.empty and (msg_ids_q == msg_id).any()) or \
                           (not msg_ids_a.empty and (msg_ids_a == msg_id).any()):
                            # Already parsed and documented this exact message. Skip instantly!
                            continue
                        
                        # If Message-ID is brand new, walk parts to extract PDF purchase orders
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
                                
                                write_worker_status(f"🧠 Running extraction engine for file...")
                                extracted = extract_po_details_native(filepath)
                                
                                # SEMANTIC DUPLICATION: Check if this Company + Date already exists in Queue
                                is_repetition = False
                                existing_po_id = None
                                
                                if not df_current_q.empty:
                                    comp_names_cleansed = df_current_q["Company Name"].astype(str).str.strip().str.lower()
                                    order_dates_cleansed = df_current_q["Order Date"].astype(str).str.strip().str.lower()
                                    
                                    target_comp = str(extracted["Company Name"]).strip().lower()
                                    target_date = str(extracted["Order Date"]).strip().lower()
                                    
                                    if (comp_names_cleansed == target_comp).any() and (order_dates_cleansed == target_date).any():
                                        is_repetition = True
                                        existing_po_id = df_current_q.loc[(comp_names_cleansed == target_comp) & (order_dates_cleansed == target_date), "PO ID"].values[0]
                                
                                if is_repetition and existing_po_id:
                                    # Overwrite existing record in queue with the latest details
                                    write_worker_status(f"🔄 Updating existing {existing_po_id} with latest data...")
                                    with excel_lock:
                                        df_latest = robust_read_excel(QUEUE_FILE)
                                        idx = df_latest[df_latest["PO ID"] == existing_po_id].index
                                        if not idx.empty:
                                            df_latest.loc[idx, "Purchase Order Number"] = str(extracted["Purchase Order Number"])
                                            df_latest.loc[idx, "Message-ID"] = msg_id
                                            df_latest.loc[idx, "System Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                            df_latest.loc[idx, "Email Sender"] = sender
                                            df_latest.loc[idx, "Email Subject"] = str(subject)
                                            df_latest.loc[idx, "Company Name"] = str(extracted["Company Name"])
                                            df_latest.loc[idx, "Order Date"] = str(extracted["Order Date"])
                                            df_latest.loc[idx, "Total Amount"] = str(extracted["Total Amount"])
                                            df_latest.loc[idx, "Address to be Delivered"] = str(extracted["Address to be Delivered"])
                                            df_latest.loc[idx, "Region"] = str(extracted["Region"])
                                            df_latest.loc[idx, "Products JSON"] = extracted["Products JSON"]
                                            df_latest.loc[idx, "PDF Path"] = str(filepath)
                                            df_latest.loc[idx, "Approval Status"] = "Pending"
                                            df_latest.to_excel(QUEUE_FILE, index=False)
                                else:
                                    # Log as a brand new unique record
                                    write_worker_status(f"📥 Logging new entry inside register spreadsheet...")
                                    new_row = {
                                        "PO ID": po_uid, 
                                        "Purchase Order Number": str(extracted["Purchase Order Number"]),
                                        "Message-ID": msg_id,
                                        "System Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        "Email Sender": sender, "Email Subject": str(subject),
                                        "Company Name": str(extracted["Company Name"]), "Order Date": str(extracted["Order Date"]),
                                        "Total Amount": str(extracted["Total Amount"]), "Address to be Delivered": str(extracted["Address to be Delivered"]),
                                        "Region": str(extracted["Region"]), "Approval Status": "Pending",
                                        "PDF Path": str(filepath), "Products JSON": extracted["Products JSON"]
                                    }
                                    with excel_lock:
                                        df_latest = robust_read_excel(QUEUE_FILE)
                                        df_new_record = pd.DataFrame([new_row])
                                        pd.concat([df_latest, df_new_record], ignore_index=True).to_excel(QUEUE_FILE, index=False)
                                
                                time.sleep(1)
                    except Exception as inner_ex:
                        pass
            
            mail.close()
            mail.logout()
            write_worker_status("Idle. Database sync completed.")
        except Exception as e:
            write_worker_status(f"🔴 Connection Error: Running reconnect procedures...")
            
        time.sleep(12)

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
    
    @st.fragment(run_every=3)
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


# --- WORKSPACE A: OPERATIONAL DASHBOARD (LIVE ISOLATED STREAM) ---
if workspace == "📊 Operational Dashboard":
    st.title("📊 Operational Management Insights")
    st.write("Real-time summary tracking structural supply document ingestion workflows.")
    st.divider()
    
    @st.fragment(run_every=3)
    def render_live_dashboard_view():
        df_q = safe_read_excel(QUEUE_FILE)
        df_a = safe_read_excel(APPROVED_FILE)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Awaiting Audit (in purchase_orders.xlsx)", value=len(df_q))
        with col2:
            st.metric(label="Successfully Verified (in approved_orders.xlsx)", value=len(df_a))
            
        st.subheader("🕒 Live Incoming Mail Feed (Unapproved Rows Only)")
        
        if not df_q.empty:
            refined_df = df_q[["PO ID", "Purchase Order Number", "Company Name", "System Timestamp"]].copy()
            refined_df.columns = ["System ID", "Doc PO Number", "Company Name", "Ingestion Inbound Time"]
            st.dataframe(
                refined_df.sort_values(by="Ingestion Inbound Time", ascending=False), 
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No temporary records are active inside your incoming email workspace.")
            
    render_live_dashboard_view()


# --- WORKSPACE B: PO SECTION (PENDING ONLY) ---
elif workspace == "📁 PO Section (Pending)":
    st.title("📁 Unverified Purchase Order Registry (Queue File)")
    st.write("Lightweight view of outstanding pipeline extractions waiting on verification desks.")
    st.divider()
    
    @st.fragment(run_every=3)
    def render_live_pending_queue():
        df_q = safe_read_excel(QUEUE_FILE)
        if not df_q.empty:
            clean_view_df = df_q[["PO ID", "Purchase Order Number", "Company Name", "System Timestamp"]].copy()
            clean_view_df.columns = ["System ID", "Doc PO Number", "Company Name", "Logged Time"]
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
                col_target = cols[idx % 3]
                with col_target:
                    with st.container(border=True):
                        st.subheader(f"✅ {row['PO ID']}")
                        st.write(f"**🔢 Doc PO #:** `{row['Purchase Order Number']}`")
                        st.write(f"**🏢 Company:** {row['Company Name']}")
                        st.write(f"**💰 Total:** {row['Total Amount']}")
                        
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
                        st.text_input("Document Purchase Order Number", value=str(selected_row["Purchase Order Number"]), disabled=True, key="view_ponum")
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


# --- WORKSPACE D: INTERACTIVE AUDIT DESK (STABLE SPLIT-SCREEN WORKSPACE) ---
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
                st.warning("Selected purchase order was not found.")
                st.session_state["selected_verification_po"] = None
                st.rerun()
            else:
                po_data = matching_rows.iloc[0]
                
                col_back_btn, col_title_lbl = st.columns([1, 4])
                with col_back_btn:
                    if st.button("⬅️ Back to Selection list", use_container_width=True):
                        st.session_state["selected_verification_po"] = None
                        st.session_state["extra_items_count"] = 0
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
                    edit_po_number = st.text_input("Purchase Order Number (PO #)", value=str(po_data["Purchase Order Number"]), key="form_po_number")
                    edit_company = st.text_input("Company Name", value=str(po_data["Company Name"]), key="form_company")
                    edit_date = st.text_input("Order Date Reference", value=str(po_data["Order Date"]), key="form_date")
                    edit_amount = st.text_input("Total Order Amount", value=str(po_data["Total Amount"]), key="form_amount")
                    edit_region = st.text_input("Assigned Region", value=str(po_data["Region"]), key="form_region")
                    edit_address = st.text_area("Delivery Address Target", value=str(po_data["Address to be Delivered"]), height=70, key="form_address")
                    
                    st.write(f"**📬 Email Source Sender:** `{po_data['Email Sender']}`")
                    st.write(f"**📧 Subject Line:** {po_data['Email Subject']}")
                    
                    st.write("##### 📦 Itemized Full Manifest Catalog")
                    try:
                        parsed_products = json.loads(po_data["Products JSON"])
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
                            e_qty = st.text_input("New Qty Metric", key=f"extra_qty_{i}")
                        if e_name:
                            updated_products_dict[e_name] = e_qty

                    st.divider()
                    
                    # Submit Verification Gate Actions
                    col_approve, col_reject = st.columns(2)
                    
                    with col_approve:
                        if st.button("✅ Commit & Approve PO Record", type="primary", use_container_width=True):
                            updated_payload = {
                                "Purchase Order Number": str(edit_po_number),
                                "Company Name": str(edit_company),
                                "Order Date": str(edit_date),
                                "Total Amount": str(edit_amount),
                                "Region": str(edit_region),
                                "Address to be Delivered": str(edit_address),
                                "Products JSON": json.dumps(updated_products_dict)
                            }
                            
                            success = safe_approve_po(target_id, updated_row_dict=updated_payload)
                            if success:
                                st.success(f"🎉 System ID {target_id} successfully synchronized to master ledgers!")
                                st.session_state["selected_verification_po"] = None
                                st.session_state["extra_items_count"] = 0
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error("❌ Thread synchronization fault. Could not commit records.")

                    with col_reject:
                        if st.button("🗑️ Delete/Drop Pending PO", type="secondary", use_container_width=True):
                            with excel_lock:
                                try:
                                    df_latest_q = robust_read_excel(QUEUE_FILE)
                                    df_latest_q = df_latest_q[df_latest_q["PO ID"] != target_id]
                                    df_latest_q.to_excel(QUEUE_FILE, index=False)
                                    
                                    st.warning(f"Pipeline entry {target_id} dropped successfully.")
                                    st.session_state["selected_verification_po"] = None
                                    st.session_state["extra_items_count"] = 0
                                    time.sleep(1)
                                    st.rerun()
                                except Exception as drop_ex:
                                    st.error(f"Drop Exception Triggered: {drop_ex}")
        else:
            # Main Entry Selection Grid Matrix Screen
            st.subheader("📬 Outstanding Document Ingestion Backlog Grid")
            st.write("Click on any row element down below to parse and verify parameters against original asset configurations.")
            
            selection_df = df_q[["PO ID", "Purchase Order Number", "Company Name", "System Timestamp"]].copy()
            selection_df.columns = ["System ID", "Doc PO Number", "Company Name", "Ingestion Inbound Time"]
            
            st.dataframe(
                selection_df.sort_values(by="Ingestion Inbound Time", ascending=False),
                use_container_width=True,
                hide_index=True
            )
            
            target_options = list(df_q["PO ID"].unique())
            selected_box = st.selectbox("🎯 Match Core System ID to Open Target Workspace Desk Panel:", ["-- Select Active PO --"] + target_options)
            
            if selected_box != "-- Select Active PO --":
                st.session_state["selected_verification_po"] = selected_box
                st.rerun()