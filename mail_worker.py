# STREAMING_CHUNK: Creating asynchronous background mail scraper thread...
import imaplib
import email
from email.header import decode_header
import os
import time
import pandas as pd
from datetime import datetime
import streamlit as st
from config import QUEUE_FILE, APPROVED_FILE, DOWNLOADS_DIR, DTYPE_MAPPING
from database import excel_lock, safe_read_excel, write_worker_status
from ai_engine import extract_po_details_native

try:
    EMAIL_USER = st.secrets["EMAIL_USER"]
    EMAIL_PASS = st.secrets["EMAIL_PASS"]
    IMAP_SERVER = st.secrets["IMAP_SERVER"]
except:
    EMAIL_USER = None

def clean_text(text):
    if isinstance(text, bytes): return text.decode(errors="ignore")
    return text

def continuous_mail_sync_loop():
    if not EMAIL_USER:
        write_worker_status("🔴 Configuration Error: Secrets missing.")
        return

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
                                
                                df_current_q = safe_read_excel(QUEUE_FILE)
                                df_current_a = safe_read_excel(APPROVED_FILE)
                                
                                if (not df_current_q.empty and msg_id in df_current_q["Message-ID"].astype(str).values) or \
                                   (not df_current_a.empty and msg_id in df_current_a["Message-ID"].astype(str).values):
                                    is_processed_successfully = True
                                    continue
                                
                                for part in msg.walk():
                                    if part.get_content_maintype() == 'multipart': continue
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
                                            "PO ID": po_uid, "Message-ID": msg_id, "System Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            "Email Sender": sender, "Email Subject": str(subject), "Company Name": str(extracted["Company Name"]),
                                            "Order Date": str(extracted["Order Date"]), "Total Amount": str(extracted["Total Amount"]), 
                                            "Address to be Delivered": str(extracted["Address to be Delivered"]), "Region": str(extracted["Region"]),
                                            "Approval Status": "Pending", "PDF Path": str(filepath), "Products JSON": extracted["Products JSON"]
                                        }
                                        
                                        with excel_lock:
                                            df_latest = pd.read_excel(QUEUE_FILE, dtype=DTYPE_MAPPING, engine='openpyxl').fillna("None")
                                            df_new_record = pd.DataFrame([new_row]).astype(DTYPE_MAPPING)
                                            pd.concat([df_latest, df_new_record], ignore_index=True).to_excel(QUEUE_FILE, index=False)
                                        
                                        is_processed_successfully = True
                                        time.sleep(0.5)
                        if not is_processed_successfully: is_processed_successfully = True
                    except Exception as inner_ex: pass
                    finally:
                        if is_processed_successfully: mail.store(mail_id, '+FLAGS', '\\Seen')
            try: mail.close(); mail.logout()
            except: pass
        except Exception as e: write_worker_status(f"🔴 Pipeline Sync Error: {str(e)}")
        time.sleep(10)