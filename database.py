# STREAMING_CHUNK: Defining database operations and locking mechanisms...
import os
import pandas as pd
import threading
from datetime import datetime
from config import QUEUE_FILE, APPROVED_FILE, DTYPE_MAPPING, STATUS_FILE

# Global thread lock for database write safety
excel_lock = threading.Lock()

def init_excel_files():
    """Initializes Excel files if they do not exist."""
    headers = list(DTYPE_MAPPING.keys())
    if not os.path.exists(QUEUE_FILE):
        pd.DataFrame(columns=headers).astype(DTYPE_MAPPING).to_excel(QUEUE_FILE, index=False)
    if not os.path.exists(APPROVED_FILE):
        pd.DataFrame(columns=headers).astype(DTYPE_MAPPING).to_excel(APPROVED_FILE, index=False)

def safe_read_excel(file_path):
    """Safely reads an Excel file with lock protection."""
    with excel_lock:
        try:
            if os.path.exists(file_path):
                df = pd.read_excel(file_path, dtype=DTYPE_MAPPING, engine='openpyxl')
                return df.fillna("None")
        except Exception as e:
            pass
        return pd.DataFrame(columns=list(DTYPE_MAPPING.keys())).astype(DTYPE_MAPPING)

def safe_approve_po(po_uid, updated_row_dict=None):
    """Transfers a row from purchase_orders.xlsx to approved_orders.xlsx securely."""
    with excel_lock:
        try:
            df_queue = pd.read_excel(QUEUE_FILE, dtype=DTYPE_MAPPING, engine='openpyxl').fillna("None")
            df_approved = pd.read_excel(APPROVED_FILE, dtype=DTYPE_MAPPING, engine='openpyxl').fillna("None")
            
            row_to_move = df_queue[df_queue["PO ID"] == po_uid]
            if not row_to_move.empty:
                row_copy = row_to_move.copy()
                
                # Apply manually reviewed edits
                if updated_row_dict:
                    for key, val in updated_row_dict.items():
                        if key in row_copy.columns:
                            row_copy[key] = str(val)
                
                row_copy["Approval Status"] = "Approved"
                row_copy["System Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # Commit to approved workbook
                df_approved = pd.concat([df_approved, row_copy], ignore_index=True)
                df_approved.astype(DTYPE_MAPPING).to_excel(APPROVED_FILE, index=False)
                
                # Pop out of incoming queue workbook
                df_queue = df_queue[df_queue["PO ID"] != po_uid]
                df_queue.astype(DTYPE_MAPPING).to_excel(QUEUE_FILE, index=False)
                return True
        except Exception as e:
            print(f"Error handling database swap: {e}")
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