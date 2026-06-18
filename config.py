# STREAMING_CHUNK: Setting up configuration paths and data schemas...
import os

QUEUE_FILE = "purchase_orders.xlsx"
APPROVED_FILE = "approved_orders.xlsx"
DOWNLOADS_DIR = os.path.join(os.getcwd(), "downloads")
STATUS_FILE = "worker_status.txt"

# Ensure the attachments directory is created
if not os.path.exists(DOWNLOADS_DIR):
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Shared master schema for both Excel files
DTYPE_MAPPING = {
    "PO ID": str, "Message-ID": str, "System Timestamp": str, "Email Sender": str, "Email Subject": str,
    "Company Name": str, "Order Date": str, "Total Amount": str,
    "Address to be Delivered": str, "Region": str, "Approval Status": str, "PDF Path": str,
    "Products JSON": str  
}