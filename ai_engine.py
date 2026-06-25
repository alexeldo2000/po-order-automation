import json
import streamlit as st
from google import genai
from google.genai import types

# Load client safely
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as e:
    client = None

def extract_po_details_native(pdf_path):
    if not client:
        return {
            "Company Name": "Extraction Error: API Key missing", "Order Date": "Unknown", "Total Amount": "Unknown",
            "Address to be Delivered": "Unknown", "Region": "Unknown", "Products JSON": json.dumps({"Error": "No API Key"})
        }
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
        
        # FIXED: Corrected string literals for backtick checking
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
        return {
            "Company Name": f"Extraction Error: {str(e)}", "Order Date": "Unknown", "Total Amount": "Unknown",
            "Address to be Delivered": "Unknown", "Region": "Missing from Document",
            "Products JSON": json.dumps({"Error": str(e)})
        }