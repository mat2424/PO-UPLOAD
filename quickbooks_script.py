import requests
import json
import pandas as pd

# QuickBooks API Configuration
REALM_ID = "9341454108635678"
ACCESS_TOKEN = "eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..gOs1XxDB0v2GEZ9-Tlagkw.rjaq_nx3LGN49AhK77dJo7_wudPOxDfYIJTzXMId_wwnGsB-PXQyoSF6J-CFKrbEadylwRNZdMKjqpFKJjak4EsDxrYPxCtmDi3dPMZToN9fVHdsep-gDUrbwK75ZaIwv4BGsXASR7HonUmI2Bh7B-Fqnd52eam8VDxIyQOCYTlYU10BFgxEScoCNAo8tcPp112z0eDjizjG15Yt-ThcEunlT-fRVCwtIimwOsujCtetdxqEpqVKnW19eyzel0lY9j20IuWW93X9pA14THXAhadB9l-1KGf6dOAEj6c7T-EjR4r_t61sHix88MH06RFuwww5gf7X6zGLhdGdj6pI2P0EQwU7TDT0ZthnTNurFqQbmxmJFGMJJAoZVF7plMC09tTTYMoBZ5V5GBqDyPNyUjjizP2QvYjtSMAAMZtnKikqBPHCSpZTKC6Vy4gRYBPfoJmCS6qfDVebf4KhwOrG2dLnZIt4hTCinhH0SovRHJ25dj1ksW_plUpiYvBzYnjzg8WIBPNoEXtPrjSTTIS_jC0sXDXGfVcWxA8Jztg7G0cSAqwFXLDxEUrDUpXOf6oQ8xVyQYwj348EcVFZFsoWbLfNsJ5wrbz4KhcT3qenvzbHYRpPa8BpYBv49RmVClR4akGIvZX8u-VecsuBDjKzWndX0sOkf6j-dVKCE9aayogwjunKB9DyR9yyK9whVDv79XUu0eqAcszh46IlrRkXzXxwri2cGG0Fs3rskzkIwhg7XjGCcp8uVz_wma3Y7x6OwZR2p1CGUtsnZs0PjCk6DSq6cs_n00gF_OZv2NzAZ5bI7jj_ETbZMtJnVoR7H0TL-bA7zofR1yZTIzGng9XaTLeWi2_NztxntjQUhRQcRpLf_6RP0ZXuju6xGmroMz8FU5g6-cBL9-9d4NnGp-kviw.xXHjbWGhAyFeW2JU9q95fg"  # Make sure it's valid
BASE_URL = f"https://sandbox-quickbooks.api.intuit.com/v3/company/{REALM_ID}"

HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Accept": "application/json",
    "Content-Type": "application/json",
    "charset": "UTF-8"
}

# Read Excel File
EXCEL_PATH = r"C:\Users\mathe\Downloads\TADS\Backlog report PO Update.xlsm"
SHEET_NAME = "PurchaseOrders"

try:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    exit()

# Function to get vendor ID
def get_existing_vendor_id(vendor_name):
    url = f"{BASE_URL}/query"
    query = f"SELECT Id FROM Vendor WHERE DisplayName = '{vendor_name}'"
    response = requests.get(f"{url}?query={query}", headers=HEADERS)

    if response.status_code == 200:
        data = response.json()
        vendors = data.get("QueryResponse", {}).get("Vendor", [])
        if vendors:
            return vendors[0]["Id"]
    print(f"❌ API Response for vendor '{vendor_name}': {response.text}")
    return None

# Function to get item ID
def get_existing_item_id(item_name):
    url = f"{BASE_URL}/query"
    query = f"SELECT Id FROM Item WHERE Name = '{item_name}'"
    response = requests.get(f"{url}?query={query}", headers=HEADERS)

    if response.status_code == 200:
        data = response.json()
        items = data.get("QueryResponse", {}).get("Item", [])
        if items:
            return items[0]["Id"]
    print(f"❌ API Response for item '{item_name}': {response.text}")
    return None

# Function to create a purchase order
def create_purchase_order(po_number, vendor_id, items):
    po_payload = {
        "APAccountRef": {"value": "33"},
        "VendorRef": {"value": vendor_id},
        "Line": items
    }
    response = requests.post(f"{BASE_URL}/purchaseorder", headers=HEADERS, json=po_payload)

    if response.status_code == 200:
        print(f"✅ Successfully created PO {po_number}")
    else:
        print(f"❌ Error creating PO {po_number}: {response.text}")

# Process each purchase order
for _, row in df.iterrows():
    vendor_name = row.get("Vendor Name", "").strip()
    item_name = row.get("Item Name", "").strip()
    po_number = str(row.get("PO Number", "")).strip()
    quantity = row.get("Quantity", 0)
    rate = row.get("Rate", 0)

    if not vendor_name or not item_name:
        print(f"⚠️ Skipping PO {po_number} due to missing vendor or item.")
        continue

    vendor_id = get_existing_vendor_id(vendor_name)
    item_id = get_existing_item_id(item_name)

    if not vendor_id or not item_id:
        print(f"⚠️ Skipping PO {po_number} due to vendor/item retrieval failure.")
        continue

    po_line = [{
        "DetailType": "ItemBasedExpenseLineDetail",
        "Amount": quantity * rate,
        "ItemBasedExpenseLineDetail": {
            "ItemRef": {"value": item_id},
            "Qty": quantity,
            "UnitPrice": rate
        }
    }]

    create_purchase_order(po_number, vendor_id, po_line)
