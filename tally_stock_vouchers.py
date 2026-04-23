import pandas as pd
import requests
import xml.etree.ElementTree as ET
import re
from xml.sax.saxutils import escape

# --- CONFIGURATION ---
HOST, PORT = "localhost", "9000"
URL = f"http://{HOST}:{PORT}"

# --- UTILITIES EXACTLY FROM APP1.PY ---
def strip_ns(tag):
    if not isinstance(tag, str): return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text))
    return text.strip()

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        value = match.group(1)
        try:
            codepoint = int(value[1:], 16) if value.lower().startswith("x") else int(value)
        except Exception: return ""
        if codepoint in (9, 10, 13) or (32 <= codepoint <= 55295) or (57344 <= codepoint <= 65533):
            return match.group(0)
        return ""
    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", fix_char_ref, xml_text)
    xml_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_text)
    xml_text = re.sub(r"&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;)", "&amp;", xml_text)
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)
    return xml_text

def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper():
            return clean_text(child.text)
    return ""

def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value

def to_float(value):
    text = clean_text(value).replace(",", "")
    if not text: return 0.0
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    if not matches: return 0.0
    token = matches[-1].group(0)
    try: return float(token)
    except: return 0.0

# --- GET COMPANY INFO (EXACT TDL FROM APP1.PY) ---
def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE>"
        "<FETCH>Name, StartingFrom, EndingAt, Guid</FETCH>"
        "<FILTER>IsActiveCompany</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=10)
        cleaned = xml_cleanup(r.text)
        root = ET.fromstring(cleaned.encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name: return name, start, end
    except: pass
    return "Unknown", "", ""

# Execute Company detection
COMPANY_NAME, RAW_FROM, RAW_TO = get_company_info(HOST, PORT)
F_FROM = format_tally_date(RAW_FROM)
F_TO = format_tally_date(RAW_TO)

# --- FETCH INVENTORY VOUCHERS ---
static_vars = [
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    f"<SVFROMDATE TYPE='Date'>{escape(RAW_FROM)}</SVFROMDATE>",
    f"<SVTODATE TYPE='Date'>{escape(RAW_TO)}</SVTODATE>"
]
if COMPANY_NAME and COMPANY_NAME != "Unknown":
    static_vars.append(f"<SVCURRENTCOMPANY>{escape(COMPANY_NAME)}</SVCURRENTCOMPANY>")

req_xml = (
    "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
    "<TYPE>COLLECTION</TYPE><ID>MyInventoryVouchers</ID></HEADER><BODY><DESC>"
    f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
    "<TDL><TDLMESSAGE>"
    "<COLLECTION NAME=\"MyInventoryVouchers\"><TYPE>Voucher</TYPE>"
    "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, "
    "InventoryEntries.*, AllInventoryEntries.*, InventoryEntriesIn.*, InventoryEntriesOut.*</FETCH>"
    "</COLLECTION>"
    "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
)

resp = requests.post(URL, data=req_xml.encode("utf-8"), timeout=120)
root = ET.fromstring(xml_cleanup(resp.text).encode("utf-8"))

rows = []
for voucher in root.iter():
    if strip_ns(voucher.tag).upper() == "VOUCHER":
        v_type = direct_child_text(voucher, "VOUCHERTYPENAME")
        if "Order" in v_type: continue
        
        v_date = format_tally_date(direct_child_text(voucher, "DATE"))
        v_number = direct_child_text(voucher, "VOUCHERNUMBER")
        v_narration = clean_text(voucher.findtext("NARRATION") or voucher.findtext("VOUCHERNARRATION"))

        # Greedy search exactly from app1.py
        inv_nodes = [child for child in voucher if "INVENTORYENTRIES" in child.tag.upper()]

        for inv in inv_nodes:
            item_name = direct_child_text(inv, "STOCKITEMNAME")
            if not item_name: continue
            
            is_pos_val = direct_child_text(inv, "ISDEEMEDPOSITIVE")
            is_inward = (is_pos_val.upper() == "YES")

            amount_val = abs(to_float(direct_child_text(inv, "AMOUNT")))
            qty_val = abs(to_float(direct_child_text(inv, "BILLEDQTY")))
            rate_val = to_float(direct_child_text(inv, "RATE"))
            
            batch_nodes = [c for c in inv if strip_ns(c.tag).upper() == "BATCHALLOCATIONS.LIST"]
            godown = direct_child_text(batch_nodes[0], "GODOWNNAME") if batch_nodes else ""
            batch = direct_child_text(batch_nodes[0], "BATCHNAME") if batch_nodes else ""

            rows.append({
                "Date": v_date,
                "VoucherTypeName": v_type,
                "VoucherNumber": v_number,
                "StockItemName": item_name,
                "BilledQty": qty_val if is_inward else -qty_val,
                "Rate": rate_val,
                "Amount": amount_val if is_inward else -amount_val,
                "GodownName": godown,
                "BatchName": batch,
                "VoucherNarration": v_narration,
                "CompanyName": COMPANY_NAME,
                "FromDate": F_FROM,
                "ToDate": F_TO,
            })

StockVoucher = pd.DataFrame(rows)