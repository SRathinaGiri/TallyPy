import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
import os
import time
import tempfile
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape
from datetime import datetime

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect
FROM_DATE = ""    # YYYYMMDD
TO_DATE = ""      # YYYYMMDD

STOCK_VOUCHER_COLUMNS = ["Date", "VoucherTypeName", "VoucherNumber", "StockItemName", "BilledQty", "Rate", "Amount", "GodownName", "BatchName", "VoucherNarration", "CompanyName", "FromDate", "ToDate"]

# Core Helpers (Line-for-line with app1.py)
def strip_ns(tag):
    if not isinstance(tag, str): return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text)).strip()

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        v = match.group(1)
        try: cp = int(v[1:], 16) if v.lower().startswith("x") else int(v)
        except: return ""
        return match.group(0) if cp in (9, 10, 13) or (32 <= cp <= 55295) or (57344 <= cp <= 65533) or (65536 <= cp <= 1114111) else ""
    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", fix_char_ref, xml_text)
    xml_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_text)
    xml_text = re.sub(r"&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;)", "&amp;", xml_text)
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)
    xml_text = re.sub(r'\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*"[^"]*"', "", xml_text)
    return xml_text

def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper(): return clean_text(child.text)
    return ""

def direct_children(elem, local_name):
    return [c for c in list(elem) if strip_ns(c.tag).upper() == local_name.upper()]

def first_non_empty_text(elem, names):
    for n in names:
        v = direct_child_text(elem, n)
        if v: return v
    return ""

def to_decimal(value, default=Decimal("0.00")):
    text = clean_text(value).replace(",", "")
    if not text: return default
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    if not matches: return default
    try: return Decimal(matches[-1].group(0))
    except: return default

def to_float(value):
    return float(to_decimal(value))

def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value): return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value

def post_to_tally(url, xml_text):
    r = requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=120)
    return r.text

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        cleaned = xml_cleanup(post_to_tally(url, xml))
        root = ET.fromstring(cleaned.encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY": return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

# --- MAIN EXECUTION ---

lock_file = os.path.join(tempfile.gettempdir(), f"tally_lock_{PORT}.lock")
for _ in range(300):
    try:
        fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY); os.close(fd); break
    except:
        if os.path.exists(lock_file) and (time.time() - os.path.getmtime(lock_file)) > 600: os.remove(lock_file)
        time.sleep(1)

try:
    cache_dir = tempfile.gettempdir()
    csv_file = os.path.join(cache_dir, f"tally_StockVoucher_{PORT}.csv")
    ready_file = os.path.join(cache_dir, f"tally_ready_StockVoucher_{PORT}.flag")

    if os.path.exists(ready_file) and (time.time() - os.path.getmtime(ready_file)) < 300:
        StockVoucher = pd.read_csv(csv_file)
    else:
        if os.path.exists(ready_file): os.remove(ready_file)
        url = f"http://{HOST}:{PORT}"
        det_name, det_start, det_end = get_company_info(HOST, PORT)
        sel_comp = COMPANY or det_name
        f_dt, t_dt = FROM_DATE or det_start, TO_DATE or det_end

        # Surgical Inventory Fetch (No .* for stability)
        sv_xml = (
            f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyInv</ID></HEADER><BODY><DESC>"
            f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY>"
            f"<SVFROMDATE TYPE='Date'>{escape(f_dt)}</SVFROMDATE><SVTODATE TYPE='Date'>{escape(t_dt)}</SVTODATE></STATICVARIABLES>"
            "<TDL><TDLMESSAGE>"
            "<COLLECTION NAME=\"MyInv\"><TYPE>Voucher</TYPE>"
            "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, "
            "InventoryEntries.StockItemName, InventoryEntries.Amount, InventoryEntries.BilledQty, InventoryEntries.Rate, "
            "InventoryEntries.IsDeemedPositive, InventoryEntries.BatchAllocations.List</FETCH>"
            "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
        )

        root = ET.fromstring(xml_cleanup(post_to_tally(url, sv_xml)))
        sv_rows = []
        for voucher in root.iter():
            if strip_ns(voucher.tag).upper() != "VOUCHER": continue
            v_type = direct_child_text(voucher, "VOUCHERTYPENAME")
            if "Order" in v_type: continue
            v_date, v_number, v_narration = format_tally_date(direct_child_text(voucher, "DATE")), direct_child_text(voucher, "VOUCHERNUMBER"), first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
            
            # Use Greedy Search matching app1.py but with surgical fetch
            inv_nodes = [child for child in voucher if "INVENTORYENTRIES" in child.tag.upper()]
            for inv in inv_nodes:
                item_name = direct_child_text(inv, "STOCKITEMNAME")
                if not item_name: continue
                is_inward = (direct_child_text(inv, "ISDEEMEDPOSITIVE").upper() == "YES")
                q_val, a_val = abs(to_float(direct_child_text(inv, "BILLEDQTY"))), abs(to_float(direct_child_text(inv, "AMOUNT")))
                batch_nodes = direct_children(inv, "BATCHALLOCATIONS.LIST")
                gn, bn = (direct_child_text(batch_nodes[0], "GODOWNNAME"), direct_child_text(batch_nodes[0], "BATCHNAME")) if batch_nodes else ("", "")
                sv_rows.append({"Date": v_date, "VoucherTypeName": v_type, "VoucherNumber": v_number, "StockItemName": item_name.strip(), "BilledQty": q_val if is_inward else -q_val, "Rate": to_float(direct_child_text(inv, "RATE")), "Amount": float(a_val if is_inward else -a_val), "GodownName": gn, "BatchName": bn, "VoucherNarration": v_narration, "CompanyName": sel_comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
        
        StockVoucher = pd.DataFrame(sv_rows, columns=STOCK_VOUCHER_COLUMNS)
        StockVoucher.to_csv(csv_file, index=False)
        with open(ready_file, 'w') as f: f.write("done")
finally:
    if os.path.exists(lock_file): os.remove(lock_file)
