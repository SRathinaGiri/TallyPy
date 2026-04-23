import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal
from xml.sax.saxutils import escape
from datetime import datetime, timedelta

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect
FROM_DATE = ""    # Leave blank for FY start (YYYYMMDD)
TO_DATE = ""      # Leave blank for FY end (YYYYMMDD)

STOCK_VOUCHER_COLUMNS = ["Date", "VoucherTypeName", "VoucherNumber", "StockItemName", "BilledQty", "Rate", "Amount", "GodownName", "BatchName", "VoucherNarration", "CompanyName", "FromDate", "ToDate"]

# Helper functions
def strip_ns(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text)).strip()

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        val = match.group(1)
        try: cp = int(val[1:], 16) if val.lower().startswith("x") else int(val)
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

def first_non_empty_text(elem, names):
    for name in names:
        v = direct_child_text(elem, name); 
        if v: return v
    return ""

def to_decimal(value, default=Decimal("0.00")):
    text = clean_text(value).replace(",", "")
    if not text: return default
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    if not matches: return default
    try: return Decimal(matches[-1].group(0))
    except: return default

def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value): return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=10)
        root = ET.fromstring(xml_cleanup(r.text).encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY": return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

# Execution
url = f"http://{HOST}:{PORT}"
c_name, s_dt, e_dt = get_company_info(HOST, PORT)
sel_comp = COMPANY or c_name
now = datetime.now()
def_start, def_end = (f"{now.year-1}0401", f"{now.year}0331") if now.month < 4 else (f"{now.year}0401", f"{now.year+1}0331")
f_dt = str(FROM_DATE or s_dt or def_start).strip()
t_dt = str(TO_DATE or e_dt or def_end).strip()
if not re.fullmatch(r"\d{8}", f_dt): f_dt = def_start
if not re.fullmatch(r"\d{8}", t_dt): t_dt = def_end

# Chunking Inventory Entries
sv_rows = []
d1, d2 = datetime.strptime(f_dt, "%Y%m%d"), datetime.strptime(t_dt, "%Y%m%d")
curr = d1
while curr <= d2:
    cs = curr.strftime("%Y%m%d"); ce = min(d2, (curr + timedelta(days=31)).replace(day=1) - timedelta(days=1)); ce_str = ce.strftime("%Y%m%d")
    chunk_sv = f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY><SVFROMDATE TYPE='Date'>{cs}</SVFROMDATE><SVTODATE TYPE='Date'>{ce_str}</SVTODATE></STATICVARIABLES>"
    try:
        rsv = requests.post(url, data=f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SV</ID></HEADER><BODY><DESC>{chunk_sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"SV\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, InventoryEntries.*, AllInventoryEntries.*</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>".encode("utf-8"), timeout=120)
        for v in ET.fromstring(xml_cleanup(rsv.text)).iter():
            if strip_ns(v.tag).upper() != "VOUCHER" or "Order" in direct_child_text(v, "VOUCHERTYPENAME"): continue
            vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
            inv_nodes = [c for c in list(v) if "INVENTORYENTRIES" in strip_ns(c.tag).upper()]
            for ent in inv_nodes:
                inm = direct_child_text(ent, "STOCKITEMNAME"); 
                if not inm: continue
                is_in = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
                q, a = abs(float(to_decimal(direct_child_text(ent, "BILLEDQTY")))), abs(float(to_decimal(direct_child_text(ent, "AMOUNT"))))
                batch_nodes = [bc for bc in list(ent) if "BATCHALLOCATIONS.LIST" in strip_ns(bc.tag).upper()]
                gn, bn = (direct_child_text(batch_nodes[0], "GODOWNNAME"), direct_child_text(batch_nodes[0], "BATCHNAME")) if batch_nodes else ("", "")
                sv_rows.append({"Date": vd, "VoucherTypeName": direct_child_text(v, "VOUCHERTYPENAME"), "VoucherNumber": vn, "StockItemName": inm, "BilledQty": q if is_in else -q, "Rate": float(to_decimal(direct_child_text(ent, "RATE"))), "Amount": a if is_in else -a, "GodownName": gn, "BatchName": bn, "VoucherNarration": v_nar, "CompanyName": sel_comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
    except: pass
    curr = ce + timedelta(days=1)

dataset = pd.DataFrame(sv_rows, columns=STOCK_VOUCHER_COLUMNS)
for c in STOCK_VOUCHER_COLUMNS:
    if c not in dataset.columns: dataset[c] = ""
dataset = dataset[STOCK_VOUCHER_COLUMNS]
