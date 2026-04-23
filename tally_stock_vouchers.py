import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal
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

def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value): return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value

def post_to_tally(url, xml_text, timeout=120):
    r = requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=timeout)
    r.raise_for_status()
    return r.content.decode(r.encoding or "utf-8", errors="replace")

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        root = ET.fromstring(xml_cleanup(post_to_tally(url, xml)))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY": return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

# Execution Flow
url = f"http://{HOST}:{PORT}"
det_name, det_start, det_end = get_company_info(HOST, PORT)
sel_comp = COMPANY or det_name
f_dt, t_dt = FROM_DATE or det_start, TO_DATE or det_end

sv_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SV</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY><SVFROMDATE TYPE='Date'>{escape(f_dt)}</SVFROMDATE><SVTODATE TYPE='Date'>{escape(t_dt)}</SVTODATE></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"SV\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, InventoryEntries.*, AllInventoryEntries.*</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"

root = ET.fromstring(xml_cleanup(post_to_tally(url, sv_xml)))
sv_rows = []
for v in root.iter():
    if strip_ns(v.tag).upper() != "VOUCHER" or "Order" in direct_child_text(v, "VOUCHERTYPENAME"): continue
    vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
    inv_nodes = [c for c in list(v) if "INVENTORYENTRIES" in strip_ns(c.tag).upper()]
    for ent in inv_nodes:
        inm = direct_child_text(ent, "STOCKITEMNAME")
        if not inm: continue
        is_in = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
        q, a = abs(float(to_decimal(direct_child_text(ent, "BILLEDQTY")))), abs(float(to_decimal(direct_child_text(ent, "AMOUNT"))))
        batch_nodes = [bc for bc in list(ent) if "BATCHALLOCATIONS.LIST" in strip_ns(bc.tag).upper()]
        gn, bn = (direct_child_text(batch_nodes[0], "GODOWNNAME"), direct_child_text(batch_nodes[0], "BATCHNAME")) if batch_nodes else ("", "")
        sv_rows.append({"Date": vd, "VoucherTypeName": direct_child_text(v, "VOUCHERTYPENAME"), "VoucherNumber": vn, "StockItemName": inm, "BilledQty": q if is_in else -q, "Rate": float(to_decimal(direct_child_text(ent, "RATE"))), "Amount": a if is_in else -a, "GodownName": gn, "BatchName": bn, "VoucherNarration": v_nar, "CompanyName": sel_comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})

dataset = pd.DataFrame(sv_rows, columns=STOCK_VOUCHER_COLUMNS)
dataset = dataset[STOCK_VOUCHER_COLUMNS]
