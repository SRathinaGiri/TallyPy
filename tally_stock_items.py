import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
import os
import time
import tempfile
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect

STOCK_ITEM_COLUMNS = ["Name", "Parent", "Category", "LedgerName", "OpeningBalance", "OpeningValue", "BasicValue", "BasicQty", "OpeningRate", "ClosingBalance", "ClosingValue", "ClosingRate", "CompanyName", "FromDate", "ToDate"]

# Core Helper Functions (Line-for-line with app1.py)
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
            if strip_ns(cmp.tag).upper() == "COMPANY":
                return clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

# Power BI Locking & CSV Caching
cache_dir = tempfile.gettempdir()
csv_file = os.path.join(cache_dir, f"tally_StockItem_{PORT}.csv")
lock_file = os.path.join(cache_dir, f"tally_lock_{PORT}.lock")
ready_file = os.path.join(cache_dir, f"tally_ready_StockItem_{PORT}.flag")

if os.path.exists(ready_file) and (time.time() - os.path.getmtime(ready_file)) < 300:
    StockItem = pd.read_csv(csv_file)
else:
    for _ in range(120):
        try:
            fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY); os.close(fd); break
        except:
            if os.path.exists(lock_file) and (time.time() - os.path.getmtime(lock_file)) > 600: os.remove(lock_file)
            time.sleep(1)
            if os.path.exists(ready_file) and (time.time() - os.path.getmtime(ready_file)) < 300:
                StockItem = pd.read_csv(csv_file); break
    else:
        try:
            if os.path.exists(ready_file): os.remove(ready_file)
            url = f"http://{HOST}:{PORT}"
            det_name, det_start, det_end = get_company_info(HOST, PORT)
            sel_comp = COMPANY or det_name
            si_req = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyStockItems</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyStockItems\"><TYPE>StockItem</TYPE><FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH><COMPUTE>ClosingBalance:$_ClosingBalance</COMPUTE><COMPUTE>ClosingValue:$_ClosingValue</COMPUTE><COMPUTE>ClosingRate:$_ClosingRate</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
            root = ET.fromstring(xml_cleanup(post_to_tally(url, si_req)))
            si_rows = []
            for elem in root.iter():
                if strip_ns(elem.tag).upper() != "STOCKITEM": continue
                name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
                if not name: continue
                si_rows.append({"Name": name, "Parent": direct_child_text(elem, "PARENT"), "Category": direct_child_text(elem, "CATEGORY"), "LedgerName": direct_child_text(elem, "LEDGERNAME"), "OpeningBalance": to_float(direct_child_text(elem, "OPENINGBALANCE")), "OpeningValue": to_float(direct_child_text(elem, "OPENINGVALUE")), "BasicValue": to_float(direct_child_text(elem, "BASICVALUE")), "BasicQty": to_float(direct_child_text(elem, "BASICQTY")), "OpeningRate": to_float(direct_child_text(elem, "OPENINGRATE")), "ClosingBalance": to_float(direct_child_text(elem, "CLOSINGBALANCE")), "ClosingValue": to_float(direct_child_text(elem, "CLOSINGVALUE")), "ClosingRate": to_float(direct_child_text(elem, "CLOSINGRATE")), "CompanyName": sel_comp, "FromDate": format_tally_date(det_start), "ToDate": format_tally_date(det_end)})
            StockItem = pd.DataFrame(si_rows, columns=STOCK_ITEM_COLUMNS)
            StockItem.to_csv(csv_file, index=False)
            with open(ready_file, 'w') as f: f.write("done")
        finally:
            if os.path.exists(lock_file): os.remove(lock_file)
