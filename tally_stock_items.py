import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect

STOCK_ITEM_COLUMNS = ["Name", "Parent", "Category", "LedgerName", "OpeningBalance", "OpeningValue", "BasicValue", "BasicQty", "OpeningRate", "ClosingBalance", "ClosingValue", "ClosingRate", "CompanyName", "FromDate", "ToDate"]

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
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
    xml += "<TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES>"
    xml += "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE>"
    xml += "<COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION>"
    xml += "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
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

si_req = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SI</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"SI\"><TYPE>StockItem</TYPE><FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH><COMPUTE>ClosingBalance:$_ClosingBalance</COMPUTE><COMPUTE>ClosingValue:$_ClosingValue</COMPUTE><COMPUTE>ClosingRate:$_ClosingRate</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"

root = ET.fromstring(xml_cleanup(post_to_tally(url, si_req)))
si_rows = []
for elem in root.iter():
    if strip_ns(elem.tag).upper() == "STOCKITEM":
        si_rows.append({"Name": clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME"), "Parent": direct_child_text(elem, "PARENT"), "Category": direct_child_text(elem, "CATEGORY"), "LedgerName": direct_child_text(elem, "LEDGERNAME"), "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))), "OpeningValue": float(to_decimal(direct_child_text(elem, "OPENINGVALUE"))), "BasicValue": float(to_decimal(direct_child_text(elem, "BASICVALUE"))), "BasicQty": float(to_decimal(direct_child_text(elem, "BASICQTY"))), "OpeningRate": float(to_decimal(direct_child_text(elem, "OPENINGRATE"))), "ClosingBalance": float(to_decimal(direct_child_text(elem, "CLOSINGBALANCE"))), "ClosingValue": float(to_decimal(direct_child_text(elem, "CLOSINGVALUE"))), "ClosingRate": float(to_decimal(direct_child_text(elem, "CLOSINGRATE"))), "CompanyName": sel_comp, "FromDate": format_tally_date(det_start), "ToDate": format_tally_date(det_end)})

dataset = pd.DataFrame(si_rows, columns=STOCK_ITEM_COLUMNS)
dataset = dataset[STOCK_ITEM_COLUMNS]
