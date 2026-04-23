import pandas as pd
import requests
import xml.etree.ElementTree as ET
import re
from xml.sax.saxutils import escape

# --- CONFIGURATION ---
HOST, PORT = "localhost", "9000"
URL = f"http://{HOST}:{PORT}"

# --- CORE UTILITIES EXACTLY FROM APP1.PY ---
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
        except: return ""
        if codepoint in (9, 10, 13) or (32 <= codepoint <= 55295) or (57344 <= codepoint <= 65533):
            return match.group(0)
        return ""
    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", fix_char_ref, xml_text)
    xml_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_text)
    xml_text = re.sub(r"&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;)", "&amp;", xml_text)
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)
    return xml_text

def to_float(value):
    text = clean_text(value).replace(",", "")
    if not text: return 0.0
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    try: return float(matches[-1].group(0)) if matches else 0.0
    except: return 0.0

# --- GET COMPANY INFO (EXACT LOGIC FROM APP1.PY) ---
def get_company_info():
    # Re-inserting the <TDL><TDLMESSAGE> wrappers exactly as per app1.py
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE>"
        "<FETCH>Name, StartingFrom, EndingAt</FETCH>"
        "<FILTER>IsActiveCompany</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        r = requests.post(URL, data=xml.encode("utf-8"), timeout=10)
        cleaned = xml_cleanup(r.text)
        root = ET.fromstring(cleaned.encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.findtext("NAME"))
                start = clean_text(cmp.findtext("STARTINGFROM"))
                end = clean_text(cmp.findtext("ENDINGAT"))
                if name: return name, start, end
    except: pass
    return "", "", ""

COMPANY_NAME, FROM_DATE, TO_DATE = get_company_info()

# --- FETCH STOCK ITEMS (EXACT LOGIC FROM APP1.PY) ---
static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
if COMPANY_NAME:
    static_vars.append(f"<SVCURRENTCOMPANY>{escape(COMPANY_NAME)}</SVCURRENTCOMPANY>")

req_xml = (
    "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
    "<TYPE>COLLECTION</TYPE><ID>MyStockItems</ID></HEADER><BODY><DESC>"
    f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
    "<TDL><TDLMESSAGE>"
    "<COLLECTION NAME=\"MyStockItems\"><TYPE>StockItem</TYPE>"
    "<FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH>"
    "<COMPUTE>ClosingBalance:$_ClosingBalance</COMPUTE>"
    "<COMPUTE>ClosingValue:$_ClosingValue</COMPUTE>"
    "<COMPUTE>ClosingRate:$_ClosingRate</COMPUTE>"
    "</COLLECTION>"
    "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
)

resp = requests.post(URL, data=req_xml.encode("utf-8"), timeout=120)
root = ET.fromstring(xml_cleanup(resp.text).encode("utf-8"))

rows = []
for elem in root.iter():
    if strip_ns(elem.tag).upper() == "STOCKITEM":
        name = clean_text(elem.get("NAME") or elem.findtext("NAME"))
        if not name: continue 

        # Godown detection from Batch Allocations if available
        godown = ""
        for child in elem:
            if "BATCHALLOCATIONS" in strip_ns(child.tag).upper():
                godown = clean_text(child.findtext("GODOWNNAME"))
                break

        rows.append({
            "Name": name,
            "Parent": clean_text(elem.findtext("PARENT")),
            "Category": clean_text(elem.findtext("CATEGORY")),
            "LedgerName": clean_text(elem.findtext("LEDGERNAME")),
            "OpeningBalance": to_float(elem.findtext("OPENINGBALANCE")),
            "OpeningValue": to_float(elem.findtext("OPENINGVALUE")),
            "ClosingBalance": to_float(elem.findtext("CLOSINGBALANCE")),
            "ClosingValue": to_float(elem.findtext("CLOSINGVALUE")),
            "GodownName": godown,
            "CompanyName": COMPANY_NAME,
            "FromDate": FROM_DATE,
            "ToDate": TO_DATE
        })

# Use the specific name for Power BI
StockItem = pd.DataFrame(rows)