import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally settings
HOST = "localhost"
PORT = "9000"
COMPANY = ""

STOCK_ITEM_OUTPUT_COLUMNS = [
    "Name",
    "Parent",
    "Category",
    "LedgerName",
    "OpeningBalance",
    "OpeningValue",
    "BasicValue",
    "BasicQty",
    "OpeningRate",
]

def strip_ns(tag):
    if not isinstance(tag, str):
        return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None:
        return ""
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text))
    return text.strip()

def cn(text):
    """Clean numeric strings from Tally XML (handles commas, signs, and Dr/Cr suffixes)"""
    if not text: return 0.0
    cleaned = text.replace(",", "")
    # Find the number part (handles -123.45 or 123.45 Dr)
    matches = re.search(r"([-+]?\d+(?:\.\d+)?)", cleaned)
    if not matches: return 0.0
    val = float(matches.group(1))
    # In Tally XML, assets (Stock) are typically positive. 
    # If it says 'Cr', we treat it as negative (rare for stock).
    if "Cr" in cleaned.upper():
        val = -abs(val)
    return val

def post_to_tally(url, xml_text):
    response = requests.post(url, data=xml_text.encode("utf-8"), timeout=120)
    return response.text

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=15)
        # Using standardized cleanup logic
        from xml.sax.saxutils import escape
        root = ET.fromstring(r.text.encode("utf-8")) # Stock items script had simpler cleanup
        cmp = root.find(".//COMPANY")
        if cmp is not None:
            name = clean_text(cmp.get("NAME")) or clean_text(cmp.findtext("NAME", ""))
            start = clean_text(cmp.findtext("STARTINGFROM", ""))
            end = clean_text(cmp.findtext("ENDINGAT", ""))
            return name, start, end
    except:
        pass
    return "", "", ""


# 1. Fetch
url = f"http://{HOST}:{PORT}"
if not COMPANY:
    COMPANY, _, _ = get_company_info(HOST, PORT)

static_vars = f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
if COMPANY:
    static_vars += f"<SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY>"
static_vars += "</STATICVARIABLES>"

xml_req = (
    "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
    "<TYPE>COLLECTION</TYPE><ID>StockMasterExtract</ID></HEADER><BODY><DESC>"
    f"{static_vars}"
    "<TDL><TDLMESSAGE><COLLECTION NAME=\"StockMasterExtract\"><TYPE>StockItem</TYPE>"
    "<FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH>"
    "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
)

xml_resp = post_to_tally(url, xml_req)
root = ET.fromstring(xml_resp.encode("utf-8"))

rows = []
for elem in root.findall(".//STOCKITEM"):
    name = clean_text(elem.get("NAME")) or elem.findtext("NAME", "")
    if not name: continue
    
    rows.append({
        "Name": name,
        "Parent": elem.findtext("PARENT", ""),
        "Category": elem.findtext("CATEGORY", ""),
        "LedgerName": elem.findtext("LEDGERNAME", ""),
        "OpeningBalance": cn(elem.findtext("OPENINGBALANCE", "0")),
        "OpeningValue": cn(elem.findtext("OPENINGVALUE", "0")),
        "BasicValue": cn(elem.findtext("BASICVALUE", "0")),
        "BasicQty": cn(elem.findtext("BASICQTY", "0")),
        "OpeningRate": cn(elem.findtext("OPENINGRATE", "0")),
    })

dataset = pd.DataFrame(rows) if rows else pd.DataFrame(columns=STOCK_ITEM_OUTPUT_COLUMNS)
