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
    "CompanyName",
    "FromDate",
    "ToDate",
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

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        value = match.group(1)
        try:
            codepoint = int(value[1:], 16) if value.lower().startswith("x") else int(value)
        except Exception:
            return ""
        if codepoint in (9, 10, 13) or (32 <= codepoint <= 55295) or (57344 <= codepoint <= 65533) or (65536 <= codepoint <= 1114111):
            return match.group(0)
        return ""

    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", fix_char_ref, xml_text)
    xml_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_text)
    xml_text = re.sub(r"&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;)", "&amp;", xml_text)

    # Strip namespace prefixes from tags (e.g., <ns0:TAG> -> <TAG>)
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)

    # Strip xmlns declarations to avoid parsing conflicts
    xml_text = re.sub(r'\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*"[^"]*"', "", xml_text)
    xml_text = re.sub(r"\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*'[^']*'", "", xml_text)

    return xml_text

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
                if name:
                    return name, start, end

        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name:
                    return name, start, end
    except:
        pass
    return "", "", ""


# 1. Fetch
url = f"http://{HOST}:{PORT}"
COMPANY_NAME, START_DATE, END_DATE = get_company_info(HOST, PORT)
if not COMPANY:
    COMPANY = COMPANY_NAME

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
root = ET.fromstring(xml_cleanup(xml_resp).encode("utf-8"))

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
dataset["CompanyName"] = COMPANY
dataset["FromDate"] = format_tally_date(START_DATE)
dataset["ToDate"] = format_tally_date(END_DATE)

for column in STOCK_ITEM_OUTPUT_COLUMNS:
    if column not in dataset.columns:
        dataset[column] = ""
dataset = dataset[STOCK_ITEM_OUTPUT_COLUMNS]
