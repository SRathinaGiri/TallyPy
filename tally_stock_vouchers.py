import pandas as pd
import requests
import xml.etree.ElementTree as ET
import re
import calendar
import hashlib
import os
import time
from datetime import date, datetime, timedelta
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

def parse_tally_date_value(value):
    text = clean_text(value)
    if not text:
        return None
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None

def tally_request_date(value):
    parsed = parse_tally_date_value(value)
    return parsed.strftime("%Y%m%d") if parsed else clean_text(value)

def split_period(start_date, end_date, mode):
    chunks = []
    current = start_date
    while current <= end_date:
        if mode == "weekly":
            chunk_end = min(end_date, current + timedelta(days=6))
        elif mode == "daily":
            chunk_end = current
        else:
            chunk_end = min(end_date, date(current.year, current.month, calendar.monthrange(current.year, current.month)[1]))
        chunks.append((current, chunk_end))
        current = chunk_end + timedelta(days=1)
    return chunks

def post_to_tally(url, xml_text):
    r = requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=120)
    r.raise_for_status()
    return r.text

def build_voucher_probe_request_xml(company, from_date, to_date):
    static_vars = [
        "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
        f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>",
        f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>",
    ]
    if company and company != "Unknown":
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVoucherProbe</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"MyVoucherProbe\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, MasterID, VoucherTypeName</FETCH>"
        "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

def plan_export_chunks(url, company, from_date, to_date):
    start_date = parse_tally_date_value(from_date)
    end_date = parse_tally_date_value(to_date)
    if not start_date or not end_date or start_date > end_date:
        return [(clean_text(from_date), clean_text(to_date))]
    try:
        root = ET.fromstring(xml_cleanup(post_to_tally(url, build_voucher_probe_request_xml(company, from_date, to_date))).encode("utf-8"))
        total = 0
        month_counts = {}
        for voucher in root.iter():
            if strip_ns(voucher.tag).upper() != "VOUCHER":
                continue
            total += 1
            voucher_date = parse_tally_date_value(direct_child_text(voucher, "DATE"))
            if voucher_date:
                key = (voucher_date.year, voucher_date.month)
                month_counts[key] = month_counts.get(key, 0) + 1
        if total <= 2000:
            return [(tally_request_date(from_date), tally_request_date(to_date))]
        mode = "monthly" if max(month_counts.values() or [0]) <= 2000 else "weekly"
    except Exception:
        mode = "monthly"
    return [(start.strftime("%Y%m%d"), end.strftime("%Y%m%d")) for start, end in split_period(start_date, end_date, mode)]

def fetch_xml_cached(url, xml_text, company, table_name, from_date, to_date):
    if os.environ.get("TALLYXML_DISABLE_CACHE") == "1":
        return post_to_tally(url, xml_text)
    cache_root = os.path.join(os.getcwd(), ".tally_cache")
    key_text = "|".join([clean_text(company), table_name, clean_text(from_date), clean_text(to_date), "xml-v3"])
    path = os.path.join(cache_root, hashlib.sha1(key_text.encode("utf-8", errors="ignore")).hexdigest() + ".xml")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8", errors="replace") as handle:
            return handle.read()
    last_error = None
    for attempt in range(3):
        try:
            xml = post_to_tally(url, xml_text)
            os.makedirs(cache_root, exist_ok=True)
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(xml)
            return xml
        except Exception as exc:
            last_error = exc
            time.sleep(1 + attempt * 2)
    raise last_error

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

rows = []
def build_flat_inventory_request_xml(company, from_date, to_date):
    static_vars = [
        "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
        f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>",
        f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>"
    ]
    if company and company != "Unknown":
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>TXMLFlatInventoryRows</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"TXMLBaseInventoryVouchers\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration</FETCH>"
        "</COLLECTION>"
        "<COLLECTION NAME=\"TXMLFlatInventoryRows\"><SOURCECOLLECTION>TXMLBaseInventoryVouchers</SOURCECOLLECTION>"
        "<WALK>All Inventory Entries</WALK>"
        "<FETCH>StockItemName, BilledQty, Rate, Amount, TXMLDate, TXMLVoucherTypeName, TXMLVoucherNumber, "
        "TXMLVoucherNarration, TXMLCompanyName, TXMLSignedQty, TXMLSignedAmount, TXMLGodownName, TXMLBatchName</FETCH>"
        "<COMPUTE>TXMLDate:$..Date</COMPUTE>"
        "<COMPUTE>TXMLVoucherTypeName:$..VoucherTypeName</COMPUTE>"
        "<COMPUTE>TXMLVoucherNumber:$..VoucherNumber</COMPUTE>"
        "<COMPUTE>TXMLVoucherNarration:$..Narration</COMPUTE>"
        "<COMPUTE>TXMLCompanyName:##SVCurrentCompany</COMPUTE>"
        "<COMPUTE>TXMLSignedQty:If $IsDeemedPositive Then $$Abs:$BilledQty Else $$Abs:$BilledQty * -1</COMPUTE>"
        "<COMPUTE>TXMLSignedAmount:If $IsDeemedPositive Then $$Abs:$Amount Else $$Abs:$Amount * -1</COMPUTE>"
        "<COMPUTE>TXMLGodownName:$GodownName</COMPUTE>"
        "<COMPUTE>TXMLBatchName:$BatchName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

def append_flat_inventory_rows(root):
    before = len(rows)
    for inv in root.iter():
        if strip_ns(inv.tag).upper() != "INVENTORYENTRY":
            continue
        item_name = direct_child_text(inv, "STOCKITEMNAME")
        v_type = direct_child_text(inv, "TXMLVOUCHERTYPENAME")
        if not item_name or "Order" in v_type:
            continue
        rows.append({
            "Date": format_tally_date(direct_child_text(inv, "TXMLDATE")),
            "VoucherTypeName": v_type,
            "VoucherNumber": direct_child_text(inv, "TXMLVOUCHERNUMBER"),
            "StockItemName": clean_text(item_name),
            "BilledQty": to_float(direct_child_text(inv, "TXMLSIGNEDQTY")),
            "Rate": to_float(direct_child_text(inv, "RATE")),
            "Amount": to_float(direct_child_text(inv, "TXMLSIGNEDAMOUNT")),
            "GodownName": direct_child_text(inv, "TXMLGODOWNNAME"),
            "BatchName": direct_child_text(inv, "TXMLBATCHNAME"),
            "VoucherNarration": direct_child_text(inv, "TXMLVOUCHERNARRATION"),
            "CompanyName": direct_child_text(inv, "TXMLCOMPANYNAME") or COMPANY_NAME,
            "FromDate": F_FROM,
            "ToDate": F_TO,
        })
    return len(rows) - before

for chunk_from, chunk_to in plan_export_chunks(URL, COMPANY_NAME, RAW_FROM, RAW_TO):
    xml = fetch_xml_cached(URL, build_flat_inventory_request_xml(COMPANY_NAME, chunk_from, chunk_to), COMPANY_NAME, "inventory_flat", chunk_from, chunk_to)
    append_flat_inventory_rows(ET.fromstring(xml_cleanup(xml).encode("utf-8")))

StockVoucher = pd.DataFrame(rows)
