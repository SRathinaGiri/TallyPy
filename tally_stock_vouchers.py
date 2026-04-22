import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally settings
HOST = "localhost"
PORT = "9000"
COMPANY = ""   # Leave blank to auto-detect first company
FROM_DATE = "20200401"              # YYYYMMDD
TO_DATE = "20210331"                # YYYYMMDD

INVENTORY_VOUCHER_TYPES = {
    "Sales",
    "Purchase",
    "Debit Note",
    "Credit Note",
    "Stock Journal",
    "Delivery Note",
    "Receipt Note",
    "Physical Stock",
}

STOCK_VOUCHER_OUTPUT_COLUMNS = [
    "Date",
    "VoucherTypeName",
    "VoucherNumber",
    "StockItemName",
    "BilledQty",
    "Rate",
    "Amount",
    "GodownName",
    "BatchName",
    "VoucherNarration",
    "CompanyName",
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
    return xml_text


def parse_xml_root(xml_text):
    return ET.fromstring(xml_cleanup(xml_text))


def direct_children(elem, local_name):
    wanted = local_name.upper()
    return [child for child in list(elem) if strip_ns(child.tag).upper() == wanted]


def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper():
            return clean_text(child.text)
    return ""


def first_descendant_text(elem, local_name):
    for child in elem.iter():
        if strip_ns(child.tag).upper() == local_name.upper():
            value = clean_text(child.text)
            if value:
                return value
    return ""


def first_non_empty_text(elem, names):
    for name in names:
        value = direct_child_text(elem, name)
        if value:
            return value
    return ""


def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value


def normalize_amount_text(value):
    text = clean_text(value).replace(",", "")
    if not text:
        return ""
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    if not matches:
        return text
    token = matches[-1].group(0)
    try:
        return f"{Decimal(token):.2f}"
    except InvalidOperation:
        return token


def to_decimal(value, default=Decimal("0.00")):
    value = normalize_amount_text(value)
    if not value:
        return default
    try:
        return Decimal(value)
    except InvalidOperation:
        return default

def to_float(value):
    return float(to_decimal(value))


def detect_company_name(root):
    for elem in root.iter():
        if strip_ns(elem.tag).upper() == "COMPANY":
            name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
            if name:
                return name
    return ""


def post_to_tally(url, xml_text, timeout=120):
    response = requests.post(
        url,
        data=xml_text.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        timeout=timeout,
    )
    response.raise_for_status()
    encoding = response.encoding or "utf-8"
    return response.content.decode(encoding, errors="replace")


def build_company_request_xml():
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>List of Companies</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "</DESC></BODY></ENVELOPE>"
    )


def build_inventory_request_xml(company, from_date, to_date):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(from_date)}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(to_date)}</SVTODATE>")

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyInventoryVouchers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyInventoryVouchers\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, "
        "InventoryEntries.StockItemName, InventoryEntries.BilledQty, "
        "InventoryEntries.Rate, InventoryEntries.Amount, "
        "InventoryEntries.IsDeemedPositive, "
        "InventoryEntries.BatchAllocations.BatchName, "
        "InventoryEntries.BatchAllocations.GodownName</FETCH>"
        "<FILTER>HasInventory</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE='Formulae' NAME='HasInventory'>NOT $$IsEmpty:$InventoryEntries</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def parse_inventory_vouchers(root, company):
    rows = []
    for voucher in root.iter():
        if strip_ns(voucher.tag).upper() != "VOUCHER":
            continue

        v_type = direct_child_text(voucher, "VOUCHERTYPENAME")
        v_date = format_tally_date(direct_child_text(voucher, "DATE"))
        v_number = direct_child_text(voucher, "VOUCHERNUMBER")
        v_narration = first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
        v_company = first_non_empty_text(voucher, ["COMPANYNAME", "SVCURRENTCOMPANY"]) or company

        inv_nodes = direct_children(voucher, "INVENTORYENTRIES.LIST")
        if not inv_nodes:
             inv_nodes = direct_children(voucher, "ALLINVENTORYENTRIES.LIST")

        for inv in inv_nodes:
            item_name = direct_child_text(inv, "STOCKITEMNAME")
            if not item_name:
                continue

            amount_val = to_decimal(direct_child_text(inv, "AMOUNT"))
            is_pos = direct_child_text(inv, "ISDEEMEDPOSITIVE").upper()
            
            # For inventory, if it's deemed positive (Inward), we usually keep it as positive in quantity.
            # Tally XML Amount for inventory entries is often negative for Sales (Outward).
            # We will follow the XML sign for Amount.
            
            qty_text = direct_child_text(inv, "BILLEDQTY")
            rate_text = direct_child_text(inv, "RATE")
            
            # Simple numeric extraction from quantity string "10 Pcs"
            qty_val = to_float(qty_text)
            rate_val = to_float(rate_text)

            # Get Godown and Batch from first BatchAllocation
            batch_nodes = direct_children(inv, "BATCHALLOCATIONS.LIST")
            godown = ""
            batch = ""
            if batch_nodes:
                godown = direct_child_text(batch_nodes[0], "GODOWNNAME")
                batch = direct_child_text(batch_nodes[0], "BATCHNAME")

            rows.append({
                "Date": v_date,
                "VoucherTypeName": v_type,
                "VoucherNumber": v_number,
                "StockItemName": item_name,
                "BilledQty": qty_val,
                "Rate": rate_val,
                "Amount": float(amount_val),
                "GodownName": godown,
                "BatchName": batch,
                "VoucherNarration": v_narration,
                "CompanyName": v_company,
            })

    return rows


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
        cleaned = xml_cleanup(r.text)
        root = ET.fromstring(cleaned.encode("utf-8"))
        cmp = root.find(".//COMPANY")
        if cmp is not None:
            name = clean_text(cmp.get("NAME")) or clean_text(cmp.findtext("NAME", ""))
            start = clean_text(cmp.findtext("STARTINGFROM", ""))
            end = clean_text(cmp.findtext("ENDINGAT", ""))
            return name, start, end
    except:
        pass
    return "", "", ""


def fetch_inventory_rows(host, port, company, from_date, to_date):
    url = f"http://{host}:{port}"

    if not company or not from_date or not to_date:
        cmp_name, cmp_start, cmp_end = get_company_info(host, port)
        if not company:
            company = cmp_name
        if not from_date:
            from_date = cmp_start
        if not to_date:
            to_date = cmp_end

    inv_root = parse_xml_root(post_to_tally(url, build_inventory_request_xml(company, from_date, to_date)))
    status = clean_text(first_descendant_text(inv_root, "STATUS"))
    if status == "0":
        error_text = first_descendant_text(inv_root, "LINEERROR") or "Tally returned STATUS=0"
        raise ValueError(error_text)

    return parse_inventory_vouchers(inv_root, company)


rows = fetch_inventory_rows(
    host=HOST,
    port=PORT,
    company=COMPANY,
    from_date=FROM_DATE,
    to_date=TO_DATE,
)

dataset = pd.DataFrame(rows)
for column in STOCK_VOUCHER_OUTPUT_COLUMNS:
    if column not in dataset.columns:
        dataset[column] = ""
dataset = dataset[STOCK_VOUCHER_OUTPUT_COLUMNS]
