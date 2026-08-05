import calendar
import hashlib
import io
import os
import re
import time
from datetime import date, datetime, timedelta
import requests
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape
from streamlit_echarts import st_echarts

st.set_page_config(layout="wide", page_title="Tally XML Explorer")

ACCOUNTING_VOUCHER_TYPES = {
    "Sales",
    "Purchase",
    "Journal",
    "Receipt",
    "Payment",
    "Debit Note",
    "Credit Note",
    "Contra",
}

PREDEFINED_VOUCHER_TYPES = {
    "Contra",
    "Payment",
    "Receipt",
    "Journal",
    "Sales",
    "Purchase",
    "Debit Note",
    "Credit Note",
    "Memorandum",
    "Reversing Journal",
    "Delivery Note",
    "Receipt Note",
    "Rejections In",
    "Rejections Out",
    "Stock Journal",
    "Physical Stock",
    "Material In",
    "Material Out",
    "Sales Order",
    "Purchase Order",
    "Job Work In Order",
    "Job Work Out Order",
    "Payroll",
    "Attendance",
}

ACCOUNTING_BASE_VOUCHER_TYPES = {
    "Contra",
    "Payment",
    "Receipt",
    "Journal",
    "Sales",
    "Purchase",
    "Debit Note",
    "Credit Note",
    "Memorandum",
    "Reversing Journal",
}

INVENTORY_BASE_VOUCHER_TYPES = {
    "Delivery Note",
    "Receipt Note",
    "Rejections In",
    "Rejections Out",
    "Stock Journal",
    "Physical Stock",
    "Material In",
    "Material Out",
}

ORDER_BASE_VOUCHER_TYPES = {
    "Sales Order",
    "Purchase Order",
    "Job Work In Order",
    "Job Work Out Order",
}

PAYROLL_BASE_VOUCHER_TYPES = {
    "Payroll",
    "Attendance",
}

# 15 Primary + 13 Sub-groups from Tally documentation
BS_PRIMARY_GROUPS = {
    "Capital Account", "Reserves & Surplus",
    "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans",
    "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors",
    "Fixed Assets",
    "Investments",
    "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors",
    "Misc. Expenses (ASSET)",
    "Suspense Account",
    "Branch / Divisions",
}

PL_PRIMARY_GROUPS = {
    "Sales Accounts",
    "Purchase Accounts",
    "Direct Incomes",
    "Indirect Incomes",
    "Direct Expenses",
    "Indirect Expenses",
}

PRIMARY_GROUPS = BS_PRIMARY_GROUPS | PL_PRIMARY_GROUPS

CURRENCY_SYMBOL_FALLBACKS = {
    "INR": "₹",
    "INDIAN RUPEE": "₹",
    "RUPEE": "₹",
    "RUPEES": "₹",
    "RS": "₹",
    "RS.": "₹",
    "USD": "$",
    "US DOLLAR": "$",
    "DOLLAR": "$",
    "EUR": "€",
    "EURO": "€",
    "GBP": "£",
    "POUND": "£",
    "POUND STERLING": "£",
    "AED": "د.إ",
    "DIRHAM": "د.إ",
    "": "",
}

VOUCHER_COLUMNS = [
    "Date",
    "VoucherTypeName",
    "BaseVoucherType",
    "VoucherNumber",
    "LedgerName",
    "MasterID",
    "Amount",
    "DrCr",
    "DebitAmount",
    "CreditAmount",
    "ParentLedger",
    "PrimaryGroup",
    "Nature",
    "NatureOfGroup",
    "PAN",
    "PartyLedgerName",
    "PartyGSTIN",
    "LedgerGSTIN",
    "VoucherNarration",
    "IsOptional",
    "CompanyName",
    "FromDate",
    "ToDate",
]

ALL_VOUCHER_COLUMNS = VOUCHER_COLUMNS + ["VoucherCategory"]

LEDGER_COLUMNS = [
    "MasterID",
    "Name",
    "PrimaryGroup",
    "Nature",
    "NatureOfGroup",
    "PAN",
    "StartingFrom",
    "CurrencyName",
    "StateName",
    "Parent",
    "PartyGSTIN",
    "OpeningBalance",
    "ClosingBalance",
    "CompanyName",
    "FromDate",
    "ToDate",
]

STOCK_ITEM_COLUMNS = [
    "Name",
    "Parent",
    "Category",
    "LedgerName",
    "OpeningBalance",
    "OpeningValue",
    "BasicValue",
    "BasicQty",
    "OpeningRate",
    "ClosingBalance",
    "ClosingValue",
    "ClosingRate",
    "CompanyName",
    "FromDate",
    "ToDate",
]

STOCK_VOUCHER_COLUMNS = [
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


def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value


def canonical_voucher_type_name(value):
    value = clean_text(value)
    lowered = value.lower()
    aliases = {
        "rejection in": "Rejections In",
        "rejections in": "Rejections In",
        "rejection out": "Rejections Out",
        "rejections out": "Rejections Out",
    }
    return aliases.get(lowered, value)


def voucher_category_from_base_type(base_v_type):
    if base_v_type in ACCOUNTING_BASE_VOUCHER_TYPES:
        return "Accounting"
    if base_v_type in INVENTORY_BASE_VOUCHER_TYPES:
        return "Inventory"
    if base_v_type in ORDER_BASE_VOUCHER_TYPES:
        return "Orders"
    if base_v_type in PAYROLL_BASE_VOUCHER_TYPES:
        return "Payroll"
    return "Unknown"


def nature_from_primary_group(primary_group):
    pg = clean_text(primary_group).lower()
    if pg in [
        "current assets", "fixed assets", "investments", "misc. expenses (asset)",
        "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)",
        "stock-in-hand", "sundry debtors"
    ]:
        return "BS", "Assets"
    elif pg in [
        "capital account", "current liabilities", "loans (liability)", "suspense account",
        "branch / divisions", "bank od a/c", "duties & taxes", "provisions",
        "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"
    ]:
        return "BS", "Liabilities"
    elif pg in ["direct incomes", "indirect incomes", "sales accounts"]:
        return "PL", "Income"
    elif pg in ["direct expenses", "indirect expenses", "purchase accounts"]:
        return "PL", "Expenses"
    return "Unknown", "Unknown"


def ledger_primary_group(ledger_name, ledger_meta):
    seen = set()
    current = clean_text(ledger_name)
    while current and current not in seen:
        seen.add(current)
        meta = ledger_meta.get(current, {})
        parent = clean_text(meta.get("Parent", ""))
        if not parent:
            return ""
        if parent in PRIMARY_GROUPS:
            return parent
        current = parent
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


def month_end(value):
    return date(value.year, value.month, calendar.monthrange(value.year, value.month)[1])


def split_period(start_date, end_date, mode):
    chunks = []
    current = start_date
    while current <= end_date:
        if mode == "weekly":
            chunk_end = min(end_date, current + timedelta(days=6))
        elif mode == "daily":
            chunk_end = current
        else:
            chunk_end = min(end_date, month_end(current))
        chunks.append((current, chunk_end))
        current = chunk_end + timedelta(days=1)
    return chunks


def is_real_voucher(elem):
    if strip_ns(elem.tag).upper() != "VOUCHER":
        return False
    if elem.get("VCHKEY") or elem.get("REMOTEID") or elem.get("VCHTYPE"):
        return True
    return bool(
        direct_child_text(elem, "DATE")
        or direct_child_text(elem, "VOUCHERTYPENAME")
        or direct_child_text(elem, "VOUCHERNUMBER")
        or direct_child_text(elem, "MASTERID")
    )


def voucher_master_id(voucher):
    text = clean_text(voucher.get("MASTERID")) or direct_child_text(voucher, "MASTERID")
    try:
        return int(text)
    except (TypeError, ValueError):
        return None


def make_chunk(from_date, to_date, master_from=None, master_to=None, expected_count=None):
    return {
        "from_date": clean_text(from_date),
        "to_date": clean_text(to_date),
        "master_from": master_from,
        "master_to": master_to,
        "expected_count": expected_count,
    }


def masterid_filter_xml(master_from=None, master_to=None):
    if master_from is None or master_to is None:
        return "", ""
    try:
        master_from = int(master_from)
        master_to = int(master_to)
    except (TypeError, ValueError):
        return "", ""
    collection_filter = "<FILTER>MasterIdRange</FILTER>"
    formula = f"<SYSTEM TYPE=\"Formulae\" NAME=\"MasterIdRange\">$MasterID &gt;= {master_from} AND $MasterID &lt;= {master_to}</SYSTEM>"
    return collection_filter, formula


def build_voucher_probe_request_xml(company, from_date, to_date):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>")
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVoucherProbe</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyVoucherProbe\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, MasterID, VoucherTypeName</FETCH>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def probe_vouchers(url, company, from_date, to_date):
    root = parse_xml_root(post_to_tally(url, build_voucher_probe_request_xml(company, from_date, to_date), timeout=180))
    counts_by_month = {}
    counts_by_day = {}
    entries = []
    total = 0
    for voucher in root.iter():
        if not is_real_voucher(voucher):
            continue
        total += 1
        voucher_date = parse_tally_date_value(direct_child_text(voucher, "DATE"))
        master_id = voucher_master_id(voucher)
        if voucher_date:
            key = (voucher_date.year, voucher_date.month)
            counts_by_month[key] = counts_by_month.get(key, 0) + 1
            counts_by_day[voucher_date.strftime("%Y%m%d")] = counts_by_day.get(voucher_date.strftime("%Y%m%d"), 0) + 1
            if master_id is not None:
                entries.append((voucher_date, master_id))
    return total, counts_by_month, counts_by_day, entries


def build_masterid_chunks(entries, start_date, end_date, max_vouchers=1000):
    master_ids = sorted(master_id for _, master_id in entries if master_id is not None)
    chunks = []
    for index in range(0, len(master_ids), max_vouchers):
        batch = master_ids[index:index + max_vouchers]
        chunks.append(make_chunk(start_date.strftime("%Y%m%d"), end_date.strftime("%Y%m%d"), min(batch), max(batch), len(batch)))
    return chunks


def plan_export_chunks(url, company, from_date, to_date):
    start_date = parse_tally_date_value(from_date)
    end_date = parse_tally_date_value(to_date)
    if not start_date or not end_date or start_date > end_date:
        return [make_chunk(from_date, to_date)]
    try:
        total, counts_by_month, counts_by_day, entries = probe_vouchers(url, company, from_date, to_date)
        target = max(250, int(os.environ.get("TALLYXML_CHUNK_VOUCHERS", "1000") or "1000"))
        if total <= target:
            return [make_chunk(tally_request_date(from_date), tally_request_date(to_date), expected_count=total)]
        if entries:
            return build_masterid_chunks(entries, start_date, end_date, max_vouchers=target)
        mode = "monthly" if max(counts_by_month.values() or [0]) <= 2000 else "weekly"
    except Exception:
        mode = "monthly"
    return [make_chunk(start.strftime("%Y%m%d"), end.strftime("%Y%m%d")) for start, end in split_period(start_date, end_date, mode)]


def cache_file_path(company, table_name, from_date, to_date, master_from=None, master_to=None):
    cache_root = os.path.join(os.getcwd(), ".tally_cache")
    key_text = "|".join([clean_text(company), table_name, clean_text(from_date), clean_text(to_date), clean_text(master_from), clean_text(master_to), "xml-v6"])
    file_name = hashlib.sha1(key_text.encode("utf-8", errors="ignore")).hexdigest() + ".xml"
    return os.path.join(cache_root, file_name)


def fetch_xml_cached(url, xml_text, company, table_name, chunk):
    if os.environ.get("TALLYXML_DISABLE_CACHE") == "1":
        return post_to_tally(url, xml_text)
    path = cache_file_path(company, table_name, chunk["from_date"], chunk["to_date"], chunk.get("master_from"), chunk.get("master_to"))
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                return handle.read()
    except OSError:
        pass
    last_error = None
    for attempt in range(3):
        try:
            xml = post_to_tally(url, xml_text)
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write(xml)
            except OSError:
                pass
            return xml
        except Exception as exc:
            last_error = exc
            time.sleep(1 + attempt * 2)
    raise last_error


def fetch_chunked_rows(url, company, from_date, to_date, table_name, build_request, parse_chunk):
    rows = []
    def process_chunk(chunk):
        chunk_from = chunk["from_date"]
        chunk_to = chunk["to_date"]
        try:
            xml = fetch_xml_cached(
                url,
                build_request(company, chunk_from, chunk_to, chunk.get("master_from"), chunk.get("master_to")),
                company,
                table_name,
                chunk,
            )
            root = parse_xml_root(xml)
            status = clean_text(first_descendant_text(root, "STATUS"))
            if status == "0":
                error_text = first_descendant_text(root, "LINEERROR") or "Tally returned STATUS=0"
                raise ValueError(error_text)
            chunk_rows = parse_chunk(root, chunk_from, chunk_to)
            if chunk.get("master_from") is not None and chunk.get("expected_count") is not None:
                detail_count = sum(1 for elem in root.iter() if is_real_voucher(elem))
                if detail_count and detail_count < chunk["expected_count"]:
                    raise ValueError(f"MasterID-filtered response returned {detail_count} voucher header(s), expected {chunk['expected_count']}.")
            rows.extend(chunk_rows)
        except Exception:
            master_from = chunk.get("master_from")
            master_to = chunk.get("master_to")
            if master_from is not None and master_to is not None:
                master_from = int(master_from)
                master_to = int(master_to)
                if master_to - master_from <= 5:
                    raise
                midpoint = (master_from + master_to) // 2
                left = dict(chunk)
                left["master_to"] = midpoint
                left["expected_count"] = None
                right = dict(chunk)
                right["master_from"] = midpoint + 1
                right["expected_count"] = None
                process_chunk(left)
                process_chunk(right)
                return
            start_date = parse_tally_date_value(chunk_from)
            end_date = parse_tally_date_value(chunk_to)
            if start_date and end_date and start_date < end_date:
                for day_start, day_end in split_period(start_date, end_date, "daily"):
                    day_from = day_start.strftime("%Y%m%d")
                    day_to = day_end.strftime("%Y%m%d")
                    day_chunk = make_chunk(day_from, day_to)
                    xml = fetch_xml_cached(url, build_request(company, day_from, day_to, None, None), company, table_name, day_chunk)
                    rows.extend(parse_chunk(parse_xml_root(xml), day_from, day_to))
            else:
                raise
    for chunk in plan_export_chunks(url, company, from_date, to_date):
        process_chunk(chunk)
    return rows


def build_company_request_xml():
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>List of Companies</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "</DESC></BODY></ENVELOPE>"
    )


def detect_company_name(root):
    for elem in root.iter():
        if strip_ns(elem.tag).upper() == "COMPANY":
            name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
            if name:
                return name
    company_name = first_descendant_text(root, "SVCURRENTCOMPANY")
    if company_name:
        return company_name
    return ""


def build_ledger_request_xml(company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyLedgers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyLedgers\"><TYPE>Ledger</TYPE>"
        "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH>"
        "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE>"
        "<COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def build_flat_voucher_request_xml(company, from_date, to_date, master_from=None, master_to=None):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>")
    collection_filter, filter_formula = masterid_filter_xml(master_from, master_to)

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>TXMLFlatVoucherRows</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"TXMLBaseVouchers\"><TYPE>Voucher</TYPE>"
        f"{collection_filter}"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional</FETCH>"
        "</COLLECTION>"
        "<COLLECTION NAME=\"TXMLFlatVoucherRows\"><SOURCECOLLECTION>TXMLBaseVouchers</SOURCECOLLECTION>"
        "<WALK>All Ledger Entries</WALK>"
        "<FETCH>LedgerName, Amount, IsDeemedPositive, TXMLDate, TXMLVoucherTypeName, TXMLVoucherNumber, "
        "TXMLPartyLedgerName, TXMLPartyGSTIN, TXMLVoucherNarration, TXMLSignedAmount, TXMLDebitAmount, "
        "TXMLCreditAmount, TXMLDrCr, TXMLEntryLedgerMasterID, TXMLEntryParentLedger, TXMLEntryPrimaryGroup, "
        "TXMLEntryLedgerGSTIN, TXMLStatusOptional, TXMLCompanyName</FETCH>"
        "<COMPUTE>TXMLDate:$..Date</COMPUTE>"
        "<COMPUTE>TXMLVoucherTypeName:$..VoucherTypeName</COMPUTE>"
        "<COMPUTE>TXMLVoucherNumber:$..VoucherNumber</COMPUTE>"
        "<COMPUTE>TXMLPartyLedgerName:If NOT $$IsEmpty:$..PartyLedgerName Then $..PartyLedgerName Else \"N/A\"</COMPUTE>"
        "<COMPUTE>TXMLPartyGSTIN:$..PartyGSTIN</COMPUTE>"
        "<COMPUTE>TXMLVoucherNarration:$..Narration</COMPUTE>"
        "<COMPUTE>TXMLSignedAmount:If $IsDeemedPositive Then $$Abs:$Amount * -1 Else $$Abs:$Amount</COMPUTE>"
        "<COMPUTE>TXMLDebitAmount:If $IsDeemedPositive Then $$Abs:$Amount Else 0</COMPUTE>"
        "<COMPUTE>TXMLCreditAmount:If $IsDeemedPositive Then 0 Else $$Abs:$Amount</COMPUTE>"
        "<COMPUTE>TXMLDrCr:If $IsDeemedPositive Then \"Dr\" Else \"Cr\"</COMPUTE>"
        "<COMPUTE>TXMLEntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLEntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLEntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLEntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLStatusOptional:If $..IsOptional Then \"Yes\" Else \"No\"</COMPUTE>"
        "<COMPUTE>TXMLCompanyName:##SVCurrentCompany</COMPUTE>"
        "</COLLECTION>"
        f"{filter_formula}"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def build_stock_item_request_xml(company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")

    return (
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


def build_flat_inventory_entries_request_xml(company, from_date, to_date, master_from=None, master_to=None):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>")
    collection_filter, filter_formula = masterid_filter_xml(master_from, master_to)

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>TXMLFlatInventoryRows</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"TXMLBaseInventoryVouchers\"><TYPE>Voucher</TYPE>"
        f"{collection_filter}"
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
        f"{filter_formula}"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def parse_ledgers(root, group_map=None):
    ledger_rows = []
    ledger_lookup = {}
    
    if group_map is None:
        group_map = {}

    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "LEDGER":
            continue

        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name:
            continue

        parent = direct_child_text(elem, "PARENT")
        
        # Use group_map to get Nature and PrimaryGroup if available
        g_info = group_map.get(parent, {})
        nature_of_group = g_info.get("Nature", "")
        primary_group = g_info.get("PrimaryGroup", "") or first_non_empty_text(elem, ["PRIMARYGROUP"]) or first_descendant_text(elem, "PRIMARYGROUP")

        row = {
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"),
            "Name": name,
            "PrimaryGroup": primary_group,
            "Nature": "",
            "NatureOfGroup": nature_of_group,
            "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]) or first_descendant_text(elem, "INCOMETAXNUMBER"),
            "StartingFrom": first_non_empty_text(elem, ["STARTINGFROM"]) or first_descendant_text(elem, "STARTINGFROM"),
            "CurrencyNameRaw": first_non_empty_text(elem, ["CURRENCYNAME"]) or first_descendant_text(elem, "CURRENCYNAME"),
            "CurrencyFormalNameRaw": first_non_empty_text(elem, ["CURRENCYFORMALNAME"]) or first_descendant_text(elem, "CURRENCYFORMALNAME"),
            "CurrencySymbolRaw": first_non_empty_text(elem, ["CURRENCYSYMBOL"]) or first_descendant_text(elem, "CURRENCYSYMBOL"),
            "CurrencyOriginalSymbolRaw": first_non_empty_text(elem, ["CURRENCYORIGINALSYMBOL"]) or first_descendant_text(elem, "CURRENCYORIGINALSYMBOL"),
            "StateName": first_non_empty_text(elem, ["STATENAME"]) or first_descendant_text(elem, "STATENAME"),
            "Parent": parent,
            "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]) or first_descendant_text(elem, "PARTYGSTIN"),
            "OpeningBalance": to_float(first_non_empty_text(elem, ["OPENINGBALANCE"]) or first_descendant_text(elem, "OPENINGBALANCE")),
            "ClosingBalance": to_float(first_non_empty_text(elem, ["CLOSINGBALANCE"]) or first_descendant_text(elem, "CLOSINGBALANCE")),
        }

        ledger_rows.append(row)
        ledger_lookup[name] = row

    for row in ledger_rows:
        if not row["PrimaryGroup"]:
            row["PrimaryGroup"] = ledger_primary_group(row["Name"], ledger_lookup)
        
        pg = row["PrimaryGroup"]
        if not row["NatureOfGroup"] and pg:
             pg_info = group_map.get(pg, {})
             row["NatureOfGroup"] = pg_info.get("Nature", "")
             
        # Resolve BS/PL and NatureOfGroup mapping
        if row["NatureOfGroup"]:
            n_val = row["NatureOfGroup"].lower()
            if n_val in ["assets", "liabilities"]: row["Nature"] = "BS"
            elif n_val in ["income", "expenses"]: row["Nature"] = "PL"
        
        if not row["Nature"] and pg:
            # Last fallback for Nature if not in group_map
            bs_pl, nog = nature_from_primary_group(pg)
            row["Nature"] = bs_pl
            row["NatureOfGroup"] = nog

        currency_key = clean_text(row.get("CurrencyFormalNameRaw") or row.get("CurrencyNameRaw")).upper()
        row["CurrencyName"] = CURRENCY_SYMBOL_FALLBACKS.get(currency_key, clean_text(row.get("CurrencySymbolRaw") or row.get("CurrencyOriginalSymbolRaw")))

    output_rows = []
    for row in sorted(ledger_rows, key=lambda item: (int(item.get("MasterID") or 0), item.get("Name", ""))):
        output_rows.append({column: row.get(column, "") for column in LEDGER_COLUMNS})
    return output_rows


def parse_flat_vouchers(root, ledger_meta, company, from_date, to_date, vtype_map=None):
    rows = []
    formatted_from_date = format_tally_date(from_date)
    formatted_to_date = format_tally_date(to_date)
    vtype_map = vtype_map or {}

    for elem in root.iter():
        ledger_name = direct_child_text(elem, "LEDGERNAME")
        amount_value = to_decimal(first_non_empty_text(elem, ["TXMLSIGNEDAMOUNT", "SIGNEDAMOUNT", "AMOUNT"]))
        if not ledger_name or amount_value == 0:
            continue

        voucher_type = canonical_voucher_type_name(first_non_empty_text(elem, ["TXMLVOUCHERTYPENAME", "VOUCHERTYPENAME"]))
        base_v_type = canonical_voucher_type_name(vtype_map.get(voucher_type, voucher_type))
        voucher_category = voucher_category_from_base_type(base_v_type)
        base_amount = abs(amount_value)
        signed_amount = amount_value
        if not first_non_empty_text(elem, ["TXMLSIGNEDAMOUNT", "SIGNEDAMOUNT"]):
            signed_amount = base_amount * (Decimal("-1") if direct_child_text(elem, "ISDEEMEDPOSITIVE").upper() == "YES" else Decimal("1"))

        debit_amount = to_decimal(first_non_empty_text(elem, ["TXMLDEBITAMOUNT", "DEBITAMOUNT"]))
        credit_amount = to_decimal(first_non_empty_text(elem, ["TXMLCREDITAMOUNT", "CREDITAMOUNT"]))
        if debit_amount == 0 and credit_amount == 0:
            debit_amount = base_amount if signed_amount < 0 else Decimal("0.00")
            credit_amount = base_amount if signed_amount > 0 else Decimal("0.00")

        meta = ledger_meta.get(ledger_name, {})
        primary_group = first_non_empty_text(elem, ["TXMLENTRYPRIMARYGROUP", "ENTRYPRIMARYGROUP", "PRIMARYGROUP"]) or meta.get("PrimaryGroup", "")
        nature = meta.get("Nature", "")
        nature_of_group = meta.get("NatureOfGroup", "")
        if not nature and primary_group:
            nature, nature_of_group = nature_from_primary_group(primary_group)

        is_optional = first_non_empty_text(elem, ["TXMLSTATUSOPTIONAL", "STATUSOPTIONAL", "ISOPTIONAL"])
        if is_optional.upper() == "YES":
            is_optional = "Yes"
        elif is_optional.upper() == "NO":
            is_optional = "No"

        rows.append({
            "Date": format_tally_date(first_non_empty_text(elem, ["TXMLDATE", "DATE"])),
            "VoucherTypeName": voucher_type,
            "BaseVoucherType": base_v_type,
            "VoucherNumber": first_non_empty_text(elem, ["TXMLVOUCHERNUMBER", "VOUCHERNUMBER"]),
            "LedgerName": ledger_name,
            "MasterID": first_non_empty_text(elem, ["TXMLENTRYLEDGERMASTERID", "ENTRYLEDGERMASTERID", "LEDMASTERID"]) or meta.get("MasterID", ""),
            "Amount": float(signed_amount),
            "DrCr": first_non_empty_text(elem, ["TXMLDRCR", "DRCR"]) or ("Dr" if signed_amount < 0 else "Cr"),
            "DebitAmount": float(debit_amount),
            "CreditAmount": float(credit_amount),
            "ParentLedger": first_non_empty_text(elem, ["TXMLENTRYPARENTLEDGER", "ENTRYPARENTLEDGER", "PARENTLEDGER"]) or meta.get("Parent", ""),
            "PrimaryGroup": primary_group,
            "Nature": nature,
            "NatureOfGroup": nature_of_group,
            "PAN": meta.get("PAN", ""),
            "PartyLedgerName": first_non_empty_text(elem, ["TXMLPARTYLEDGERNAME", "PARTYLEDGERNAME"]) or "N/A",
            "PartyGSTIN": first_non_empty_text(elem, ["TXMLPARTYGSTIN", "PARTYGSTIN"]),
            "LedgerGSTIN": first_non_empty_text(elem, ["TXMLENTRYLEDGERGSTIN", "ENTRYLEDGERGSTIN", "LEDGERGSTIN"]) or meta.get("PartyGSTIN", ""),
            "VoucherNarration": first_non_empty_text(elem, ["TXMLVOUCHERNARRATION", "NARRATION", "VOUCHERNARRATION"]),
            "IsOptional": is_optional or "No",
            "CompanyName": first_non_empty_text(elem, ["TXMLCOMPANYNAME", "COMPANYNAME"]) or company,
            "FromDate": formatted_from_date,
            "ToDate": formatted_to_date,
            "VoucherCategory": voucher_category,
        })

    return rows


def parse_stock_items(root):
    rows = []
    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "STOCKITEM":
            continue
        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name:
            continue
        rows.append({
            "Name": name,
            "Parent": direct_child_text(elem, "PARENT"),
            "Category": direct_child_text(elem, "CATEGORY"),
            "LedgerName": direct_child_text(elem, "LEDGERNAME"),
            "OpeningBalance": to_float(direct_child_text(elem, "OPENINGBALANCE")),
            "OpeningValue": to_float(direct_child_text(elem, "OPENINGVALUE")),
            "BasicValue": to_float(direct_child_text(elem, "BASICVALUE")),
            "BasicQty": to_float(direct_child_text(elem, "BASICQTY")),
            "OpeningRate": to_float(direct_child_text(elem, "OPENINGRATE")),
            "ClosingBalance": to_float(direct_child_text(elem, "CLOSINGBALANCE")),
            "ClosingValue": to_float(direct_child_text(elem, "CLOSINGVALUE")),
            "ClosingRate": to_float(direct_child_text(elem, "CLOSINGRATE")),
        })
    return rows


def parse_flat_inventory_entries(root, company):
    rows = []
    for inv in root.iter():
        if strip_ns(inv.tag).upper() != "INVENTORYENTRY":
            continue
        item_name = first_non_empty_text(inv, ["STOCKITEMNAME"])
        v_type = first_non_empty_text(inv, ["TXMLVOUCHERTYPENAME", "VOUCHERTYPENAME"])
        if not item_name or "Order" in v_type:
            continue
        rows.append({
            "Date": format_tally_date(first_non_empty_text(inv, ["TXMLDATE", "DATE"])),
            "VoucherTypeName": v_type,
            "VoucherNumber": first_non_empty_text(inv, ["TXMLVOUCHERNUMBER", "VOUCHERNUMBER"]),
            "StockItemName": clean_text(item_name),
            "BilledQty": to_float(first_non_empty_text(inv, ["TXMLSIGNEDQTY", "BILLEDQTY"])),
            "Rate": to_float(first_non_empty_text(inv, ["RATE"])),
            "Amount": to_float(first_non_empty_text(inv, ["TXMLSIGNEDAMOUNT", "AMOUNT"])),
            "GodownName": first_non_empty_text(inv, ["TXMLGODOWNNAME", "GODOWNNAME"]),
            "BatchName": first_non_empty_text(inv, ["TXMLBATCHNAME", "BATCHNAME"]),
            "VoucherNarration": first_non_empty_text(inv, ["TXMLVOUCHERNARRATION", "NARRATION", "VOUCHERNARRATION"]),
            "CompanyName": first_non_empty_text(inv, ["TXMLCOMPANYNAME", "COMPANYNAME"]) or company,
        })
    return rows


def get_company_info(host, port):
    url = f"http://{host}:{port}"
    # Requesting Company info specifically from the context of the active session
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
        
        # Try to find the active company in the returned collection
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name:
                    return name, start, end

        # Fallback: If filtered list is empty, try to get the first company found
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name:
                    return name, start, end

    except Exception as e:
        print(f"Error in get_company_info: {e}")
    
    return "", "", ""


def fetch_tally_metadata(url, company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    
    # 1. Fetch Voucher Types
    vtype_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllVoucherTypes</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllVoucherTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    
    # 2. Fetch Groups
    group_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllGroups</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllGroups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    
    vtype_map = {}
    group_map = {}
    
    try:
        # Request 1
        resp_v = post_to_tally(url, vtype_xml)
        root_v = parse_xml_root(resp_v)
        for vt in root_v.iter():
            if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
                name = canonical_voucher_type_name(clean_text(vt.get("NAME")) or direct_child_text(vt, "NAME"))
                parent = canonical_voucher_type_name(direct_child_text(vt, "PARENT"))
                if name:
                    vtype_map[name] = parent or name
        
        # Request 2
        resp_g = post_to_tally(url, group_xml)
        root_g = parse_xml_root(resp_g)
        for g in root_g.iter():
            if strip_ns(g.tag).upper() == "GROUP":
                name = direct_child_text(g, "NAME")
                parent = direct_child_text(g, "PARENT")
                nature = direct_child_text(g, "NATURE")
                primary = direct_child_text(g, "_PRIMARYGROUP")
                if name:
                    group_map[name] = {
                        "Parent": parent,
                        "Nature": nature,
                        "PrimaryGroup": primary
                    }
        
        # Resolve Voucher Types recursively
        base_types = set(PREDEFINED_VOUCHER_TYPES)
        for _ in range(5):
            for vt_name, parent_name in vtype_map.items():
                if parent_name and parent_name not in base_types and parent_name in vtype_map:
                    vtype_map[vt_name] = vtype_map[parent_name]

        # Resolve Group nature recursively
        for _ in range(5):
            for g_name, g_info in group_map.items():
                parent = g_info.get("Parent")
                if parent and not g_info.get("Nature") and parent in group_map:
                    g_info["Nature"] = group_map[parent].get("Nature")
                if parent and not g_info.get("PrimaryGroup") and parent in group_map:
                    g_info["PrimaryGroup"] = group_map[parent].get("PrimaryGroup")

    except:
        pass
        
    return vtype_map, group_map


@st.cache_data(show_spinner=False)
def load_tally_data(host, port, company, from_date, to_date):
    url = f"http://{host}:{port}"
    selected_company = clean_text(company)

    if not selected_company or not from_date or not to_date:
        cmp_name, cmp_start, cmp_end = get_company_info(host, port)
        if not selected_company:
            selected_company = cmp_name
        if not from_date:
            from_date = cmp_start
        if not to_date:
            to_date = cmp_end

    vtype_map, group_map = fetch_tally_metadata(url, selected_company)

    ledger_root = parse_xml_root(post_to_tally(url, build_ledger_request_xml(selected_company)))
    ledger_rows = parse_ledgers(ledger_root, group_map)
    ledger_meta = {row["Name"]: row for row in ledger_rows}

    voucher_rows = fetch_chunked_rows(
        url,
        selected_company,
        from_date,
        to_date,
        "vouchers_flat",
        build_flat_voucher_request_xml,
        lambda root, chunk_from, chunk_to: parse_flat_vouchers(root, ledger_meta, selected_company, chunk_from, chunk_to, vtype_map),
    )
    
    # New Stock Data
    stock_item_root = parse_xml_root(post_to_tally(url, build_stock_item_request_xml(selected_company)))
    stock_item_rows = parse_stock_items(stock_item_root)
    
    inventory_rows = fetch_chunked_rows(
        url,
        selected_company,
        from_date,
        to_date,
        "inventory_flat",
        build_flat_inventory_entries_request_xml,
        lambda root, chunk_from, chunk_to: parse_flat_inventory_entries(root, selected_company),
    )

    all_voucher_df = pd.DataFrame(voucher_rows)
    if all_voucher_df.empty:
        all_voucher_df = pd.DataFrame(columns=ALL_VOUCHER_COLUMNS)
        voucher_df = pd.DataFrame(columns=VOUCHER_COLUMNS)
    else:
        voucher_df = all_voucher_df[all_voucher_df["VoucherCategory"] == "Accounting"].copy()
    ledger_df = pd.DataFrame(ledger_rows)
    stock_item_df = pd.DataFrame(stock_item_rows)
    inventory_df = pd.DataFrame(inventory_rows)

    f_from = format_tally_date(from_date)
    f_to = format_tally_date(to_date)

    # Populate and reorder columns
    df_configs = [
        (voucher_df, VOUCHER_COLUMNS),
        (all_voucher_df, ALL_VOUCHER_COLUMNS),
        (ledger_df, LEDGER_COLUMNS),
        (stock_item_df, STOCK_ITEM_COLUMNS),
        (inventory_df, STOCK_VOUCHER_COLUMNS)
    ]

    final_dfs = []
    for df, cols in df_configs:
        df["CompanyName"] = selected_company
        df["FromDate"] = f_from
        df["ToDate"] = f_to
        for column in cols:
            if column not in df.columns:
                df[column] = ""
        final_dfs.append(df[cols])

    return selected_company, from_date, to_date, *final_dfs


def to_excel_bytes(voucher_df, all_voucher_df, ledger_df, stock_item_df, inventory_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        voucher_df.to_excel(writer, index=False, sheet_name="Vouchers")
        all_voucher_df.to_excel(writer, index=False, sheet_name="All Vouchers")
        ledger_df.to_excel(writer, index=False, sheet_name="Ledgers")
        stock_item_df.to_excel(writer, index=False, sheet_name="Stock Items")
        inventory_df.to_excel(writer, index=False, sheet_name="Stock Vouchers")
    output.seek(0)
    return output.getvalue()


def prepare_dashboard_df(voucher_df):
    df = voucher_df.copy()
    if df.empty:
        return df
    df["DateObj"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Month"] = df["DateObj"].dt.strftime("%b-%Y").fillna("Unknown")
    df["MonthSort"] = df["DateObj"].dt.strftime("%Y%m").fillna("000000")
    df["AbsoluteAmount"] = pd.to_numeric(df["Amount"], errors="coerce").abs().fillna(0)
    return df


st.title("Tally XML Explorer")

with st.sidebar:
    st.header("Connection")
    host = st.text_input("Server", "localhost")
    port = st.text_input("Port", "9000")
    company = st.text_input("Company (optional)", "")
    from_date = st.text_input("From Date (YYYYMMDD, optional)", "")
    to_date = st.text_input("To Date (YYYYMMDD, optional)", "")
    load_btn = st.button("Load Tables", type="primary")
    if st.button("Clear Cache"):
        st.cache_data.clear()
        st.info("Cache cleared.")

if "voucher_df" not in st.session_state:
    st.session_state.voucher_df = None
    st.session_state.all_voucher_df = None
    st.session_state.ledger_df = None
    st.session_state.stock_item_df = None
    st.session_state.inventory_df = None
    st.session_state.company_name = ""
    st.session_state.from_date = ""
    st.session_state.to_date = ""

if load_btn:
    try:
        # Clear existing state for fresh load
        st.session_state.company_name = ""
        st.session_state.from_date = ""
        st.session_state.to_date = ""

        company_name, start_date, end_date, vdf, avdf, ldf, sidf, ivdf = load_tally_data(host, port, company, from_date, to_date)
        
        if not company_name:
            st.error("❌ Failed to detect company name. Please enter it manually in the sidebar.")
        
        st.session_state.voucher_df = vdf
        st.session_state.all_voucher_df = avdf
        st.session_state.ledger_df = ldf
        st.session_state.stock_item_df = sidf
        st.session_state.inventory_df = ivdf
        st.session_state.company_name = company_name or "Unknown Company"
        st.session_state.from_date = start_date
        st.session_state.to_date = end_date
        
        st.success(f"✅ Loaded: {st.session_state.company_name}")
    except Exception as exc:
        st.error(f"Error fetching data: {exc}")

vdf = st.session_state.voucher_df
avdf = st.session_state.all_voucher_df
ldf = st.session_state.ledger_df
sidf = st.session_state.stock_item_df
ivdf = st.session_state.inventory_df

if vdf is not None:
    c_name = st.session_state.company_name
    f_date = format_tally_date(st.session_state.from_date) or "N/A"
    t_date = format_tally_date(st.session_state.to_date) or "N/A"
    st.caption(f"🏢 Company: **{c_name}** | 📅 Period: **{f_date}** to **{t_date}**")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Voucher Rows", len(vdf))
    m2.metric("All Voucher Rows", len(avdf))
    m3.metric("Ledger Rows", len(ldf))
    m4.metric("Stock Item Rows", len(sidf))
    m5.metric("Stock Voucher Rows", len(ivdf))

    v_csv = vdf.to_csv(index=False).encode("utf-8")
    av_csv = avdf.to_csv(index=False).encode("utf-8")
    l_csv = ldf.to_csv(index=False).encode("utf-8")
    si_csv = sidf.to_csv(index=False).encode("utf-8")
    iv_csv = ivdf.to_csv(index=False).encode("utf-8")
    wb_bytes = to_excel_bytes(vdf, avdf, ldf, sidf, ivdf)

    dl1, dl2, dl3, dl4, dl5, dl6 = st.columns(6)
    dl1.download_button("Vouchers CSV", v_csv, "vouchers.csv", "text/csv")
    dl2.download_button("All Vouchers CSV", av_csv, "allvouchers.csv", "text/csv")
    dl3.download_button("Ledgers CSV", l_csv, "ledgers.csv", "text/csv")
    dl4.download_button("Stock Items CSV", si_csv, "stock_items.csv", "text/csv")
    dl5.download_button("Stock Vouchers CSV", iv_csv, "stock_vouchers.csv", "text/csv")
    dl6.download_button("All in Excel", wb_bytes, "tally_export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    dashboard_df = prepare_dashboard_df(vdf)
    tabs = st.tabs(["Dashboard", "Vouchers", "All Vouchers", "Ledgers", "Stock Items", "Stock Vouchers"])

    with tabs[0]:
        left, right = st.columns([1, 1])
        with left:
            selected_types = st.multiselect(
                "Voucher Types",
                options=sorted(dashboard_df["VoucherTypeName"].dropna().unique().tolist()),
                default=sorted(dashboard_df["VoucherTypeName"].dropna().unique().tolist()),
            )
        with right:
            selected_groups = st.multiselect(
                "Primary Groups",
                options=sorted(dashboard_df["PrimaryGroup"].dropna().unique().tolist()),
                default=sorted(dashboard_df["PrimaryGroup"].dropna().unique().tolist()),
            )

        filtered_df = dashboard_df[
            dashboard_df["VoucherTypeName"].isin(selected_types)
            & dashboard_df["PrimaryGroup"].isin(selected_groups)
        ].copy()

        if not filtered_df.empty:
            voucher_chart = filtered_df.groupby("VoucherTypeName")[["DebitAmount", "CreditAmount"]].sum().reset_index()
            bar_options = {
                "tooltip": {"trigger": "axis"},
                "legend": {"data": ["Debit", "Credit"]},
                "xAxis": {"type": "category", "data": voucher_chart["VoucherTypeName"].tolist()},
                "yAxis": {"type": "value"},
                "series": [
                    {"name": "Debit", "type": "bar", "data": voucher_chart["DebitAmount"].tolist()},
                    {"name": "Credit", "type": "bar", "data": voucher_chart["CreditAmount"].tolist()},
                ],
            }
            st.subheader("Voucher Type Totals")
            st_echarts(options=bar_options, height="380px")

            st.subheader("Filtered Voucher Data")
            st.dataframe(filtered_df[VOUCHER_COLUMNS], use_container_width=True, hide_index=True)
        else:
            st.warning("No data for current filters.")

    with tabs[1]:
        st.subheader("Voucher Table")
        st.dataframe(vdf, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader("All Voucher Table")
        st.dataframe(avdf, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader("Ledger Table")
        st.dataframe(ldf, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.subheader("Stock Item Table")
        st.dataframe(sidf, use_container_width=True, hide_index=True)

    with tabs[5]:
        st.subheader("Stock Voucher Table (Inventory Entries)")
        st.dataframe(ivdf, use_container_width=True, hide_index=True)
else:
    st.info("Load data from the sidebar to view the extracted tables.")
