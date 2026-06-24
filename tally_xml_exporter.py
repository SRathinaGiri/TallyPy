import calendar
import hashlib
import os
import re
import shutil
import threading
import time
import tkinter as tk
import traceback
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape

import pandas as pd
import requests


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
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)
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
        "stock-in-hand", "sundry debtors",
    ]:
        return "BS", "Assets"
    if pg in [
        "capital account", "current liabilities", "loans (liability)", "suspense account",
        "branch / divisions", "bank od a/c", "duties & taxes", "provisions",
        "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans",
    ]:
        return "BS", "Liabilities"
    if pg in ["direct incomes", "indirect incomes", "sales accounts"]:
        return "PL", "Income"
    if pg in ["direct expenses", "indirect expenses", "purchase accounts"]:
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


class ExportCancelled(Exception):
    pass


class IncompleteChunkError(Exception):
    pass


class ExportRunContext:
    def __init__(self):
        self.cancel_event = threading.Event()
        self.session = requests.Session()

    def cancel(self):
        self.cancel_event.set()
        try:
            self.session.close()
        except Exception:
            pass

    def check_cancelled(self):
        if self.cancel_event.is_set():
            raise ExportCancelled("Operation cancelled by user.")


def post_to_tally(url, xml_text, timeout=120, context=None):
    if context:
        context.check_cancelled()
    client = context.session if context else requests
    response = client.post(
        url,
        data=xml_text.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        timeout=timeout,
    )
    if context:
        context.check_cancelled()
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


def voucher_master_id(voucher):
    text = clean_text(voucher.get("MASTERID")) or direct_child_text(voucher, "MASTERID")
    try:
        return int(text)
    except (TypeError, ValueError):
        return None


def chunk_label(chunk):
    label = f"{format_tally_date(chunk['from_date'])} to {format_tally_date(chunk['to_date'])}"
    if chunk.get("master_from") is not None and chunk.get("master_to") is not None:
        label += f", MasterID {chunk['master_from']}-{chunk['master_to']}"
    if chunk.get("expected_count") is not None:
        label += f", expected vouchers {chunk['expected_count']}"
    return label


def make_date_chunk(from_date, to_date, expected_count=None):
    return {
        "from_date": clean_text(from_date),
        "to_date": clean_text(to_date),
        "master_from": None,
        "master_to": None,
        "expected_count": expected_count,
    }


def create_voucher_diagnostics():
    return {
        "headers": {},
        "type_counts": {},
        "category_counts": {},
        "zero_export_reasons": {},
    }


def voucher_header_key(voucher, voucher_date, voucher_type, voucher_number):
    master_id = clean_text(voucher.get("MASTERID")) or direct_child_text(voucher, "MASTERID")
    if master_id:
        return f"master:{master_id}"
    return "|".join(["fallback", voucher_date, voucher_type, voucher_number, clean_text(voucher.get("VCHKEY"))])


def record_voucher_diagnostic(diag, key, voucher_type, category, export_rows, reason):
    if diag is None:
        return
    is_new = key not in diag["headers"]
    header = diag["headers"].setdefault(key, {
        "voucher_type": voucher_type or "Unknown",
        "category": category or "Unknown",
        "rows": 0,
        "reason": "",
    })
    header["rows"] += export_rows
    if not header["reason"] and reason:
        header["reason"] = reason
    if is_new:
        diag["type_counts"][header["voucher_type"]] = diag["type_counts"].get(header["voucher_type"], 0) + 1
        diag["category_counts"][header["category"]] = diag["category_counts"].get(header["category"], 0) + 1


def finalize_voucher_diagnostics(diag):
    if not diag:
        return {}
    headers = list(diag["headers"].values())
    headers_with_rows = sum(1 for item in headers if item["rows"] > 0)
    headers_without_rows = len(headers) - headers_with_rows
    reason_counts = {}
    for item in headers:
        if item["rows"] == 0:
            reason = item.get("reason") or "No exported ledger rows"
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
    return {
        "header_count": len(headers),
        "headers_with_rows": headers_with_rows,
        "headers_without_rows": headers_without_rows,
        "type_counts": dict(sorted(diag["type_counts"].items(), key=lambda item: (-item[1], item[0]))),
        "category_counts": dict(sorted(diag["category_counts"].items(), key=lambda item: (-item[1], item[0]))),
        "reason_counts": dict(sorted(reason_counts.items(), key=lambda item: (-item[1], item[0]))),
    }


def format_top_counts(counts, limit=8):
    if not counts:
        return "none"
    items = list(counts.items())[:limit]
    text = ", ".join(f"{name}: {count}" for name, count in items)
    if len(counts) > limit:
        text += ", ..."
    return text


def verbose_chunk_logging():
    return os.environ.get("TALLYXML_VERBOSE_CHUNKS") == "1"


def should_log_chunk_detail(display_name):
    if verbose_chunk_logging():
        return True
    match = re.search(r"chunk\s+(\d+)/(\d+)", display_name)
    if not match:
        return False
    index = int(match.group(1))
    total = int(match.group(2))
    if total <= 20:
        return True
    return index <= 5 or index == total or index % 50 == 0


def dataframe_value_counts(df, column):
    if df.empty or column not in df.columns:
        return {}
    counts = df[column].fillna("").replace("", "Blank").value_counts()
    return {str(key): int(value) for key, value in counts.items()}


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


def count_voucher_headers(root):
    return sum(1 for elem in root.iter() if is_real_voucher(elem))


def tally_formula_date(value):
    parsed = parse_tally_date_value(value)
    if not parsed:
        return format_tally_date(value)
    return f"{parsed.day}-{calendar.month_abbr[parsed.month]}-{parsed.year}"


def collection_filter_xml(from_date, to_date, master_from=None, master_to=None):
    filters = []
    formulas = []
    if master_from is not None and master_to is not None:
        try:
            master_from = int(master_from)
            master_to = int(master_to)
        except (TypeError, ValueError):
            master_from = None
            master_to = None
        if master_from is not None and master_to is not None:
            filters.append("<FILTER>MasterIdRange</FILTER>")
            formulas.append(
                "<SYSTEM TYPE=\"Formulae\" NAME=\"MasterIdRange\">"
                f"$MasterID &gt;= {master_from} AND $MasterID &lt;= {master_to}"
                "</SYSTEM>"
            )
    else:
        filters.append("<FILTER>DateRange</FILTER>")
        formulas.append(
            "<SYSTEM TYPE=\"Formulae\" NAME=\"DateRange\">"
            f"$Date &gt;= $$Date:\"{escape(tally_formula_date(from_date))}\" "
            f"AND $Date &lt;= $$Date:\"{escape(tally_formula_date(to_date))}\""
            "</SYSTEM>"
        )
    return "".join(filters), "".join(formulas)


def build_voucher_probe_request_xml(company, from_date, to_date):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>")
    collection_filter, filter_formula = collection_filter_xml(from_date, to_date)
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVoucherProbe</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"MyVoucherProbe\"><TYPE>Voucher</TYPE>"
        f"{collection_filter}"
        "<FETCH>Date, MasterID, VoucherTypeName</FETCH>"
        "</COLLECTION>"
        f"{filter_formula}"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def build_data_exception_report_xml(company, report_id):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        f"<TYPE>DATA</TYPE><ID>{escape(report_id)}</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "</DESC></BODY></ENVELOPE>"
    )


def check_tally_data_exceptions(url, company, log=None, context=None):
    report_ids = [
        "Data Exceptions",
        "All Exceptions",
        "Repair/Migrate Exceptions",
        "Import Exceptions",
        "Synchronisation Exceptions",
    ]
    exception_markers = (
        "voucher-related exceptions",
        "master-related exceptions",
        "repair/migrate exceptions",
        "import exceptions",
        "synchronisation exceptions",
        "data exceptions exist",
    )
    unsupported_markers = (
        "could not find",
        "unknown report",
        "invalid report",
        "lineerror",
        "does not exist",
    )
    for report_id in report_ids:
        try:
            if context:
                context.check_cancelled()
            xml = post_to_tally(url, build_data_exception_report_xml(company, report_id), timeout=20, context=context)
        except ExportCancelled:
            raise
        except Exception as exc:
            if log:
                log(f"Data exception preflight: report '{report_id}' not available or timed out: {exc}")
            continue
        plain_text = clean_text(re.sub(r"<[^>]+>", " ", xml_cleanup(xml)))
        lowered = plain_text.lower()
        if any(marker in lowered for marker in unsupported_markers):
            if log:
                log(f"Data exception preflight: report '{report_id}' did not return usable exception data.")
            continue
        if any(marker in lowered for marker in exception_markers) and re.search(r"\b[1-9]\d*\b", lowered):
            summary = plain_text[:500] + ("..." if len(plain_text) > 500 else "")
            if log:
                log(f"Data exception preflight: possible exceptions found in '{report_id}': {summary}")
            return report_id, summary
        if log:
            log(f"Data exception preflight: report '{report_id}' returned no obvious exception markers.")
    return None, ""


def probe_vouchers(url, company, from_date, to_date, log=None, context=None):
    if log:
        log(f"Probe started for {format_tally_date(from_date)} to {format_tally_date(to_date)}")
    started = time.perf_counter()
    root = parse_xml_root(post_to_tally(url, build_voucher_probe_request_xml(company, from_date, to_date), timeout=180, context=context))
    counts_by_month = {}
    counts_by_day = {}
    entries = []
    total = 0
    for voucher in root.iter():
        if context:
            context.check_cancelled()
        if not is_real_voucher(voucher):
            continue
        total += 1
        voucher_date = parse_tally_date_value(direct_child_text(voucher, "DATE"))
        master_id = voucher_master_id(voucher)
        if voucher_date:
            key = (voucher_date.year, voucher_date.month)
            counts_by_month[key] = counts_by_month.get(key, 0) + 1
            day_key = voucher_date.strftime("%Y%m%d")
            counts_by_day[day_key] = counts_by_day.get(day_key, 0) + 1
            if master_id is not None:
                entries.append((voucher_date, master_id))
    if log:
        elapsed = time.perf_counter() - started
        month_text = ", ".join(f"{year}-{month:02d}: {count}" for (year, month), count in sorted(counts_by_month.items()))
        peak_day = max(counts_by_day.items(), key=lambda item: item[1], default=None)
        peak_text = f"{format_tally_date(peak_day[0])}: {peak_day[1]}" if peak_day else "none"
        log(f"Probe completed in {elapsed:.1f}s. Voucher headers found: {total}. MasterIDs found: {len(entries)}. Peak day: {peak_text}. Monthly counts: {month_text or 'none'}")
    return total, counts_by_month, counts_by_day, entries


def build_masterid_chunks(entries, start_date, end_date, max_vouchers=1000):
    if not entries:
        return []
    chunks = []
    master_ids = sorted(master_id for _, master_id in entries if master_id is not None)
    for index in range(0, len(master_ids), max_vouchers):
        batch = master_ids[index:index + max_vouchers]
        chunks.append({
            "from_date": start_date.strftime("%Y%m%d"),
            "to_date": end_date.strftime("%Y%m%d"),
            "master_from": min(batch),
            "master_to": max(batch),
            "expected_count": len(batch),
        })
    return chunks


def build_count_date_chunks(counts_by_day, start_date, end_date, max_vouchers=1000):
    chunks = []
    current_start = None
    current_end = None
    current_count = 0

    for day_text in sorted(counts_by_day):
        day = parse_tally_date_value(day_text)
        if not day:
            continue
        day_count = counts_by_day.get(day_text, 0)
        if current_start is None:
            current_start = day
            current_end = day
            current_count = day_count
            continue
        if current_count and current_count + day_count > max_vouchers:
            chunks.append(make_date_chunk(current_start.strftime("%Y%m%d"), current_end.strftime("%Y%m%d"), current_count))
            current_start = day
            current_end = day
            current_count = day_count
        else:
            current_end = day
            current_count += day_count

    if current_start is not None:
        chunks.append(make_date_chunk(current_start.strftime("%Y%m%d"), current_end.strftime("%Y%m%d"), current_count))

    if not chunks:
        return [make_date_chunk(start_date.strftime("%Y%m%d"), end_date.strftime("%Y%m%d"), 0)]

    first_day = parse_tally_date_value(chunks[0]["from_date"])
    last_day = parse_tally_date_value(chunks[-1]["to_date"])
    if first_day and first_day > start_date:
        chunks.insert(0, make_date_chunk(start_date.strftime("%Y%m%d"), (first_day - timedelta(days=1)).strftime("%Y%m%d"), 0))
    if last_day and last_day < end_date:
        chunks.append(make_date_chunk((last_day + timedelta(days=1)).strftime("%Y%m%d"), end_date.strftime("%Y%m%d"), 0))
    return chunks


def probe_summary_from_counts(total, counts_by_day, entries):
    if not counts_by_day:
        return {
            "total": total,
            "master_count": len(entries),
            "first_voucher_date": "",
            "last_voucher_date": "",
            "peak_day": "",
            "peak_day_count": 0,
        }
    days = sorted(counts_by_day)
    peak_day, peak_count = max(counts_by_day.items(), key=lambda item: item[1])
    return {
        "total": total,
        "master_count": len(entries),
        "first_voucher_date": days[0],
        "last_voucher_date": days[-1],
        "peak_day": peak_day,
        "peak_day_count": peak_count,
    }


def probe_coverage_warning(from_date, to_date, summary):
    requested_from = parse_tally_date_value(from_date)
    requested_to = parse_tally_date_value(to_date)
    first_voucher = parse_tally_date_value(summary.get("first_voucher_date"))
    last_voucher = parse_tally_date_value(summary.get("last_voucher_date"))
    if not requested_from or not requested_to or summary.get("total", 0) == 0:
        return ""
    warnings = []
    if first_voucher and first_voucher > requested_from + timedelta(days=31):
        warnings.append(f"first voucher returned is {first_voucher.isoformat()}, much later than requested from-date {requested_from.isoformat()}")
    if last_voucher and last_voucher < requested_to - timedelta(days=31):
        warnings.append(f"last voucher returned is {last_voucher.isoformat()}, much earlier than requested to-date {requested_to.isoformat()}")
    missing_master_ids = summary.get("total", 0) - summary.get("master_count", 0)
    if missing_master_ids > max(10, int(summary.get("total", 0) * 0.01)):
        warnings.append(f"probe found {summary['total']} voucher headers but {missing_master_ids} did not expose MasterID")
    if not warnings:
        return ""
    return "Potential incomplete voucher coverage from Tally: " + "; ".join(warnings) + ". Check active company, selected period, data exceptions, and rewrite/repair status."


def plan_export_chunks(url, company, from_date, to_date, log=None, context=None, warnings=None):
    start_date = parse_tally_date_value(from_date)
    end_date = parse_tally_date_value(to_date)
    if not start_date or not end_date or start_date > end_date:
        if log:
            log("Could not parse date range for chunk planning; using one request.")
        return [make_date_chunk(from_date, to_date)]
    try:
        total, counts_by_month, counts_by_day, entries = probe_vouchers(url, company, from_date, to_date, log, context)
        summary = probe_summary_from_counts(total, counts_by_day, entries)
        warning = probe_coverage_warning(from_date, to_date, summary)
        if warning:
            if warnings is not None:
                warnings.append(warning)
            if log:
                log(f"WARNING: {warning}")
        target_vouchers_per_chunk = int(os.environ.get("TALLYXML_CHUNK_VOUCHERS", "1000") or "1000")
        target_vouchers_per_chunk = max(250, target_vouchers_per_chunk)
        if total <= target_vouchers_per_chunk:
            chunks = [make_date_chunk(tally_request_date(from_date), tally_request_date(to_date), total)]
            if log:
                log(f"Chunk plan: one full-period request because probe count is <= {target_vouchers_per_chunk}.")
            return chunks
        if entries:
            chunks = build_masterid_chunks(entries, start_date, end_date, max_vouchers=target_vouchers_per_chunk)
            if log:
                first_chunk = chunk_label(chunks[0]) if chunks else "none"
                last_chunk = chunk_label(chunks[-1]) if chunks else "none"
                max_chunk_count = max((chunk.get("expected_count") or 0) for chunk in chunks) if chunks else 0
                log(
                    f"Chunk plan: {len(chunks)} MasterID chunk(s), target <= {target_vouchers_per_chunk} "
                    f"probed vouchers each, max planned {max_chunk_count}. First: {first_chunk}. Last: {last_chunk}."
                )
            return chunks
        if counts_by_day:
            chunks = build_count_date_chunks(counts_by_day, start_date, end_date, max_vouchers=target_vouchers_per_chunk)
            if log:
                first_chunk = chunk_label(chunks[0]) if chunks else "none"
                last_chunk = chunk_label(chunks[-1]) if chunks else "none"
                max_chunk_count = max((chunk.get("expected_count") or 0) for chunk in chunks) if chunks else 0
                log(
                    f"Chunk plan: {len(chunks)} date-range chunk(s), target <= {target_vouchers_per_chunk} "
                    f"probed vouchers each, max planned {max_chunk_count}. First: {first_chunk}. Last: {last_chunk}."
                )
            return chunks
        max_month = max(counts_by_month.values() or [0])
        mode = "monthly" if max_month <= 2000 else "weekly"
    except Exception as exc:
        mode = "monthly"
        if log:
            log(f"Probe failed: {exc}. Falling back to monthly chunks.")
    chunks = [make_date_chunk(start.strftime("%Y%m%d"), end.strftime("%Y%m%d")) for start, end in split_period(start_date, end_date, mode)]
    if log:
        first_chunk = chunk_label(chunks[0]) if chunks else "none"
        last_chunk = chunk_label(chunks[-1]) if chunks else "none"
        log(f"Chunk plan: {len(chunks)} {mode} chunk(s). First: {first_chunk}. Last: {last_chunk}.")
    return chunks


def cache_file_path(company, table_name, from_date, to_date, master_from=None, master_to=None):
    cache_root = os.path.join(os.getcwd(), ".tally_cache")
    key_text = "|".join([
        clean_text(company),
        table_name,
        clean_text(from_date),
        clean_text(to_date),
        clean_text(master_from),
        clean_text(master_to),
        "xml-v5",
    ])
    file_name = hashlib.sha1(key_text.encode("utf-8", errors="ignore")).hexdigest() + ".xml"
    return os.path.join(cache_root, file_name)


def remove_cached_chunk(company, table_name, chunk):
    path = cache_file_path(company, table_name, chunk["from_date"], chunk["to_date"], chunk.get("master_from"), chunk.get("master_to"))
    try:
        if os.path.exists(path):
            os.remove(path)
            return True
    except OSError:
        pass
    return False


def fetch_xml_cached(url, xml_text, company, table_name, chunk, log=None, context=None):
    label = chunk_label(chunk)
    verbose = verbose_chunk_logging()
    if context:
        context.check_cancelled()
    if os.environ.get("TALLYXML_DISABLE_CACHE") == "1":
        if log and verbose:
            log(f"{table_name} {label}: cache disabled; requesting Tally.")
        return post_to_tally(url, xml_text, context=context)
    path = cache_file_path(company, table_name, chunk["from_date"], chunk["to_date"], chunk.get("master_from"), chunk.get("master_to"))
    try:
        if os.path.exists(path):
            if log and verbose:
                size_kb = os.path.getsize(path) / 1024
                log(f"{table_name} {label}: cache hit ({size_kb:.1f} KB).")
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                return handle.read()
    except OSError:
        pass
    last_error = None
    for attempt in range(3):
        try:
            if context:
                context.check_cancelled()
            if log and verbose:
                log(f"{table_name} {label}: requesting Tally (attempt {attempt + 1}/3).")
            started = time.perf_counter()
            xml = post_to_tally(url, xml_text, context=context)
            elapsed = time.perf_counter() - started
            if context:
                context.check_cancelled()
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write(xml)
            except OSError:
                pass
            if log and verbose:
                log(f"{table_name} {label}: received {len(xml) / 1024:.1f} KB in {elapsed:.1f}s and cached.")
            return xml
        except Exception as exc:
            if isinstance(exc, ExportCancelled):
                raise
            last_error = exc
            if log:
                log(f"{table_name} {label}: attempt {attempt + 1}/3 failed: {exc}")
            for _ in range(10 + attempt * 20):
                if context:
                    context.check_cancelled()
                time.sleep(0.1)
    raise last_error


def fetch_chunked_rows(url, company, from_date, to_date, table_name, build_request, parse_chunk, log=None, context=None, warnings=None):
    rows = []
    chunks = plan_export_chunks(url, company, from_date, to_date, log, context, warnings)
    if log:
        log(f"{table_name}: starting {len(chunks)} chunk(s).")
    incomplete_chunks = 0

    def split_masterid_chunk(chunk, display_name, cause=None):
        master_from = chunk.get("master_from")
        master_to = chunk.get("master_to")
        if master_from is None or master_to is None:
            if cause:
                raise cause
            raise ValueError("Cannot split chunk without a MasterID range.")
        master_from = int(master_from)
        master_to = int(master_to)
        if master_to - master_from <= 5:
            message = (
                "Voucher detail export failed for a very small MasterID range "
                f"({master_from}-{master_to}). This usually points to Tally data health issues "
                "or data exceptions in one of these vouchers. Resolve Tally data exceptions/rewrite data, "
                "then retry this utility."
            )
            if cause:
                raise ValueError(message) from cause
            raise ValueError(message)
        midpoint = (master_from + master_to) // 2
        if log:
            log(f"{table_name}: splitting MasterID range {master_from}-{master_to} into {master_from}-{midpoint} and {midpoint + 1}-{master_to}.")
        left_chunk = dict(chunk)
        left_chunk["master_to"] = midpoint
        left_chunk["expected_count"] = None
        right_chunk = dict(chunk)
        right_chunk["master_from"] = midpoint + 1
        right_chunk["expected_count"] = None
        process_chunk(left_chunk, f"{display_name}.1")
        process_chunk(right_chunk, f"{display_name}.2")

    def process_chunk(chunk, display_name):
        nonlocal incomplete_chunks
        if context:
            context.check_cancelled()
        detail_log = log and should_log_chunk_detail(display_name)
        chunk_from = chunk["from_date"]
        chunk_to = chunk["to_date"]
        try:
            if detail_log:
                log(f"{table_name}: {display_name} started ({chunk_label(chunk)}).")
            xml = fetch_xml_cached(url, build_request(company, chunk_from, chunk_to, chunk.get("master_from"), chunk.get("master_to")), company, table_name, chunk, log, context)
            root = parse_xml_root(xml)
            if context:
                context.check_cancelled()
            status = clean_text(first_descendant_text(root, "STATUS"))
            if status == "0":
                error_text = first_descendant_text(root, "LINEERROR") or "Tally returned STATUS=0"
                raise ValueError(error_text)
            detail_header_count = count_voucher_headers(root)
            expected_count = chunk.get("expected_count")
            if chunk.get("master_from") is not None and expected_count is not None and detail_header_count < expected_count:
                raise IncompleteChunkError(
                    f"MasterID-filtered response returned {detail_header_count} detail voucher header(s), "
                    f"but probe saw {expected_count} voucher header(s)."
                )
            chunk_rows = parse_chunk(root, chunk_from, chunk_to)
            rows.extend(chunk_rows)
            if detail_log:
                log(f"{table_name}: {display_name} parsed {len(chunk_rows)} row(s) from {detail_header_count} voucher header(s). Total rows: {len(rows)}.")
        except IncompleteChunkError as exc:
            incomplete_chunks += 1
            if detail_log:
                log(f"{table_name}: {display_name} incomplete: {exc}")
                if remove_cached_chunk(company, table_name, chunk):
                    log(f"{table_name}: removed incomplete cached MasterID response for {chunk_label(chunk)}.")
            else:
                remove_cached_chunk(company, table_name, chunk)
            split_masterid_chunk(chunk, display_name)
        except ExportCancelled:
            raise
        except Exception as exc:
            if log:
                log(f"{table_name}: {display_name} failed: {exc}")
            master_from = chunk.get("master_from")
            master_to = chunk.get("master_to")
            if master_from is not None and master_to is not None:
                split_masterid_chunk(chunk, display_name, exc)
                return
            start_date = parse_tally_date_value(chunk_from)
            end_date = parse_tally_date_value(chunk_to)
            if start_date and end_date and start_date < end_date:
                if log:
                    log(f"{table_name}: splitting failed chunk into daily requests.")
                for day_start, day_end in split_period(start_date, end_date, "daily"):
                    if context:
                        context.check_cancelled()
                    day_from = day_start.strftime("%Y%m%d")
                    day_to = day_end.strftime("%Y%m%d")
                    day_chunk = make_date_chunk(day_from, day_to)
                    xml = fetch_xml_cached(url, build_request(company, day_from, day_to, None, None), company, table_name, day_chunk, log, context)
                    day_rows = parse_chunk(parse_xml_root(xml), day_from, day_to)
                    rows.extend(day_rows)
                    if log:
                        log(f"{table_name}: daily chunk {format_tally_date(day_from)} parsed {len(day_rows)} row(s). Total rows: {len(rows)}.")
            else:
                raise

    for index, chunk in enumerate(chunks, start=1):
        process_chunk(chunk, f"chunk {index}/{len(chunks)}")
    if log:
        if incomplete_chunks:
            log(f"{table_name}: fallback summary: {incomplete_chunks} incomplete chunk(s) were split into smaller MasterID ranges.")
        log(f"{table_name}: completed with {len(rows)} row(s).")
    return rows


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


def build_voucher_request_xml(company, from_date, to_date, master_from=None, master_to=None):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>")
    collection_filter, filter_formula = collection_filter_xml(from_date, to_date, master_from, master_to)

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVouchers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<OBJECT NAME=\"All Ledger Entries\">"
        "<COMPUTE>EntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>"
        "</OBJECT>"
        "<COLLECTION NAME=\"MyVouchers\"><TYPE>Voucher</TYPE>"
        f"{collection_filter}"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, "
        "PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, "
        "AllLedgerEntries.IsDeemedPositive, AllLedgerEntries.EntryLedgerMasterID, "
        "AllLedgerEntries.EntryParentLedger, AllLedgerEntries.EntryPrimaryGroup, "
        "AllLedgerEntries.EntryLedgerGSTIN</FETCH>"
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


def build_inventory_entries_request_xml(company, from_date, to_date, master_from=None, master_to=None):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(tally_request_date(from_date))}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(tally_request_date(to_date))}</SVTODATE>")
    collection_filter, filter_formula = collection_filter_xml(from_date, to_date, master_from, master_to)

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyInventoryVouchers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyInventoryVouchers\"><TYPE>Voucher</TYPE>"
        f"{collection_filter}"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, "
        "InventoryEntries.*, AllInventoryEntries.*, InventoryEntriesIn.*, InventoryEntriesOut.*</FETCH>"
        "</COLLECTION>"
        f"{filter_formula}"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def get_company_info(host, port, context=None):
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
    response = (context.session if context else requests).post(url, data=xml.encode("utf-8"), timeout=10)
    if context:
        context.check_cancelled()
    response.raise_for_status()
    root = ET.fromstring(xml_cleanup(response.text).encode("utf-8"))

    for cmp in root.iter():
        if strip_ns(cmp.tag).upper() == "COMPANY":
            name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
            start = direct_child_text(cmp, "STARTINGFROM")
            end = direct_child_text(cmp, "ENDINGAT")
            if name:
                return name, start, end

    return "", "", ""


def fetch_tally_metadata(url, company, context=None):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")

    vtype_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllVoucherTypes</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllVoucherTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

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
        root_v = parse_xml_root(post_to_tally(url, vtype_xml, context=context))
        for vt in root_v.iter():
            if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
                name = canonical_voucher_type_name(clean_text(vt.get("NAME")) or direct_child_text(vt, "NAME"))
                parent = canonical_voucher_type_name(direct_child_text(vt, "PARENT"))
                if name:
                    vtype_map[name] = parent or name

        root_g = parse_xml_root(post_to_tally(url, group_xml, context=context))
        for grp in root_g.iter():
            if strip_ns(grp.tag).upper() == "GROUP":
                name = direct_child_text(grp, "NAME")
                parent = direct_child_text(grp, "PARENT")
                nature = direct_child_text(grp, "NATURE")
                primary = direct_child_text(grp, "_PRIMARYGROUP")
                if name:
                    group_map[name] = {
                        "Parent": parent,
                        "Nature": nature,
                        "PrimaryGroup": primary,
                    }

        base_types = set(PREDEFINED_VOUCHER_TYPES)
        for _ in range(5):
            for vt_name, parent_name in list(vtype_map.items()):
                if parent_name and parent_name not in base_types and parent_name in vtype_map:
                    vtype_map[vt_name] = vtype_map[parent_name]

        for _ in range(5):
            for _, g_info in group_map.items():
                parent = g_info.get("Parent")
                if parent and not g_info.get("Nature") and parent in group_map:
                    g_info["Nature"] = group_map[parent].get("Nature")
                if parent and not g_info.get("PrimaryGroup") and parent in group_map:
                    g_info["PrimaryGroup"] = group_map[parent].get("PrimaryGroup")
    except ExportCancelled:
        raise
    except Exception:
        pass

    return vtype_map, group_map


def parse_ledgers(root, group_map=None):
    ledger_rows = []
    ledger_lookup = {}
    group_map = group_map or {}

    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "LEDGER":
            continue

        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name:
            continue

        parent = direct_child_text(elem, "PARENT")
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

        if row["NatureOfGroup"]:
            n_val = row["NatureOfGroup"].lower()
            if n_val in ["assets", "liabilities"]:
                row["Nature"] = "BS"
            elif n_val in ["income", "expenses"]:
                row["Nature"] = "PL"

        if not row["Nature"] and pg:
            row["Nature"], row["NatureOfGroup"] = nature_from_primary_group(pg)

        currency_key = clean_text(row.get("CurrencyFormalNameRaw") or row.get("CurrencyNameRaw")).upper()
        row["CurrencyName"] = CURRENCY_SYMBOL_FALLBACKS.get(currency_key, clean_text(row.get("CurrencySymbolRaw") or row.get("CurrencyOriginalSymbolRaw")))

    return [{column: row.get(column, "") for column in LEDGER_COLUMNS} for row in sorted(ledger_rows, key=lambda item: (int(item.get("MasterID") or 0), item.get("Name", "")))]


def parse_vouchers(root, ledger_meta, company, from_date, to_date, vtype_map=None, diagnostics=None):
    rows = []
    formatted_from_date = format_tally_date(from_date)
    formatted_to_date = format_tally_date(to_date)
    vtype_map = vtype_map or {}

    for voucher in root.iter():
        if not is_real_voucher(voucher):
            continue

        voucher_type = canonical_voucher_type_name(direct_child_text(voucher, "VOUCHERTYPENAME"))
        base_v_type = canonical_voucher_type_name(vtype_map.get(voucher_type, voucher_type))
        voucher_category = voucher_category_from_base_type(base_v_type)

        voucher_date = format_tally_date(direct_child_text(voucher, "DATE"))
        voucher_number = direct_child_text(voucher, "VOUCHERNUMBER")
        header_key = voucher_header_key(voucher, voucher_date, voucher_type, voucher_number)
        party_ledger_name = direct_child_text(voucher, "PARTYLEDGERNAME") or "N/A"
        voucher_gstin = direct_child_text(voucher, "PARTYGSTIN")
        voucher_narration = first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
        is_optional = "Yes" if direct_child_text(voucher, "ISOPTIONAL").upper() == "YES" else "No"
        voucher_company = first_non_empty_text(voucher, ["COMPANYNAME", "SVCURRENTCOMPANY"]) or company

        entry_nodes = direct_children(voucher, "ALLLEDGERENTRIES.LIST")
        if not entry_nodes:
            entry_nodes = direct_children(voucher, "LEDGERENTRIES.LIST")

        rows_before = len(rows)
        skipped_missing_ledger = 0
        skipped_zero_amount = 0
        for entry in entry_nodes:
            ledger_name = direct_child_text(entry, "LEDGERNAME")
            amount_value = to_decimal(direct_child_text(entry, "AMOUNT"))
            is_deemed_positive = direct_child_text(entry, "ISDEEMEDPOSITIVE").upper()

            if not ledger_name:
                skipped_missing_ledger += 1
                continue
            if amount_value == 0:
                skipped_zero_amount += 1
                continue

            base_amount = abs(amount_value)
            signed_amount = base_amount * Decimal("-1") if is_deemed_positive == "YES" else base_amount
            dr_cr = "Dr" if signed_amount < 0 else "Cr"
            debit_amount = base_amount if signed_amount < 0 else Decimal("0.00")
            credit_amount = base_amount if signed_amount > 0 else Decimal("0.00")

            meta = ledger_meta.get(ledger_name, {})
            ledger_master_id = meta.get("MasterID", "")
            primary_group = meta.get("PrimaryGroup", "")
            parent_ledger = meta.get("Parent", "")
            ledger_gstin = meta.get("PartyGSTIN", "")
            pan = meta.get("PAN", "")
            nature = meta.get("Nature", "")
            nature_of_group = meta.get("NatureOfGroup", "")

            entry_level_master_id = direct_child_text(entry, "ENTRYLEDGERMASTERID")
            entry_level_parent = direct_child_text(entry, "ENTRYPARENTLEDGER")
            entry_level_primary_group = direct_child_text(entry, "ENTRYPRIMARYGROUP")
            entry_level_gstin = direct_child_text(entry, "ENTRYLEDGERGSTIN")

            if entry_level_master_id:
                ledger_master_id = entry_level_master_id
            if entry_level_parent:
                parent_ledger = entry_level_parent
            if entry_level_primary_group:
                primary_group = entry_level_primary_group
            if entry_level_gstin:
                ledger_gstin = entry_level_gstin
            if not nature and primary_group:
                nature, nature_of_group = nature_from_primary_group(primary_group)

            rows.append({
                "Date": voucher_date,
                "VoucherTypeName": voucher_type,
                "BaseVoucherType": base_v_type,
                "VoucherNumber": voucher_number,
                "LedgerName": ledger_name,
                "MasterID": ledger_master_id,
                "Amount": float(signed_amount),
                "DrCr": dr_cr,
                "DebitAmount": float(debit_amount),
                "CreditAmount": float(credit_amount),
                "ParentLedger": parent_ledger,
                "PrimaryGroup": primary_group,
                "Nature": nature,
                "NatureOfGroup": nature_of_group,
                "PAN": pan,
                "PartyLedgerName": party_ledger_name,
                "PartyGSTIN": voucher_gstin,
                "LedgerGSTIN": ledger_gstin,
                "VoucherNarration": voucher_narration,
                "IsOptional": is_optional,
                "CompanyName": voucher_company,
                "FromDate": formatted_from_date,
                "ToDate": formatted_to_date,
                "VoucherCategory": voucher_category,
            })

        exported_count = len(rows) - rows_before
        reason = ""
        if exported_count == 0:
            if not entry_nodes:
                reason = "No ledger entry nodes"
            elif skipped_zero_amount and skipped_zero_amount == len(entry_nodes):
                reason = "Only zero-amount ledger entries"
            elif skipped_missing_ledger and skipped_missing_ledger == len(entry_nodes):
                reason = "Ledger name missing in entries"
            else:
                reason = "No exported ledger rows"
            if voucher_category != "Accounting":
                rows.append({
                    "Date": voucher_date,
                    "VoucherTypeName": voucher_type,
                    "BaseVoucherType": base_v_type,
                    "VoucherNumber": voucher_number,
                    "LedgerName": "",
                    "MasterID": "",
                    "Amount": 0.0,
                    "DrCr": "",
                    "DebitAmount": 0.0,
                    "CreditAmount": 0.0,
                    "ParentLedger": "",
                    "PrimaryGroup": "",
                    "Nature": "",
                    "NatureOfGroup": "",
                    "PAN": "",
                    "PartyLedgerName": party_ledger_name,
                    "PartyGSTIN": voucher_gstin,
                    "LedgerGSTIN": "",
                    "VoucherNarration": voucher_narration,
                    "IsOptional": is_optional,
                    "CompanyName": voucher_company,
                    "FromDate": formatted_from_date,
                    "ToDate": formatted_to_date,
                    "VoucherCategory": voucher_category,
                })
                exported_count = 1
        record_voucher_diagnostic(diagnostics, header_key, voucher_type, voucher_category, exported_count, reason)

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


def parse_inventory_entries(root, company):
    rows = []
    for voucher in root.iter():
        if not is_real_voucher(voucher):
            continue
        v_type = direct_child_text(voucher, "VOUCHERTYPENAME")
        if "Order" in v_type:
            continue
        v_date = format_tally_date(direct_child_text(voucher, "DATE"))
        v_number = direct_child_text(voucher, "VOUCHERNUMBER")
        v_narration = first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
        v_company = first_non_empty_text(voucher, ["COMPANYNAME", "SVCURRENTCOMPANY"]) or company

        inv_nodes = [child for child in voucher if "INVENTORYENTRIES" in child.tag.upper()]
        for inv in inv_nodes:
            item_name = direct_child_text(inv, "STOCKITEMNAME")
            if not item_name:
                continue

            is_pos_val = direct_child_text(inv, "ISDEEMEDPOSITIVE")
            is_inward = is_pos_val.upper() == "YES"
            amount_val = abs(to_decimal(direct_child_text(inv, "AMOUNT")))
            qty_val = abs(to_float(direct_child_text(inv, "BILLEDQTY")))
            rate_val = to_float(direct_child_text(inv, "RATE"))

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
                "StockItemName": item_name.strip(),
                "BilledQty": qty_val if is_inward else -qty_val,
                "Rate": rate_val,
                "Amount": float(amount_val if is_inward else -amount_val),
                "GodownName": godown,
                "BatchName": batch,
                "VoucherNarration": v_narration,
                "CompanyName": v_company,
            })
    return rows


def load_tally_data(host, port, company, from_date, to_date, log=None, context=None):
    load_started = time.perf_counter()
    warnings = []
    url = f"http://{host}:{port}"
    selected_company = clean_text(company)
    from_date = clean_text(from_date)
    to_date = clean_text(to_date)
    if log:
        log(f"Load started. Target: {url}. Input company='{selected_company or '(auto)'}', from='{from_date or '(auto)'}', to='{to_date or '(auto)'}'.")
    if context:
        context.check_cancelled()

    if not selected_company or not from_date or not to_date:
        if log:
            log("Company/date input incomplete; requesting active company metadata.")
        cmp_name, cmp_start, cmp_end = get_company_info(host, port, context)
        if log:
            log(f"Company metadata received: company='{cmp_name}', starting='{format_tally_date(cmp_start)}', ending='{format_tally_date(cmp_end)}'.")
        if not selected_company:
            selected_company = cmp_name
        if not from_date:
            from_date = cmp_start
        if not to_date:
            to_date = cmp_end

    if log:
        log(f"Using company='{selected_company}', period={format_tally_date(from_date)} to {format_tally_date(to_date)}.")
        log("Running best-effort Tally data exception preflight.")
    exception_report, exception_summary = check_tally_data_exceptions(url, selected_company, log, context)
    if exception_report:
        raise ValueError(
            "Tally data exceptions appear to exist in "
            f"'{exception_report}'. Resolve Tally data exceptions before running voucher export. "
            f"Details: {exception_summary}"
        )

    if log:
        log("Fetching voucher type and group metadata.")
    started = time.perf_counter()
    vtype_map, group_map = fetch_tally_metadata(url, selected_company, context)
    if log:
        log(f"Metadata received in {time.perf_counter() - started:.1f}s. Voucher types: {len(vtype_map)}. Groups: {len(group_map)}.")

    if log:
        log("Fetching ledgers.")
    started = time.perf_counter()
    ledger_root = parse_xml_root(post_to_tally(url, build_ledger_request_xml(selected_company), context=context))
    ledger_rows = parse_ledgers(ledger_root, group_map)
    ledger_meta = {row["Name"]: row for row in ledger_rows}
    if log:
        log(f"Ledgers parsed in {time.perf_counter() - started:.1f}s. Rows: {len(ledger_rows)}.")

    if log:
        log("Fetching accounting/all voucher rows with probe-based chunking.")
    voucher_diagnostics = create_voucher_diagnostics()
    voucher_rows = fetch_chunked_rows(
        url,
        selected_company,
        from_date,
        to_date,
        "vouchers",
        build_voucher_request_xml,
        lambda root, chunk_from, chunk_to: parse_vouchers(root, ledger_meta, selected_company, chunk_from, chunk_to, vtype_map, voucher_diagnostics),
        log,
        context,
        warnings,
    )
    voucher_diagnostic_summary = finalize_voucher_diagnostics(voucher_diagnostics)
    if log:
        log(
            "Voucher diagnostics: "
            f"headers in detail XML: {voucher_diagnostic_summary.get('header_count', 0)}; "
            f"headers with exported rows: {voucher_diagnostic_summary.get('headers_with_rows', 0)}; "
            f"headers without exported rows: {voucher_diagnostic_summary.get('headers_without_rows', 0)}."
        )
        log(f"Voucher diagnostics by category: {format_top_counts(voucher_diagnostic_summary.get('category_counts', {}))}")
        log(f"Voucher diagnostics by type: {format_top_counts(voucher_diagnostic_summary.get('type_counts', {}))}")
        log(f"Voucher diagnostics no-row reasons: {format_top_counts(voucher_diagnostic_summary.get('reason_counts', {}))}")

    if log:
        log("Fetching stock items.")
    started = time.perf_counter()
    stock_item_root = parse_xml_root(post_to_tally(url, build_stock_item_request_xml(selected_company), context=context))
    stock_item_rows = parse_stock_items(stock_item_root)
    if log:
        log(f"Stock items parsed in {time.perf_counter() - started:.1f}s. Rows: {len(stock_item_rows)}.")

    if log:
        log("Fetching stock voucher rows with probe-based chunking.")
    inventory_rows = fetch_chunked_rows(
        url,
        selected_company,
        from_date,
        to_date,
        "inventory",
        build_inventory_entries_request_xml,
        lambda root, chunk_from, chunk_to: parse_inventory_entries(root, selected_company),
        log,
        context,
        warnings,
    )

    if log:
        log("Building final dataframes.")
    try:
        all_voucher_df = pd.DataFrame(voucher_rows)
        if all_voucher_df.empty:
            all_voucher_df = pd.DataFrame(columns=ALL_VOUCHER_COLUMNS)
            voucher_df = pd.DataFrame(columns=VOUCHER_COLUMNS)
        else:
            voucher_df = all_voucher_df[all_voucher_df["VoucherCategory"] == "Accounting"].copy()
        ledger_df = pd.DataFrame(ledger_rows)
        stock_item_df = pd.DataFrame(stock_item_rows)
        inventory_df = pd.DataFrame(inventory_rows)

        formatted_from = format_tally_date(from_date)
        formatted_to = format_tally_date(to_date)

        raw_dfs = {
            "voucher_df": (voucher_df, VOUCHER_COLUMNS),
            "all_voucher_df": (all_voucher_df, ALL_VOUCHER_COLUMNS),
            "ledger_df": (ledger_df, LEDGER_COLUMNS),
            "stock_item_df": (stock_item_df, STOCK_ITEM_COLUMNS),
            "inventory_df": (inventory_df, STOCK_VOUCHER_COLUMNS),
        }
        final_dfs = {}
        for name, (df, columns) in raw_dfs.items():
            if log:
                log(f"Finalizing dataframe '{name}' with {len(df)} row(s).")
            df = df.copy()
            df["CompanyName"] = selected_company
            df["FromDate"] = formatted_from
            df["ToDate"] = formatted_to
            for column in columns:
                if column not in df.columns:
                    df[column] = ""
            final_dfs[name] = df.loc[:, columns]
    except Exception:
        if log:
            log("Dataframe assembly failed:\n" + traceback.format_exc())
        raise

    if log:
        category_counts = dataframe_value_counts(final_dfs["all_voucher_df"], "VoucherCategory")
        log(
            "Load completed in "
            f"{time.perf_counter() - load_started:.1f}s. "
            f"Accounting voucher rows: {len(final_dfs['voucher_df'])}. All voucher rows: {len(final_dfs['all_voucher_df'])}. "
            f"Ledgers: {len(final_dfs['ledger_df'])}. Stock items: {len(final_dfs['stock_item_df'])}. Stock vouchers: {len(final_dfs['inventory_df'])}."
        )
        log(f"All voucher rows by category: {format_top_counts(category_counts)}")

    return {
        "company_name": selected_company,
        "from_date": from_date,
        "to_date": to_date,
        "voucher_df": final_dfs["voucher_df"],
        "all_voucher_df": final_dfs["all_voucher_df"],
        "ledger_df": final_dfs["ledger_df"],
        "stock_item_df": final_dfs["stock_item_df"],
        "inventory_df": final_dfs["inventory_df"],
        "warnings": warnings,
        "voucher_diagnostics": voucher_diagnostic_summary,
    }


class TallyDesktopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tally XML Desktop Exporter")
        self.root.geometry("1260x780")

        self.company_var = tk.StringVar()
        self.from_date_var = tk.StringVar()
        self.to_date_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.stats_var = tk.StringVar(value="Voucher Rows: 0 | All Voucher Rows: 0 | Ledgers: 0 | Stock Items: 0 | Stock Voucher Rows: 0")

        self.tables = {
            "voucher_df": pd.DataFrame(columns=VOUCHER_COLUMNS),
            "all_voucher_df": pd.DataFrame(columns=ALL_VOUCHER_COLUMNS),
            "ledger_df": pd.DataFrame(columns=LEDGER_COLUMNS),
            "stock_item_df": pd.DataFrame(columns=STOCK_ITEM_COLUMNS),
            "inventory_df": pd.DataFrame(columns=STOCK_VOUCHER_COLUMNS),
        }

        self.treeviews = {}
        self.log_file_path = self._create_log_file_path()
        self.current_context = None
        self.is_closing = False
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._log(f"Log file: {self.log_file_path}")

    def _create_log_file_path(self):
        log_dir = os.path.join(os.getcwd(), ".tally_logs")
        os.makedirs(log_dir, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(log_dir, f"tally_xml_exporter_{stamp}.log")

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top = ttk.Frame(self.root, padding=12)
        top.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        connection = ttk.LabelFrame(top, text="Connection", padding=10)
        connection.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        details = ttk.LabelFrame(top, text="Selection", padding=10)
        details.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ttk.Label(connection, text="Host").grid(row=0, column=0, sticky="w")
        ttk.Label(connection, text="Port").grid(row=1, column=0, sticky="w")
        ttk.Label(connection, text="Company").grid(row=2, column=0, sticky="w")
        ttk.Label(connection, text="From Date").grid(row=3, column=0, sticky="w")
        ttk.Label(connection, text="To Date").grid(row=4, column=0, sticky="w")

        self.host_entry = ttk.Entry(connection, width=28)
        self.host_entry.insert(0, "localhost")
        self.host_entry.grid(row=0, column=1, sticky="ew", pady=2)

        self.port_entry = ttk.Entry(connection, width=28)
        self.port_entry.insert(0, "9000")
        self.port_entry.grid(row=1, column=1, sticky="ew", pady=2)

        self.company_entry = ttk.Entry(connection, width=28)
        self.company_entry.grid(row=2, column=1, sticky="ew", pady=2)

        self.from_entry = ttk.Entry(connection, width=28)
        self.from_entry.grid(row=3, column=1, sticky="ew", pady=2)

        self.to_entry = ttk.Entry(connection, width=28)
        self.to_entry.grid(row=4, column=1, sticky="ew", pady=2)

        button_row = ttk.Frame(connection)
        button_row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.connect_button = ttk.Button(button_row, text="Connect", command=self.connect_tally)
        self.connect_button.pack(side="left")
        self.load_button = ttk.Button(button_row, text="Load Tables", command=self.load_tables)
        self.load_button.pack(side="left", padx=6)
        self.cancel_button = ttk.Button(button_row, text="Cancel", command=self.cancel_current_operation, state="disabled")
        self.cancel_button.pack(side="left")
        self.progress = ttk.Progressbar(button_row, mode="indeterminate", length=180)
        self.progress.pack(side="left", padx=(10, 0))

        maintenance_row = ttk.Frame(connection)
        maintenance_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.clear_cache_button = ttk.Button(maintenance_row, text="Clear Cache", command=self.clear_cache)
        self.clear_cache_button.pack(side="left")
        self.clear_logs_button = ttk.Button(maintenance_row, text="Clear Logs", command=self.clear_logs)
        self.clear_logs_button.pack(side="left", padx=6)

        connection.columnconfigure(1, weight=1)

        ttk.Label(details, text="Detected Company").grid(row=0, column=0, sticky="w")
        ttk.Entry(details, textvariable=self.company_var, state="readonly", width=36).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Label(details, text="Detected From").grid(row=1, column=0, sticky="w")
        ttk.Entry(details, textvariable=self.from_date_var, state="readonly", width=36).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Label(details, text="Detected To").grid(row=2, column=0, sticky="w")
        ttk.Entry(details, textvariable=self.to_date_var, state="readonly", width=36).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(details, text="Status").grid(row=3, column=0, sticky="w")
        ttk.Entry(details, textvariable=self.status_var, state="readonly", width=36).grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Label(details, textvariable=self.stats_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 4))

        export_row = ttk.Frame(details)
        export_row.grid(row=5, column=0, columnspan=2, sticky="ew")
        ttk.Button(export_row, text="Export All CSVs", command=self.export_all_csvs).pack(side="left")
        ttk.Button(export_row, text="Export Vouchers", command=lambda: self.export_single_csv("voucher_df", "vouchers.csv")).pack(side="left", padx=4)
        ttk.Button(export_row, text="Export All Vouchers", command=lambda: self.export_single_csv("all_voucher_df", "allvouchers.csv")).pack(side="left", padx=4)
        ttk.Button(export_row, text="Export Ledgers", command=lambda: self.export_single_csv("ledger_df", "ledgers.csv")).pack(side="left", padx=4)
        ttk.Button(export_row, text="Export Stock Items", command=lambda: self.export_single_csv("stock_item_df", "stock_items.csv")).pack(side="left", padx=4)
        ttk.Button(export_row, text="Export Stock Vouchers", command=lambda: self.export_single_csv("inventory_df", "stock_vouchers.csv")).pack(side="left", padx=4)

        details.columnconfigure(1, weight=1)

        body = ttk.Panedwindow(self.root, orient="vertical")
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        notebook_frame = ttk.Frame(body)
        log_frame = ttk.LabelFrame(body, text="Log", padding=8)
        body.add(notebook_frame, weight=5)
        body.add(log_frame, weight=2)

        notebook = ttk.Notebook(notebook_frame)
        notebook.pack(fill="both", expand=True)

        for title, key, columns in [
            ("Vouchers", "voucher_df", VOUCHER_COLUMNS),
            ("All Vouchers", "all_voucher_df", ALL_VOUCHER_COLUMNS),
            ("Ledgers", "ledger_df", LEDGER_COLUMNS),
            ("Stock Items", "stock_item_df", STOCK_ITEM_COLUMNS),
            ("Stock Vouchers", "inventory_df", STOCK_VOUCHER_COLUMNS),
        ]:
            frame = ttk.Frame(notebook, padding=6)
            notebook.add(frame, text=title)
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)
            tree = ttk.Treeview(frame, show="headings")
            vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            self._configure_tree(tree, columns)
            self.treeviews[key] = tree

        self.log_text = tk.Text(log_frame, height=10, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def _configure_tree(self, tree, columns):
        tree["columns"] = columns
        for column in columns:
            tree.heading(column, text=column)
            tree.column(column, width=130, anchor="w", stretch=True)

    def _log(self, message):
        line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{line}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as handle:
                handle.write(f"{line}\n")
        except OSError:
            pass

    def _log_from_worker(self, message):
        self._call_ui(lambda msg=message: self._log(msg))

    def _call_ui(self, callback):
        if self.is_closing:
            return
        try:
            self.root.after(0, callback)
        except tk.TclError:
            pass

    def clear_cache(self):
        cache_dir = os.path.join(os.getcwd(), ".tally_cache")
        try:
            if os.path.isdir(cache_dir):
                file_count = sum(len(files) for _, _, files in os.walk(cache_dir))
                shutil.rmtree(cache_dir)
            else:
                file_count = 0
            self._log(f"Cache cleared. Removed {file_count} cached XML file(s).")
            messagebox.showinfo("Cache Cleared", f"Removed {file_count} cached XML file(s).")
        except Exception as exc:
            self._handle_error(f"Clear cache failed: {exc}")

    def clear_logs(self):
        log_dir = os.path.join(os.getcwd(), ".tally_logs")
        deleted = 0
        try:
            if os.path.isdir(log_dir):
                for name in os.listdir(log_dir):
                    if not name.lower().endswith(".log"):
                        continue
                    path = os.path.join(log_dir, name)
                    if os.path.isfile(path):
                        os.remove(path)
                        deleted += 1
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")
            self.log_file_path = self._create_log_file_path()
            self._log(f"Logs cleared. Removed {deleted} log file(s). New log file: {self.log_file_path}")
            messagebox.showinfo("Logs Cleared", f"Removed {deleted} log file(s).")
        except Exception as exc:
            self._handle_error(f"Clear logs failed: {exc}")

    def _set_status(self, message):
        self.status_var.set(message)
        self._log(message)

    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for entry in [self.host_entry, self.port_entry, self.company_entry, self.from_entry, self.to_entry]:
            entry.configure(state=state)
        self.connect_button.configure(state=state)
        self.load_button.configure(state=state)
        self.clear_cache_button.configure(state=state)
        self.clear_logs_button.configure(state=state)
        self.cancel_button.configure(state="normal" if busy else "disabled")
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()

    def cancel_current_operation(self):
        context = self.current_context
        if context:
            self._log("Cancel requested. Closing active HTTP session and stopping after current request returns.")
            context.cancel()
            self.status_var.set("Cancelling...")

    def _on_close(self):
        self.is_closing = True
        self.cancel_current_operation()
        self.root.destroy()

    def _run_background(self, target):
        self._set_busy(True)
        thread = threading.Thread(target=target, daemon=True)
        thread.start()

    def connect_tally(self):
        def work():
            try:
                host = self.host_entry.get().strip() or "localhost"
                port = self.port_entry.get().strip() or "9000"
                company_name, start_date, end_date = get_company_info(host, port)
                if not company_name:
                    raise ValueError("Active company could not be detected from Tally.")
                self._call_ui(lambda: self._apply_connection(company_name, start_date, end_date))
            except Exception as exc:
                message = f"Connect failed: {exc}"
                self._call_ui(lambda message=message: self._handle_error(message))
            finally:
                self._call_ui(lambda: self._set_busy(False))

        self._set_status("Connecting to Tally...")
        self._run_background(work)

    def _apply_connection(self, company_name, start_date, end_date):
        self.company_var.set(company_name)
        self.from_date_var.set(format_tally_date(start_date))
        self.to_date_var.set(format_tally_date(end_date))
        if not self.company_entry.get().strip():
            self.company_entry.insert(0, company_name)
        if not self.from_entry.get().strip():
            self.from_entry.insert(0, start_date)
        if not self.to_entry.get().strip():
            self.to_entry.insert(0, end_date)
        self._set_status(f"Connected to {company_name}")

    def load_tables(self):
        def work():
            context = ExportRunContext()
            self.current_context = context
            try:
                host = self.host_entry.get().strip() or "localhost"
                port = self.port_entry.get().strip() or "9000"
                company = self.company_entry.get().strip()
                from_date = self.from_entry.get().strip()
                to_date = self.to_entry.get().strip()
                data = load_tally_data(host, port, company, from_date, to_date, log=self._log_from_worker, context=context)
                self._call_ui(lambda: self._apply_loaded_data(data))
            except ExportCancelled as exc:
                message = str(exc)
                self._call_ui(lambda message=message: self._set_status(message))
            except Exception as exc:
                message = f"Load failed: {exc}"
                self._call_ui(lambda message=message: self._handle_error(message))
            finally:
                try:
                    context.session.close()
                except Exception:
                    pass
                if self.current_context is context:
                    self.current_context = None
                self._call_ui(lambda: self._set_busy(False))

        self._set_status("Loading tables from Tally...")
        self._run_background(work)

    def _apply_loaded_data(self, data):
        self.company_var.set(data["company_name"])
        self.from_date_var.set(format_tally_date(data["from_date"]))
        self.to_date_var.set(format_tally_date(data["to_date"]))

        for key in self.tables:
            self.tables[key] = data[key]
            self._populate_tree(self.treeviews[key], data[key])

        self.stats_var.set(
            f"Voucher Rows: {len(data['voucher_df'])} | All Voucher Rows: {len(data['all_voucher_df'])} | Ledgers: {len(data['ledger_df'])} | "
            f"Stock Items: {len(data['stock_item_df'])} | Stock Voucher Rows: {len(data['inventory_df'])}"
        )
        if data.get("warnings"):
            unique_warnings = list(dict.fromkeys(data["warnings"]))
            warning_text = "\n\n".join(unique_warnings)
            diag = data.get("voucher_diagnostics") or {}
            if diag:
                warning_text += (
                    "\n\nVoucher diagnostics:\n"
                    f"Headers in detail XML: {diag.get('header_count', 0)}\n"
                    f"Headers with exported rows: {diag.get('headers_with_rows', 0)}\n"
                    f"Headers without exported rows: {diag.get('headers_without_rows', 0)}\n"
                    f"No-row reasons: {format_top_counts(diag.get('reason_counts', {}), limit=5)}"
                )
            self._set_status(f"Loaded with warnings: {unique_warnings[0]}")
            messagebox.showwarning("Tally Data Warning", warning_text)
        else:
            self._set_status(f"Loaded data for {data['company_name']}")

    def _populate_tree(self, tree, df):
        tree.delete(*tree.get_children())
        if df.empty:
            return
        preview = df.head(500).fillna("")
        for row in preview.itertuples(index=False, name=None):
            tree.insert("", "end", values=row)

    def _require_data(self):
        if all(df.empty for df in self.tables.values()):
            messagebox.showwarning("No data", "Load data from Tally before exporting.")
            return False
        return True

    def export_single_csv(self, key, default_name):
        if self.tables[key].empty:
            messagebox.showwarning("No data", "This table is empty.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=default_name,
        )
        if not path:
            return
        self.tables[key].to_csv(path, index=False, encoding="utf-8-sig")
        self._set_status(f"Saved {os.path.basename(path)}")
        messagebox.showinfo("Export complete", f"Saved:\n{path}")

    def export_all_csvs(self):
        if not self._require_data():
            return
        folder = filedialog.askdirectory()
        if not folder:
            return

        file_map = {
            "voucher_df": "vouchers.csv",
            "all_voucher_df": "allvouchers.csv",
            "ledger_df": "ledgers.csv",
            "stock_item_df": "stock_items.csv",
            "inventory_df": "stock_vouchers.csv",
        }
        for key, filename in file_map.items():
            self.tables[key].to_csv(os.path.join(folder, filename), index=False, encoding="utf-8-sig")

        self._set_status(f"Exported all CSVs to {folder}")
        messagebox.showinfo("Export complete", f"All CSVs saved to:\n{folder}")

    def _handle_error(self, message):
        self._set_status(message)
        messagebox.showerror("Tally XML Exporter", message)


if __name__ == "__main__":
    root = tk.Tk()
    app = TallyDesktopApp(root)
    root.mainloop()
