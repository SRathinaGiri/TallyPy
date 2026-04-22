import io
import re
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
    "PartyLedgerName",
    "PartyGSTIN",
    "LedgerGSTIN",
    "VoucherNarration",
    "IsOptional",
    "CompanyName",
    "FromDate",
    "ToDate",
]

LEDGER_COLUMNS = [
    "MasterID",
    "Name",
    "PrimaryGroup",
    "Nature",
    "StartingFrom",
    "CurrencyName",
    "StateName",
    "Parent",
    "PartyGSTIN",
    "OpeningBalance",
    "ClosingBalance",
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


def nature_from_primary_group(primary_group):
    pg = clean_text(primary_group).lower()
    if pg in [
        "current assets", "fixed assets", "investments", "misc. expenses (asset)",
        "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)",
        "stock-in-hand", "sundry debtors"
    ]:
        return "Assets"
    elif pg in [
        "capital account", "current liabilities", "loans (liability)", "suspense account",
        "branch / divisions", "bank od a/c", "duties & taxes", "provisions",
        "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"
    ]:
        return "Liabilities"
    elif pg in ["direct incomes", "indirect incomes", "sales accounts"]:
        return "Income"
    elif pg in ["direct expenses", "indirect expenses", "purchase accounts"]:
        return "Expenses"
    return "Unknown"


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
        "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance</FETCH>"
        "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE>"
        "<COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def build_voucher_request_xml(company, from_date, to_date):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(from_date)}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(to_date)}</SVTODATE>")

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVouchers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<SYSTEM TYPE='Formulae' NAME='IsAccountingVoucher'>"
        "($VoucherTypeName = \"Sales\") OR ($VoucherTypeName = \"Purchase\") OR "
        "($VoucherTypeName = \"Journal\") OR ($VoucherTypeName = \"Receipt\") OR "
        "($VoucherTypeName = \"Payment\") OR ($VoucherTypeName = \"Debit Note\") OR "
        "($VoucherTypeName = \"Credit Note\")"
        "</SYSTEM>"
        "<OBJECT NAME=\"All Ledger Entries\">"
        "<COMPUTE>EntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>"
        "</OBJECT>"
        "<COLLECTION NAME=\"MyVouchers\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, "
        "PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, "
        "AllLedgerEntries.IsDeemedPositive, AllLedgerEntries.EntryLedgerMasterID, "
        "AllLedgerEntries.EntryParentLedger, AllLedgerEntries.EntryPrimaryGroup, "
        "AllLedgerEntries.EntryLedgerGSTIN</FETCH>"
        "<FILTER>IsAccountingVoucher</FILTER>"
        "</COLLECTION>"
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


def build_inventory_entries_request_xml(company, from_date, to_date):
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
        "InventoryEntries.*, AllInventoryEntries.*, InventoryEntriesIn.*, InventoryEntriesOut.*</FETCH>"
        "</COLLECTION>"
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
        # Otherwise fallback to pre-defined logic or empty
        g_info = group_map.get(parent, {})
        nature = g_info.get("Nature", "")
        primary_group = g_info.get("PrimaryGroup", "") or first_non_empty_text(elem, ["PRIMARYGROUP"]) or first_descendant_text(elem, "PRIMARYGROUP")

        row = {
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"),
            "Name": name,
            "PrimaryGroup": primary_group,
            "Nature": nature,
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
        
        # If nature is still empty, try to resolve it from the PrimaryGroup
        if not row["Nature"] and row["PrimaryGroup"]:
             pg_info = group_map.get(row["PrimaryGroup"], {})
             row["Nature"] = pg_info.get("Nature", "")
             
        if not row["Nature"] and row["PrimaryGroup"]:
            # Last fallback for Nature if not in group_map
            row["Nature"] = nature_from_primary_group(row["PrimaryGroup"])

        currency_key = clean_text(row.get("CurrencyFormalNameRaw") or row.get("CurrencyNameRaw")).upper()
        row["CurrencyName"] = CURRENCY_SYMBOL_FALLBACKS.get(currency_key, clean_text(row.get("CurrencySymbolRaw") or row.get("CurrencyOriginalSymbolRaw")))

    output_rows = []
    for row in sorted(ledger_rows, key=lambda item: (int(item.get("MasterID") or 0), item.get("Name", ""))):
        output_rows.append({column: row.get(column, "") for column in LEDGER_COLUMNS})
    return output_rows


def parse_vouchers(root, ledger_meta, company, from_date, to_date, vtype_map=None):
    rows = []
    formatted_from_date = format_tally_date(from_date)
    formatted_to_date = format_tally_date(to_date)
    
    if vtype_map is None:
        vtype_map = {}

    for voucher in root.iter():
        if strip_ns(voucher.tag).upper() != "VOUCHER":
            continue

        voucher_type = direct_child_text(voucher, "VOUCHERTYPENAME")
        base_v_type = vtype_map.get(voucher_type, voucher_type)
        
        # We filter based on the BASE voucher type now
        if base_v_type not in ACCOUNTING_VOUCHER_TYPES:
            continue

        voucher_date = format_tally_date(direct_child_text(voucher, "DATE"))
        voucher_number = direct_child_text(voucher, "VOUCHERNUMBER")
        party_ledger_name = direct_child_text(voucher, "PARTYLEDGERNAME") or "N/A"
        voucher_gstin = direct_child_text(voucher, "PARTYGSTIN")
        voucher_narration = first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
        is_optional = "Yes" if direct_child_text(voucher, "ISOPTIONAL").upper() == "YES" else "No"
        voucher_company = first_non_empty_text(voucher, ["COMPANYNAME", "SVCURRENTCOMPANY"]) or company

        entry_nodes = direct_children(voucher, "ALLLEDGERENTRIES.LIST")
        if not entry_nodes:
            entry_nodes = direct_children(voucher, "LEDGERENTRIES.LIST")

        for entry in entry_nodes:
            ledger_name = direct_child_text(entry, "LEDGERNAME")
            amount_value = to_decimal(direct_child_text(entry, "AMOUNT"))
            is_deemed_positive = direct_child_text(entry, "ISDEEMEDPOSITIVE").upper()

            if not ledger_name or amount_value == 0:
                continue

            base_amount = abs(amount_value)
            signed_amount = base_amount * Decimal("-1") if is_deemed_positive == "YES" else base_amount
            dr_cr = "Dr" if signed_amount < 0 else "Cr"
            debit_amount = base_amount if signed_amount < 0 else Decimal("0.00")
            credit_amount = base_amount if signed_amount > 0 else Decimal("0.00")

            meta = ledger_meta.get(ledger_name, {})
            primary_group = meta.get("PrimaryGroup", "")
            parent_ledger = meta.get("Parent", "")
            ledger_gstin = meta.get("PartyGSTIN", "")
            ledger_master_id = meta.get("MasterID", "")
            nature = meta.get("Nature", "")

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
                "PartyLedgerName": party_ledger_name,
                "PartyGSTIN": voucher_gstin,
                "LedgerGSTIN": ledger_gstin,
                "VoucherNarration": voucher_narration,
                "IsOptional": is_optional,
                "CompanyName": voucher_company,
                "FromDate": formatted_from_date,
                "ToDate": formatted_to_date,
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


def parse_inventory_entries(root, company):
    rows = []
    for voucher in root.iter():
        if strip_ns(voucher.tag).upper() != "VOUCHER":
            continue
        v_type = direct_child_text(voucher, "VOUCHERTYPENAME")
        if "Order" in v_type:
            continue
        v_date = format_tally_date(direct_child_text(voucher, "DATE"))
        v_number = direct_child_text(voucher, "VOUCHERNUMBER")
        v_narration = first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
        v_company = first_non_empty_text(voucher, ["COMPANYNAME", "SVCURRENTCOMPANY"]) or company

        # GREEDY SEARCH: Find ANY tag that contains inventory data
        inv_nodes = [child for child in voucher if "INVENTORYENTRIES" in child.tag.upper()]

        for inv in inv_nodes:
            item_name = direct_child_text(inv, "STOCKITEMNAME")
            if not item_name:
                continue
            
            is_pos_val = direct_child_text(inv, "ISDEEMEDPOSITIVE")
            is_inward = (is_pos_val.upper() == "YES")

            amount_val = abs(to_decimal(direct_child_text(inv, "AMOUNT")))
            qty_text = direct_child_text(inv, "BILLEDQTY")
            rate_text = direct_child_text(inv, "RATE")
            qty_val = abs(to_float(qty_text))
            rate_val = to_float(rate_text)
            
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
                name = direct_child_text(vt, "NAME")
                parent = direct_child_text(vt, "PARENT")
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
        base_types = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra", "Stock Journal"}
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

    voucher_root = parse_xml_root(post_to_tally(url, build_voucher_request_xml(selected_company, from_date, to_date)))
    status = clean_text(first_descendant_text(voucher_root, "STATUS"))
    if status == "0":
        error_text = first_descendant_text(voucher_root, "LINEERROR") or "Tally returned STATUS=0"
        raise ValueError(error_text)

    voucher_rows = parse_vouchers(voucher_root, ledger_meta, selected_company, from_date, to_date, vtype_map)
    
    # New Stock Data
    stock_item_root = parse_xml_root(post_to_tally(url, build_stock_item_request_xml(selected_company)))
    stock_item_rows = parse_stock_items(stock_item_root)
    
    inventory_root = parse_xml_root(post_to_tally(url, build_inventory_entries_request_xml(selected_company, from_date, to_date)))
    inventory_rows = parse_inventory_entries(inventory_root, selected_company)

    voucher_df = pd.DataFrame(voucher_rows)
    ledger_df = pd.DataFrame(ledger_rows)
    stock_item_df = pd.DataFrame(stock_item_rows)
    inventory_df = pd.DataFrame(inventory_rows)

    for df, cols in [(voucher_df, VOUCHER_COLUMNS), (ledger_df, LEDGER_COLUMNS), (stock_item_df, STOCK_ITEM_COLUMNS), (inventory_df, STOCK_VOUCHER_COLUMNS)]:
        for column in cols:
            if column not in df.columns:
                df[column] = ""
        df = df[cols]

    return selected_company, voucher_df, ledger_df, stock_item_df, inventory_df


def to_excel_bytes(voucher_df, ledger_df, stock_item_df, inventory_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        voucher_df.to_excel(writer, index=False, sheet_name="Vouchers")
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

if "voucher_df" not in st.session_state:
    st.session_state.voucher_df = None
    st.session_state.ledger_df = None
    st.session_state.stock_item_df = None
    st.session_state.inventory_df = None
    st.session_state.company_name = ""

if load_btn:
    try:
        company_name, vdf, ldf, sidf, ivdf = load_tally_data(host, port, company, from_date, to_date)
        st.session_state.voucher_df = vdf
        st.session_state.ledger_df = ldf
        st.session_state.stock_item_df = sidf
        st.session_state.inventory_df = ivdf
        st.session_state.company_name = company_name
    except Exception as exc:
        st.error(f"Error fetching data: {exc}")

vdf = st.session_state.voucher_df
ldf = st.session_state.ledger_df
sidf = st.session_state.stock_item_df
ivdf = st.session_state.inventory_df

if vdf is not None:
    st.caption(f"Company: {st.session_state.company_name}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Voucher Rows", len(vdf))
    m2.metric("Ledger Rows", len(ldf))
    m3.metric("Stock Item Rows", len(sidf))
    m4.metric("Stock Voucher Rows", len(ivdf))

    v_csv = vdf.to_csv(index=False).encode("utf-8")
    l_csv = ldf.to_csv(index=False).encode("utf-8")
    si_csv = sidf.to_csv(index=False).encode("utf-8")
    iv_csv = ivdf.to_csv(index=False).encode("utf-8")
    wb_bytes = to_excel_bytes(vdf, ldf, sidf, ivdf)

    dl1, dl2, dl3, dl4, dl5 = st.columns(5)
    dl1.download_button("Vouchers CSV", v_csv, "vouchers.csv", "text/csv")
    dl2.download_button("Ledgers CSV", l_csv, "ledgers.csv", "text/csv")
    dl3.download_button("Stock Items CSV", si_csv, "stock_items.csv", "text/csv")
    dl4.download_button("Stock Vouchers CSV", iv_csv, "stock_vouchers.csv", "text/csv")
    dl5.download_button("All in Excel", wb_bytes, "tally_export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    dashboard_df = prepare_dashboard_df(vdf)
    tabs = st.tabs(["Dashboard", "Vouchers", "Ledgers", "Stock Items", "Stock Vouchers"])

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
        st.subheader("Ledger Table")
        st.dataframe(ldf, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader("Stock Item Table")
        st.dataframe(sidf, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.subheader("Stock Voucher Table (Inventory Entries)")
        st.dataframe(ivdf, use_container_width=True, hide_index=True)
else:
    st.info("Load data from the sidebar to view the extracted tables.")
