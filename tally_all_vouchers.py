import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

HOST = "localhost"
PORT = "9000"
COMPANY = ""
FROM_DATE = ""
TO_DATE = ""

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

PRIMARY_GROUPS = {
    "Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c",
    "Secured Loans", "Unsecured Loans", "Current Liabilities", "Duties & Taxes",
    "Provisions", "Sundry Creditors", "Fixed Assets", "Investments", "Current Assets",
    "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts",
    "Cash-in-hand", "Sundry Debtors", "Misc. Expenses (ASSET)", "Suspense Account",
    "Branch / Divisions", "Sales Accounts", "Purchase Accounts", "Direct Incomes",
    "Indirect Incomes", "Direct Expenses", "Indirect Expenses",
}


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


def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper():
            return clean_text(child.text)
    return ""


def direct_children(elem, local_name):
    return [c for c in list(elem) if strip_ns(c.tag).upper() == local_name.upper()]


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


def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value


def canonical_voucher_type_name(value):
    value = clean_text(value)
    aliases = {
        "rejection in": "Rejections In",
        "rejections in": "Rejections In",
        "rejection out": "Rejections Out",
        "rejections out": "Rejections Out",
    }
    return aliases.get(value.lower(), value)


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
    if pg in ["current assets", "fixed assets", "investments", "misc. expenses (asset)", "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)", "stock-in-hand", "sundry debtors"]:
        return "BS", "Assets"
    if pg in ["capital account", "current liabilities", "loans (liability)", "suspense account", "branch / divisions", "bank od a/c", "duties & taxes", "provisions", "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"]:
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


def post_to_tally(url, xml_text):
    response = requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=120)
    response.raise_for_status()
    return response.text


def parse_xml_root(xml_text):
    cleaned = xml_cleanup(xml_text)
    return ET.fromstring(cleaned.encode("utf-8"))


def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    root = parse_xml_root(post_to_tally(url, xml))
    for cmp in root.iter():
        if strip_ns(cmp.tag).upper() == "COMPANY":
            name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
            start = direct_child_text(cmp, "STARTINGFROM")
            end = direct_child_text(cmp, "ENDINGAT")
            if name:
                return name, start, end
    return "", "", ""


def fetch_tally_metadata(url, company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    voucher_xml = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>AllVTypes</ID></HEADER>"
        f"<BODY><DESC><STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES><TDL><TDLMESSAGE>"
        f"<COLLECTION NAME=\"AllVTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION>"
        f"</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    group_xml = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>AllGroups</ID></HEADER>"
        f"<BODY><DESC><STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES><TDL><TDLMESSAGE>"
        f"<COLLECTION NAME=\"AllGroups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION>"
        f"</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

    vtype_map, group_map = {}, {}
    voucher_root = parse_xml_root(post_to_tally(url, voucher_xml))
    for vt in voucher_root.iter():
        if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
            name = canonical_voucher_type_name(clean_text(vt.get("NAME")) or direct_child_text(vt, "NAME"))
            parent = canonical_voucher_type_name(direct_child_text(vt, "PARENT"))
            if name:
                vtype_map[name] = parent or name

    group_root = parse_xml_root(post_to_tally(url, group_xml))
    for grp in group_root.iter():
        if strip_ns(grp.tag).upper() == "GROUP":
            name = direct_child_text(grp, "NAME")
            parent = direct_child_text(grp, "PARENT")
            nature = direct_child_text(grp, "NATURE")
            primary_group = direct_child_text(grp, "_PRIMARYGROUP")
            if name:
                group_map[name] = {"Parent": parent, "Nature": nature, "PrimaryGroup": primary_group}

    base_types = set(PREDEFINED_VOUCHER_TYPES)
    for _ in range(5):
        for voucher_name, parent_name in list(vtype_map.items()):
            if parent_name and parent_name not in base_types and parent_name in vtype_map:
                vtype_map[voucher_name] = vtype_map[parent_name]
        for _, group_info in group_map.items():
            parent = group_info.get("Parent")
            if parent and not group_info.get("Nature") and parent in group_map:
                group_info["Nature"] = group_map[parent].get("Nature")
            if parent and not group_info.get("PrimaryGroup") and parent in group_map:
                group_info["PrimaryGroup"] = group_map[parent].get("PrimaryGroup")

    return vtype_map, group_map


def parse_ledgers(root, group_map):
    ledger_rows, ledger_lookup = [], {}
    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "LEDGER":
            continue
        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name:
            continue
        parent = direct_child_text(elem, "PARENT")
        group_info = group_map.get(parent, {})
        row = {
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"),
            "Name": name,
            "PrimaryGroup": group_info.get("PrimaryGroup") or first_non_empty_text(elem, ["PRIMARYGROUP"]),
            "Nature": "",
            "NatureOfGroup": group_info.get("Nature", ""),
            "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]),
            "Parent": parent,
            "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]),
        }
        ledger_rows.append(row)
        ledger_lookup[name] = row

    for row in ledger_rows:
        if not row["PrimaryGroup"]:
            row["PrimaryGroup"] = ledger_primary_group(row["Name"], ledger_lookup)
        primary_group = row["PrimaryGroup"]
        if not row["NatureOfGroup"] and primary_group:
            row["NatureOfGroup"] = group_map.get(primary_group, {}).get("Nature", "")
        if row["NatureOfGroup"]:
            lowered = row["NatureOfGroup"].lower()
            if lowered in ["assets", "liabilities"]:
                row["Nature"] = "BS"
            elif lowered in ["income", "expenses"]:
                row["Nature"] = "PL"
        if not row["Nature"] and primary_group:
            row["Nature"], row["NatureOfGroup"] = nature_from_primary_group(primary_group)

    return {row["Name"]: row for row in ledger_rows}


def parse_voucher_rows(root, ledger_meta, company, from_date, to_date, vtype_map):
    rows = []
    formatted_from_date = format_tally_date(from_date)
    formatted_to_date = format_tally_date(to_date)

    for voucher in root.iter():
        if strip_ns(voucher.tag).upper() != "VOUCHER":
            continue

        voucher_type = canonical_voucher_type_name(direct_child_text(voucher, "VOUCHERTYPENAME"))
        base_voucher_type = canonical_voucher_type_name(vtype_map.get(voucher_type, voucher_type))
        voucher_category = voucher_category_from_base_type(base_voucher_type)
        voucher_date = format_tally_date(direct_child_text(voucher, "DATE"))
        voucher_number = direct_child_text(voucher, "VOUCHERNUMBER")
        voucher_narration = first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])

        entries = direct_children(voucher, "ALLLEDGERENTRIES.LIST") or direct_children(voucher, "LEDGERENTRIES.LIST")
        for entry in entries:
            ledger_name = direct_child_text(entry, "LEDGERNAME")
            amount_value = to_decimal(direct_child_text(entry, "AMOUNT"))
            if not ledger_name or amount_value == 0:
                continue

            is_positive = direct_child_text(entry, "ISDEEMEDPOSITIVE").upper() == "YES"
            signed_amount = abs(amount_value) * (Decimal("-1") if is_positive else Decimal("1"))
            meta = ledger_meta.get(ledger_name, {})

            rows.append({
                "Date": voucher_date,
                "VoucherTypeName": voucher_type,
                "BaseVoucherType": base_voucher_type,
                "VoucherNumber": voucher_number,
                "LedgerName": ledger_name,
                "MasterID": direct_child_text(entry, "ENTRYLEDGERMASTERID") or meta.get("MasterID", ""),
                "Amount": float(signed_amount),
                "DrCr": "Dr" if signed_amount < 0 else "Cr",
                "DebitAmount": float(abs(signed_amount)) if signed_amount < 0 else 0.0,
                "CreditAmount": float(abs(signed_amount)) if signed_amount > 0 else 0.0,
                "ParentLedger": direct_child_text(entry, "ENTRYPARENTLEDGER") or meta.get("Parent", ""),
                "PrimaryGroup": direct_child_text(entry, "ENTRYPRIMARYGROUP") or meta.get("PrimaryGroup", ""),
                "Nature": meta.get("Nature", ""),
                "NatureOfGroup": meta.get("NatureOfGroup", ""),
                "PAN": meta.get("PAN", ""),
                "PartyLedgerName": direct_child_text(voucher, "PARTYLEDGERNAME"),
                "PartyGSTIN": direct_child_text(voucher, "PARTYGSTIN"),
                "LedgerGSTIN": direct_child_text(entry, "ENTRYLEDGERGSTIN") or meta.get("PartyGSTIN", ""),
                "VoucherNarration": voucher_narration,
                "IsOptional": "Yes" if direct_child_text(voucher, "ISOPTIONAL").upper() == "YES" else "No",
                "CompanyName": company,
                "FromDate": formatted_from_date,
                "ToDate": formatted_to_date,
                "VoucherCategory": voucher_category,
            })

    return rows


def load_voucher_data(accounting_only=False):
    url = f"http://{HOST}:{PORT}"
    detected_company, detected_from, detected_to = get_company_info(HOST, PORT)
    selected_company = COMPANY or detected_company
    selected_from = FROM_DATE or detected_from
    selected_to = TO_DATE or detected_to

    vtype_map, group_map = fetch_tally_metadata(url, selected_company)
    ledger_request = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyLedgers</ID></HEADER>"
        f"<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(selected_company)}</SVCURRENTCOMPANY></STATICVARIABLES>"
        f"<TDL><TDLMESSAGE><COLLECTION NAME=\"MyLedgers\"><TYPE>Ledger</TYPE>"
        f"<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH>"
        f"<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    ledger_meta = parse_ledgers(parse_xml_root(post_to_tally(url, ledger_request)), group_map)

    voucher_request = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyVouchers</ID></HEADER>"
        f"<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(selected_company)}</SVCURRENTCOMPANY>"
        f"<SVFROMDATE TYPE='Date'>{escape(selected_from)}</SVFROMDATE><SVTODATE TYPE='Date'>{escape(selected_to)}</SVTODATE></STATICVARIABLES>"
        f"<TDL><TDLMESSAGE><OBJECT NAME=\"All Ledger Entries\">"
        f"<COMPUTE>EntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE>"
        f"<COMPUTE>EntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>"
        f"<COMPUTE>EntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>"
        f"<COMPUTE>EntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>"
        f"</OBJECT><COLLECTION NAME=\"MyVouchers\"><TYPE>Voucher</TYPE>"
        f"<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, AllLedgerEntries.IsDeemedPositive, AllLedgerEntries.EntryLedgerMasterID, AllLedgerEntries.EntryParentLedger, AllLedgerEntries.EntryPrimaryGroup, AllLedgerEntries.EntryLedgerGSTIN</FETCH>"
        f"</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    voucher_rows = parse_voucher_rows(parse_xml_root(post_to_tally(url, voucher_request)), ledger_meta, selected_company, selected_from, selected_to, vtype_map)

    df = pd.DataFrame(voucher_rows)
    if df.empty:
        columns = VOUCHER_COLUMNS if accounting_only else ALL_VOUCHER_COLUMNS
        return pd.DataFrame(columns=columns)

    if accounting_only:
        df = df[df["VoucherCategory"] == "Accounting"].copy()
        return df[VOUCHER_COLUMNS]
    return df[ALL_VOUCHER_COLUMNS]


AllVouchers = load_voucher_data(accounting_only=False)
