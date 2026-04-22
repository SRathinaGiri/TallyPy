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
CHUNK_DAYS = 31

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
    "Fixed Assets", "Investments",
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

TDL_OUTPUT_COLUMNS = [
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


def iterate_date_chunks(from_date, to_date, chunk_days):
    start = datetime.strptime(from_date, "%Y%m%d").date()
    end = datetime.strptime(to_date, "%Y%m%d").date()
    current = start
    while current <= end:
        chunk_end = min(current + timedelta(days=chunk_days - 1), end)
        yield current.strftime("%Y%m%d"), chunk_end.strftime("%Y%m%d")
        current = chunk_end + timedelta(days=1)


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


def fetch_tally_metadata(url, company):
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
        resp_v = post_to_tally(url, vtype_xml)
        root_v = parse_xml_root(resp_v)
        for vt in root_v.iter():
            if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
                name = direct_child_text(vt, "NAME")
                parent = direct_child_text(vt, "PARENT")
                if name:
                    vtype_map[name] = parent or name
        
        # Resolve Voucher Types recursively
        base_types = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra", "Stock Journal"}
        for _ in range(5): 
            for vt_name, parent_name in vtype_map.items():
                if parent_name and parent_name not in base_types and parent_name in vtype_map:
                    vtype_map[vt_name] = vtype_map[parent_name]

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
        "<FETCH>Name, Parent, GSTIN, PartyGSTIN, MasterID, IncomeTaxNumber</FETCH>"
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


def parse_ledgers(root, group_map=None):
    ledger_meta = {}
    if group_map is None:
        group_map = {}
        
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
            "Parent": parent,
            "GSTIN": first_non_empty_text(elem, ["GSTIN", "PARTYGSTIN"]) or first_descendant_text(elem, "PARTYGSTIN"),
            "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]) or first_descendant_text(elem, "INCOMETAXNUMBER"),
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"),
            "Nature": "",
            "NatureOfGroup": nature_of_group,
            "PrimaryGroup": primary_group
        }
        existing = ledger_meta.get(name)
        if existing is None:
            ledger_meta[name] = row
            continue

        try:
            existing_id = int(existing.get("MasterID") or 0)
        except ValueError:
            existing_id = 0
        try:
            new_id = int(row.get("MasterID") or 0)
        except ValueError:
            new_id = 0

        if new_id >= existing_id:
            ledger_meta[name] = row

    for name in list(ledger_meta):
        if not ledger_meta[name]["PrimaryGroup"]:
            ledger_meta[name]["PrimaryGroup"] = ledger_primary_group(name, ledger_meta)
        
        pg = ledger_meta[name]["PrimaryGroup"]
        if not ledger_meta[name]["NatureOfGroup"] and pg:
            ledger_meta[name]["NatureOfGroup"] = group_map.get(pg, {}).get("Nature", "")
        
        if ledger_meta[name]["NatureOfGroup"]:
            n_val = ledger_meta[name]["NatureOfGroup"].lower()
            if n_val in ["assets", "liabilities"]: ledger_meta[name]["Nature"] = "BS"
            elif n_val in ["income", "expenses"]: ledger_meta[name]["Nature"] = "PL"
            
        if not ledger_meta[name]["Nature"] and pg:
            bs_pl, nog = nature_from_primary_group(pg)
            ledger_meta[name]["Nature"] = bs_pl
            ledger_meta[name]["NatureOfGroup"] = nog
            
    return ledger_meta


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
            ledger_gstin = meta.get("GSTIN", "")
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


def fetch_voucher_rows(host, port, company, from_date, to_date, chunk_days):
    url = f"http://{host}:{port}"

    if not company or not from_date or not to_date:
        cmp_name, cmp_start, cmp_end = get_company_info(host, port)
        if not company:
            company = cmp_name
        if not from_date:
            from_date = cmp_start
        if not to_date:
            to_date = cmp_end

    vtype_map, group_map = fetch_tally_metadata(url, company)

    ledger_root = parse_xml_root(post_to_tally(url, build_ledger_request_xml(company)))
    ledger_meta = parse_ledgers(ledger_root, group_map)

    voucher_root = parse_xml_root(post_to_tally(url, build_voucher_request_xml(company, from_date, to_date)))
    status = clean_text(first_descendant_text(voucher_root, "STATUS"))
    if status == "0":
        error_text = first_descendant_text(voucher_root, "LINEERROR") or f"Tally returned STATUS=0 for {from_date} to {to_date}"
        raise ValueError(error_text)

    return parse_vouchers(voucher_root, ledger_meta, company, from_date, to_date, vtype_map)


rows = fetch_voucher_rows(
    host=HOST,
    port=PORT,
    company=COMPANY,
    from_date=FROM_DATE,
    to_date=TO_DATE,
    chunk_days=CHUNK_DAYS,
)

dataset = pd.DataFrame(rows)
for column in TDL_OUTPUT_COLUMNS:
    if column not in dataset.columns:
        dataset[column] = ""
dataset = dataset[[column for column in TDL_OUTPUT_COLUMNS if column in dataset.columns]]
