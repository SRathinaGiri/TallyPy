import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

HOST = "localhost"
PORT = "9000"
COMPANY = ""
FROM_DATE = "20200401"
TO_DATE = "20210331"

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

OUTPUT_COLUMNS = [
    "Date",
    "VoucherTypeName",
    "BaseVoucherType",
    "VoucherNumber",
    "LedgerName",
    "Amount",
    "IsDeemedPositive",
    "PartyLedgerName",
    "PartyGSTIN",
    "VoucherNarration",
    "IsOptional",
    "DebitAmount",
    "CreditAmount",
    "ParentLedger",
    "PrimaryGroup",
    "Nature",
    "DrCr",
    "LedgerGSTIN",
    "StatusOptional",
    "CompanyName",
    "FromDate",
    "ToDate",
    "LedMasterID",
    "VoucherKey",
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


def to_float(value):
    value = normalize_amount_text(value)
    if not value:
        return 0.0
    try:
        return float(Decimal(value))
    except InvalidOperation:
        return 0.0


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


def detect_company_name(root):
    for elem in root.iter():
        if strip_ns(elem.tag).upper() == "COMPANY":
            name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
            if name:
                return name
    return ""


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
        # Resolve Group nature recursively if missing
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


def build_request_xml(company, from_date, to_date):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    static_vars.append(f"<SVFROMDATE TYPE='Date'>{escape(from_date)}</SVFROMDATE>")
    static_vars.append(f"<SVTODATE TYPE='Date'>{escape(to_date)}</SVTODATE>")

    tdl = """
<SYSTEM TYPE="Formulae" NAME="IsAccountingVoucher">
($VoucherTypeName = "Sales") OR ($VoucherTypeName = "Purchase") OR
($VoucherTypeName = "Journal") OR ($VoucherTypeName = "Receipt") OR
($VoucherTypeName = "Payment") OR ($VoucherTypeName = "Debit Note") OR
($VoucherTypeName = "Credit Note")
</SYSTEM>
<SYSTEM TYPE="Formulae" NAME="VoucherNarration">$Narration:Voucher</SYSTEM>
<COLLECTION NAME="AllVouchersForODBC">
<TYPE>Voucher</TYPE>
<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, MasterID</FETCH>
<COMPUTE NAME="VoucherNarration">$Narration</COMPUTE>
</COLLECTION>
<COLLECTION NAME="_All_Vouchers">
<SOURCECOLLECTION>AllVouchersForODBC</SOURCECOLLECTION>
<WALK>All Ledger Entries</WALK>
<FETCH>Date, VoucherTypeName, VoucherNumber, LedgerName, Amount, IsDeemedPositive, PartyLedgerName, PartyGSTIN, VoucherNarration, IsOptional</FETCH>
<COMPUTE NAME="AmountSignedForDRCR">If (If $IsDeemedPositive Then $Amount * -1 Else $Amount) &lt; 0 Then $$Abs:$Amount * -1 Else $$Abs:$Amount</COMPUTE>
<COMPUTE NAME="DebitAmount">If $$AsAmount:##AmountSignedForDRCR &lt; 0 Then $$Abs:$Amount Else 0</COMPUTE>
<COMPUTE NAME="CreditAmount">If $$AsAmount:##AmountSignedForDRCR &gt; 0 Then $$Abs:$Amount Else 0</COMPUTE>
<COMPUTE NAME="DrCr">If $$AsAmount:##AmountSignedForDRCR &lt; 0 Then "Dr" Else "Cr"</COMPUTE>
<COMPUTE NAME="ParentLedger">$Parent:Ledger:$LedgerName</COMPUTE>
<COMPUTE NAME="PrimaryGroup">$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>
<COMPUTE NAME="PAN">$IncomeTaxNumber:Ledger:$LedgerName</COMPUTE>
<COMPUTE NAME="LedMasterID">$MasterID:Ledger:$LedgerName</COMPUTE>
<COMPUTE NAME="LedgerGSTIN">$PartyGSTIN:Ledger:$LedgerName</COMPUTE>
<COMPUTE NAME="StatusOptional">If $IsOptional Then "Yes" Else "No"</COMPUTE>
<COMPUTE NAME="CompanyName">##SVCurrentCompany</COMPUTE>
<COMPUTE NAME="FromDate">If $$IsEmpty:##SVFromDate Then $_ThisYearBeg:Company:##SVCurrentCompany Else ##SVFromDate</COMPUTE>
<COMPUTE NAME="ToDate">If $$IsEmpty:##SVToDate Then $_ThisYearEnd:Company:##SVCurrentCompany Else ##SVToDate</COMPUTE>
<FILTER>IsAccountingVoucher</FILTER>
</COLLECTION>
"""

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>_All_Vouchers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        f"<TDL><TDLMESSAGE>{tdl}</TDLMESSAGE></TDL>"
        "</DESC></BODY></ENVELOPE>"
    )


def parse_rows(root, vtype_map=None, group_map=None):
    rows = []
    if vtype_map is None: vtype_map = {}
    if group_map is None: group_map = {}
    
    for obj in root.iter():
        tag = strip_ns(obj.tag).upper()
        if tag not in {"LEDGERENTRY"}:
            continue
        ledger_name = direct_child_text(obj, "LEDGERNAME")
        if not ledger_name:
            continue
        amount_raw = direct_child_text(obj, "AMOUNT")
        amount_val = to_float(amount_raw)
        if amount_val == 0:
            continue
        dr_cr = direct_child_text(obj, "DRCR")
        normalized_amount = abs(amount_val)
        if dr_cr == "Dr":
            normalized_amount *= -1
        primary_group = direct_child_text(obj, "PRIMARYGROUP")
        status_optional = direct_child_text(obj, "STATUSOPTIONAL") or direct_child_text(obj, "ISOPTIONAL")
        voucher_number = direct_child_text(obj, "VOUCHERNUMBER")
        voucher_type = direct_child_text(obj, "VOUCHERTYPENAME")
        
        base_v_type = vtype_map.get(voucher_type, voucher_type)
        nature_of_group = group_map.get(primary_group, {}).get("Nature", "")
        nature = ""
        
        if nature_of_group:
            nv = nature_of_group.lower()
            if nv in ["assets", "liabilities"]: nature = "BS"
            elif nv in ["income", "expenses"]: nature = "PL"
            
        if not nature:
            bs_pl, nog = nature_from_primary_group(primary_group)
            nature = bs_pl
            nature_of_group = nog

        row = {
            "Date": format_tally_date(direct_child_text(obj, "DATE")),
            "VoucherTypeName": voucher_type,
            "BaseVoucherType": base_v_type,
            "VoucherNumber": voucher_number,
            "LedgerName": ledger_name,
            "Amount": normalized_amount,
            "IsDeemedPositive": 1 if dr_cr == "Dr" else 0,
            "PartyLedgerName": direct_child_text(obj, "PARTYLEDGERNAME"),
            "PartyGSTIN": direct_child_text(obj, "PARTYGSTIN"),
            "VoucherNarration": direct_child_text(obj, "VOUCHERNARRATION") or direct_child_text(obj, "NARRATION"),
            "IsOptional": 1 if status_optional == "Yes" else 0,
            "DrCr": dr_cr,
            "DebitAmount": to_float(direct_child_text(obj, "DEBITAMOUNT")),
            "CreditAmount": to_float(direct_child_text(obj, "CREDITAMOUNT")),
            "ParentLedger": direct_child_text(obj, "PARENTLEDGER"),
            "PrimaryGroup": primary_group,
            "Nature": nature,
            "NatureOfGroup": nature_of_group,
            "PAN": direct_child_text(obj, "PAN"),
            "LedgerGSTIN": direct_child_text(obj, "LEDGERGSTIN"),
            "StatusOptional": status_optional,
            "CompanyName": direct_child_text(obj, "COMPANYNAME"),
            "FromDate": format_tally_date(direct_child_text(obj, "FROMDATE")),
            "ToDate": format_tally_date(direct_child_text(obj, "TODATE")),
            "LedMasterID": direct_child_text(obj, "LEDMASTERID"),
            "VoucherKey": f"{voucher_type}:{voucher_number}" if voucher_type and voucher_number else "",
        }
        if row["VoucherTypeName"]:
            rows.append(row)
    return rows


def fetch_rows():
    url = f"http://{host}:{port}" # Wait, should be HOST and PORT or arguments
    url = f"http://{HOST}:{PORT}"
    company = COMPANY
    if not company:
        company_root = parse_xml_root(post_to_tally(url, build_company_request_xml()))
        company = detect_company_name(company_root)

    vtype_map, group_map = fetch_tally_metadata(url, company)
    
    root = parse_xml_root(post_to_tally(url, build_request_xml(company, FROM_DATE, TO_DATE)))
    status = clean_text(first_descendant_text(root, "STATUS"))
    if status == "0":
        error_text = first_descendant_text(root, "LINEERROR") or "Tally returned STATUS=0"
        raise ValueError(error_text)
    return parse_rows(root, vtype_map, group_map)


rows = fetch_rows()
dataset = pd.DataFrame(rows)
for column in OUTPUT_COLUMNS:
    if column not in dataset.columns:
        dataset[column] = ""
dataset = dataset[OUTPUT_COLUMNS]
