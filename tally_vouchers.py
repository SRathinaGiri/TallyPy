import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape
from datetime import datetime

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect
FROM_DATE = ""    # YYYYMMDD
TO_DATE = ""      # YYYYMMDD

ACCOUNTING_VOUCHER_TYPES = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra"}
VOUCHER_COLUMNS = ["Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName", "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName", "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName", "FromDate", "ToDate"]
PRIMARY_GROUPS = {"Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans", "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors", "Fixed Assets", "Investments", "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors", "Misc. Expenses (ASSET)", "Suspense Account", "Branch / Divisions", "Sales Accounts", "Purchase Accounts", "Direct Incomes", "Indirect Incomes", "Direct Expenses", "Indirect Expenses"}
CURRENCY_SYMBOL_FALLBACKS = {"INR": "₹", "INDIAN RUPEE": "₹", "RUPEE": "₹", "RUPEES": "₹", "RS": "₹", "RS.": "₹", "USD": "$", "US DOLLAR": "$", "DOLLAR": "$", "EUR": "€", "EURO": "€", "GBP": "£", "POUND": "£", "POUND STERLING": "£", "AED": "د.இ", "DIRHAM": "د.இ", "": ""}

# --- HELPERS EXACT FROM app1.py ---

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
        # FIX: Ensure 'codepoint' is used consistently
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
        v = direct_child_text(elem, name)
        if v: return v
    return ""

def first_descendant_text(elem, local_name):
    for child in elem.iter():
        if strip_ns(child.tag).upper() == local_name.upper():
            val = clean_text(child.text)
            if val: return val
    return ""

def normalize_amount_text(value):
    text = clean_text(value).replace(",", "")
    if not text: return ""
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    if not matches: return text
    token = matches[-1].group(0)
    try: return f"{Decimal(token):.2f}"
    except InvalidOperation: return token

def to_decimal(value, default=Decimal("0.00")):
    value = normalize_amount_text(value)
    if not value: return default
    try: return Decimal(value)
    except InvalidOperation: return default

def to_float(value):
    return float(to_decimal(value))

def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value

def nature_from_primary_group(primary_group):
    pg = clean_text(primary_group).lower()
    if pg in ["current assets", "fixed assets", "investments", "misc. expenses (asset)", "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)", "stock-in-hand", "sundry debtors"]:
        return "BS", "Assets"
    elif pg in ["capital account", "current liabilities", "loans (liability)", "suspense account", "branch / divisions", "bank od a/c", "duties & taxes", "provisions", "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"]:
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
        if not parent: return ""
        if parent in PRIMARY_GROUPS: return parent
        current = parent
    return ""

def post_to_tally(url, xml_text):
    r = requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=120)
    r.raise_for_status()
    return r.text

def parse_xml_root(xml_text):
    if not xml_text: return ET.Element("ROOT")
    try:
        cleaned = xml_cleanup(xml_text)
        return ET.fromstring(cleaned.encode("utf-8"))
    except: return ET.Element("ROOT")

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
    try:
        root = parse_xml_root(post_to_tally(url, xml))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name: return name, start, end
    except: pass
    return "", "", ""

def fetch_tally_metadata(url, company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company: static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    v_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>AllVTypes</ID></HEADER><BODY><DESC><STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"AllVTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    g_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>AllGroups</ID></HEADER><BODY><DESC><STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"AllGroups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    vtype_map, group_map = {}, {}
    try:
        rv = parse_xml_root(post_to_tally(url, v_xml))
        for vt in rv.iter():
            if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
                n, p = direct_child_text(vt, "NAME"), direct_child_text(vt, "PARENT")
                if n: vtype_map[n] = p or n
        rg = parse_xml_root(post_to_tally(url, g_xml))
        for g in rg.iter():
            if strip_ns(g.tag).upper() == "GROUP":
                n, p, nat, pri = direct_child_text(g, "NAME"), direct_child_text(g, "PARENT"), direct_child_text(g, "NATURE"), direct_child_text(g, "_PRIMARYGROUP")
                if n: group_map[n] = {"Parent": p, "Nature": nat, "PrimaryGroup": pri}
        base_types = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra"}
        for _ in range(5):
            for vt_n, p_n in vtype_map.items():
                if p_n and p_n not in base_types and p_n in vtype_map: vtype_map[vt_n] = vtype_map[p_n]
            for g_n, gi in group_map.items():
                p = gi.get("Parent")
                if p and not gi.get("Nature") and p in group_map: gi["Nature"] = group_map[p].get("Nature")
                if p and not gi.get("PrimaryGroup") and p in group_map: gi["PrimaryGroup"] = group_map[p].get("PrimaryGroup")
    except: pass
    return vtype_map, group_map

def parse_ledgers(root, group_map):
    ledger_rows, ledger_lookup = [], {}
    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "LEDGER": continue
        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name: continue
        parent = direct_child_text(elem, "PARENT"); gi = group_map.get(parent, {})
        row = {
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"), 
            "Name": name, 
            "PrimaryGroup": gi.get("PrimaryGroup") or first_non_empty_text(elem, ["PRIMARYGROUP"]), 
            "Nature": "", "NatureOfGroup": gi.get("Nature", ""), 
            "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]), 
            "Parent": parent, 
            "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]), 
            "OpeningBalance": to_float(first_non_empty_text(elem, ["OPENINGBALANCE"])), 
            "ClosingBalance": to_float(first_non_empty_text(elem, ["CLOSINGBALANCE"]))
        }
        ledger_rows.append(row); ledger_lookup[name] = row
    for r in ledger_rows:
        if not r["PrimaryGroup"]: r["PrimaryGroup"] = ledger_primary_group(r["Name"], ledger_lookup)
        pg = r["PrimaryGroup"]
        if not r["NatureOfGroup"] and pg: r["NatureOfGroup"] = group_map.get(pg, {}).get("Nature", "")
        if r["NatureOfGroup"]:
            nv = r["NatureOfGroup"].lower()
            if nv in ["assets", "liabilities"]: r["Nature"] = "BS"
            elif nv in ["income", "expenses"]: r["Nature"] = "PL"
        if not r["Nature"] and pg:
            bs_pl, nog = nature_from_primary_group(pg); r["Nature"], r["NatureOfGroup"] = bs_pl, nog
    return {r["Name"]: r for r in ledger_rows}

# --- EXECUTION ---

url = f"http://{HOST}:{PORT}"
det_name, det_start, det_end = get_company_info(HOST, PORT)
sel_comp = COMPANY or det_name
f_dt, t_dt = FROM_DATE or det_start, TO_DATE or det_end
vtype_map, group_map = fetch_tally_metadata(url, sel_comp)

# Fetch Ledger meta for mapping (matching app1.py)
l_req = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyLedgers</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyLedgers\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH><COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
l_meta = parse_ledgers(parse_xml_root(post_to_tally(url, l_req)), group_map)

# Final Journal extraction
v_xml = (
    f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyVouchers</ID></HEADER><BODY><DESC>"
    f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY>"
    f"<SVFROMDATE TYPE='Date'>{escape(f_dt)}</SVFROMDATE><SVTODATE TYPE='Date'>{escape(t_dt)}</SVTODATE></STATICVARIABLES>"
    "<TDL><TDLMESSAGE><SYSTEM TYPE='Formulae' NAME='IsAccountingVoucher'>"
    "($VoucherTypeName = \"Sales\") OR ($VoucherTypeName = \"Purchase\") OR ($VoucherTypeName = \"Journal\") OR ($VoucherTypeName = \"Receipt\") OR ($VoucherTypeName = \"Payment\") OR ($VoucherTypeName = \"Debit Note\") OR ($VoucherTypeName = \"Credit Note\") OR ($VoucherTypeName = \"Contra\")</SYSTEM>"
    "<OBJECT NAME=\"All Ledger Entries\"><COMPUTE>EntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE><COMPUTE>EntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE><COMPUTE>EntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE><COMPUTE>EntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE></OBJECT>"
    "<COLLECTION NAME=\"MyVouchers\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, AllLedgerEntries.IsDeemedPositive, AllLedgerEntries.EntryLedgerMasterID, AllLedgerEntries.EntryParentLedger, AllLedgerEntries.EntryPrimaryGroup, AllLedgerEntries.EntryLedgerGSTIN</FETCH><FILTER>IsAccountingVoucher</FILTER></COLLECTION>"
    "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
)

root = parse_xml_root(post_to_tally(url, v_xml))
rows = []
for voucher in root.iter():
    if strip_ns(voucher.tag).upper() != "VOUCHER": continue
    vt = direct_child_text(voucher, "VOUCHERTYPENAME"); base_vt = vtype_map.get(vt, vt)
    if base_vt not in ACCOUNTING_VOUCHER_TYPES: continue
    vd, vn, v_nar = format_tally_date(direct_child_text(voucher, "DATE")), direct_child_text(voucher, "VOUCHERNUMBER"), first_non_empty_text(voucher, ["NARRATION", "VOUCHERNARRATION"])
    
    entries = direct_children(voucher, "ALLLEDGERENTRIES.LIST") or direct_children(voucher, "LEDGERENTRIES.LIST")
    for ent in entries:
        ln = direct_child_text(ent, "LEDGERNAME"); amt_v = to_decimal(direct_child_text(ent, "AMOUNT"))
        if not ln or amt_v == 0: continue
        is_pos = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
        signed = abs(amt_v) * (Decimal("-1") if is_pos else Decimal("1"))
        
        # Mapping from l_meta
        meta = l_meta.get(ln, {})
        
        rows.append({
            "Date": vd, "VoucherTypeName": vt, "BaseVoucherType": base_vt, "VoucherNumber": vn, "LedgerName": ln, 
            "MasterID": direct_child_text(ent, "ENTRYLEDGERMASTERID") or meta.get("MasterID", ""), 
            "Amount": float(signed), "DrCr": "Dr" if signed < 0 else "Cr", 
            "DebitAmount": float(abs(signed)) if signed < 0 else 0.0, 
            "CreditAmount": float(abs(signed)) if signed > 0 else 0.0, 
            "ParentLedger": direct_child_text(ent, "ENTRYPARENTLEDGER") or meta.get("Parent", ""), 
            "PrimaryGroup": direct_child_text(ent, "ENTRYPRIMARYGROUP") or meta.get("PrimaryGroup", ""), 
            "Nature": meta.get("Nature", ""), "NatureOfGroup": meta.get("NatureOfGroup", ""), 
            "PAN": meta.get("PAN", ""), 
            "PartyLedgerName": direct_child_text(voucher, "PARTYLEDGERNAME"), 
            "PartyGSTIN": direct_child_text(voucher, "PARTYGSTIN"), 
            "LedgerGSTIN": direct_child_text(ent, "ENTRYLEDGERGSTIN") or meta.get("PartyGSTIN", ""), 
            "VoucherNarration": v_nar, "IsOptional": "Yes" if direct_child_text(voucher, "ISOPTIONAL").upper() == "YES" else "No", 
            "CompanyName": sel_comp, "FromDate": format_tally_date(det_start), "ToDate": format_tally_date(det_end)
        })

Journal = pd.DataFrame(rows, columns=VOUCHER_COLUMNS)
Journal = Journal[VOUCHER_COLUMNS]
