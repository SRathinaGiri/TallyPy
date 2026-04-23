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
BS_PRIMARY_GROUPS = {"Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans", "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors", "Fixed Assets", "Investments", "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors", "Misc. Expenses (ASSET)", "Suspense Account", "Branch / Divisions"}
PL_PRIMARY_GROUPS = {"Sales Accounts", "Purchase Accounts", "Direct Incomes", "Indirect Incomes", "Direct Expenses", "Indirect Expenses"}
PRIMARY_GROUPS = BS_PRIMARY_GROUPS | PL_PRIMARY_GROUPS

VOUCHER_COLUMNS = ["Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName", "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName", "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName", "FromDate", "ToDate"]

# Core Helpers (Line-for-line with app1.py)
def strip_ns(tag):
    if not isinstance(tag, str): return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text)).strip()

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        v = match.group(1)
        try: cp = int(v[1:], 16) if v.lower().startswith("x") else int(v)
        except: return ""
        return match.group(0) if cp in (9, 10, 13) or (32 <= cp <= 55295) or (57344 <= cp <= 65533) or (65536 <= cp <= 1114111) else ""
    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", fix_char_ref, xml_text)
    xml_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_text)
    xml_text = re.sub(r"&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;)", "&amp;", xml_text)
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)
    xml_text = re.sub(r'\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*"[^"]*"', "", xml_text)
    return xml_text

def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper(): return clean_text(child.text)
    return ""

def first_non_empty_text(elem, names):
    for n in names:
        v = direct_child_text(elem, n)
        if v: return v
    return ""

def to_decimal(value, default=Decimal("0.00")):
    text = clean_text(value).replace(",", "")
    if not text: return default
    matches = list(re.finditer(r"[-+]?\d+(?:\.\d+)?", text))
    if not matches: return default
    try: return Decimal(matches[-1].group(0))
    except: return default

def format_tally_date(value):
    value = clean_text(value)
    if re.fullmatch(r"\d{8}", value): return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value

def nature_from_primary_group(primary_group):
    pg = clean_text(primary_group).lower()
    if pg in ["current assets", "fixed assets", "investments", "misc. expenses (asset)", "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)", "stock-in-hand", "sundry debtors"]: return "BS", "Assets"
    elif pg in ["capital account", "current liabilities", "loans (liability)", "suspense account", "branch / divisions", "bank od a/c", "duties & taxes", "provisions", "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"]: return "BS", "Liabilities"
    elif pg in ["direct incomes", "indirect incomes", "sales accounts"]: return "PL", "Income"
    elif pg in ["direct expenses", "indirect expenses", "purchase accounts"]: return "PL", "Expenses"
    return "Unknown", "Unknown"

def post_to_tally(url, xml_text, timeout=120):
    r = requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=timeout)
    r.raise_for_status()
    return r.content.decode(r.encoding or "utf-8", errors="replace")

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        root = ET.fromstring(xml_cleanup(post_to_tally(url, xml)))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY": return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

def fetch_tally_metadata(url, company):
    sv = f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>" if company else ""
    v_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>VT</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{sv}</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"VT\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    g_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>GR</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{sv}</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"GR\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    vm, gm = {}, {}
    try:
        rv = ET.fromstring(xml_cleanup(post_to_tally(url, v_xml)))
        for vt in rv.iter():
            if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
                n, p = direct_child_text(vt, "NAME"), direct_child_text(vt, "PARENT")
                if n: vm[n] = p or n
        rg = ET.fromstring(xml_cleanup(post_to_tally(url, g_xml)))
        for g in rg.iter():
            if strip_ns(g.tag).upper() == "GROUP":
                n, p, nat, pri = direct_child_text(g, "NAME"), direct_child_text(g, "PARENT"), direct_child_text(g, "NATURE"), direct_child_text(g, "_PRIMARYGROUP")
                if n: gm[n] = {"Parent": p, "Nature": nat, "PrimaryGroup": pri}
        base_types = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra"}
        for _ in range(5):
            for vt_n, p_n in vm.items():
                if p_n and p_n not in base_types and p_n in vm: vm[vt_n] = vm[p_n]
            for g_n, gi in gm.items():
                p = gi.get("Parent")
                if p and not gi.get("Nature") and p in gm: gi["Nature"] = gm[p].get("Nature")
                if p and not gi.get("PrimaryGroup") and p in gm: gi["PrimaryGroup"] = gm[p].get("PrimaryGroup")
    except: pass
    return vm, gm

def build_voucher_request_xml(company, from_date, to_date):
    sv = [f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company: sv.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    sv.append(f"<SVFROMDATE TYPE='Date'>{escape(from_date)}</SVFROMDATE>")
    sv.append(f"<SVTODATE TYPE='Date'>{escape(to_date)}</SVTODATE>")
    return (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>V</ID></HEADER><BODY><DESC><STATICVARIABLES>{''.join(sv)}</STATICVARIABLES><TDL><TDLMESSAGE>"
        "<SYSTEM TYPE='Formulae' NAME='IsAccountingVoucher'>($VoucherTypeName = \"Sales\") OR ($VoucherTypeName = \"Purchase\") OR ($VoucherTypeName = \"Journal\") OR ($VoucherTypeName = \"Receipt\") OR ($VoucherTypeName = \"Payment\") OR ($VoucherTypeName = \"Debit Note\") OR ($VoucherTypeName = \"Credit Note\")</SYSTEM>"
        "<OBJECT NAME=\"All Ledger Entries\"><COMPUTE>EntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE><COMPUTE>EntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE><COMPUTE>EntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE><COMPUTE>EntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE></OBJECT>"
        "<COLLECTION NAME=\"V\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, AllLedgerEntries.IsDeemedPositive, AllLedgerEntries.EntryLedgerMasterID, AllLedgerEntries.EntryParentLedger, AllLedgerEntries.EntryPrimaryGroup, AllLedgerEntries.EntryLedgerGSTIN</FETCH><FILTER>IsAccountingVoucher</FILTER></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )

# Execution Flow
url = f"http://{HOST}:{PORT}"
det_name, det_start, det_end = get_company_info(HOST, PORT)
sel_comp = COMPANY or det_name
f_dt, t_dt = FROM_DATE or det_start, TO_DATE or det_end
v_map, g_map = fetch_tally_metadata(url, sel_comp)

# Fetch Ledgers for meta
l_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>LM</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"LM\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, MasterID, PartyGSTIN</FETCH><COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
l_meta = {}
try:
    rl = ET.fromstring(xml_cleanup(post_to_tally(url, l_xml)))
    for elem in rl.iter():
        if strip_ns(elem.tag).upper() == "LEDGER":
            name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
            if not name: continue
            parent = direct_child_text(elem, "PARENT"); gi = g_map.get(parent, {})
            pg = gi.get("PrimaryGroup") or direct_child_text(elem, "PRIMARYGROUP")
            nog = gi.get("Nature", "")
            nat = "BS" if nog and nog.lower() in ["assets", "liabilities"] else ("PL" if nog and nog.lower() in ["income", "expenses"] else "")
            if not nat and pg: nat, nog = nature_from_primary_group(pg)
            l_meta[name] = {"Name": name, "Parent": parent, "PrimaryGroup": pg, "MasterID": elem.get("MASTERID") or direct_child_text(elem, "MASTERID"), "Nature": nat, "NatureOfGroup": nog, "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]), "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"])}
except: pass

# Fetch Vouchers
root = ET.fromstring(xml_cleanup(post_to_tally(url, build_voucher_request_xml(sel_comp, f_dt, t_dt))))
rows = []
for v in root.iter():
    if strip_ns(v.tag).upper() != "VOUCHER": continue
    vtype = direct_child_text(v, "VOUCHERTYPENAME"); base_vt = v_map.get(vtype, vtype)
    if base_vt not in ACCOUNTING_VOUCHER_TYPES: continue
    vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
    entries = [c for c in list(v) if "LEDGERENTRIES.LIST" in strip_ns(c.tag).upper() or "ALLLEDGERENTRIES.LIST" in strip_ns(c.tag).upper()]
    for ent in entries:
        ln = direct_child_text(ent, "LEDGERNAME"); amt_v = to_decimal(direct_child_text(ent, "AMOUNT"))
        if not ln or amt_v == 0: continue
        is_pos = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
        signed = abs(amt_v) * (Decimal("-1") if is_pos else Decimal("1"))
        meta = l_meta.get(ln, {})
        rows.append({"Date": vd, "VoucherTypeName": vtype, "BaseVoucherType": base_vt, "VoucherNumber": vn, "LedgerName": ln, "MasterID": meta.get("MasterID", ""), "Amount": float(signed), "DrCr": "Dr" if signed < 0 else "Cr", "DebitAmount": float(abs(signed)) if signed < 0 else 0.0, "CreditAmount": float(abs(signed)) if signed > 0 else 0.0, "ParentLedger": meta.get("Parent", ""), "PrimaryGroup": meta.get("PrimaryGroup", ""), "Nature": meta.get("Nature", ""), "NatureOfGroup": meta.get("NatureOfGroup", ""), "PAN": meta.get("PAN", ""), "PartyLedgerName": direct_child_text(v, "PARTYLEDGERNAME"), "PartyGSTIN": direct_child_text(v, "PARTYGSTIN"), "LedgerGSTIN": meta.get("PartyGSTIN", ""), "VoucherNarration": v_nar, "IsOptional": direct_child_text(v, "ISOPTIONAL"), "CompanyName": sel_comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})

Journal = pd.DataFrame(rows, columns=VOUCHER_COLUMNS)
Journal['CompanyName'] = sel_comp
Journal['FromDate'] = format_tally_date(f_dt)
Journal['ToDate'] = format_tally_date(t_dt)
Journal = Journal[VOUCHER_COLUMNS]
