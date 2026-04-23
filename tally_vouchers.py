import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal
from xml.sax.saxutils import escape
from datetime import datetime, timedelta

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect
FROM_DATE = ""    # Leave blank for FY start (YYYYMMDD)
TO_DATE = ""      # Leave blank for FY end (YYYYMMDD)

# Mapping constants
ACCOUNTING_VOUCHER_TYPES = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra"}

VOUCHER_COLUMNS = ["Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName", "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName", "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName", "FromDate", "ToDate"]

# Helper functions
def strip_ns(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text)).strip()

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        val = match.group(1)
        try: cp = int(val[1:], 16) if val.lower().startswith("x") else int(val)
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
    for name in names:
        v = direct_child_text(elem, name); 
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

def nature_from_pg(pg):
    pg = clean_text(pg).lower()
    if pg in ["current assets", "fixed assets", "investments", "misc. expenses (asset)", "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)", "stock-in-hand", "sundry debtors"]: return "BS", "Assets"
    elif pg in ["capital account", "current liabilities", "loans (liability)", "suspense account", "branch / divisions", "bank od a/c", "duties & taxes", "provisions", "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"]: return "BS", "Liabilities"
    elif pg in ["direct incomes", "indirect incomes", "sales accounts"]: return "PL", "Income"
    elif pg in ["direct expenses", "indirect expenses", "purchase accounts"]: return "PL", "Expenses"
    return "Unknown", "Unknown"

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=10)
        root = ET.fromstring(xml_cleanup(r.text).encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY": return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

def fetch_metadata(url, company):
    sv = f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY></STATICVARIABLES>"
    v_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>VT</ID></HEADER><BODY><DESC>{sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"VT\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    g_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>GR</ID></HEADER><BODY><DESC>{sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"GR\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    vm, gm = {}, {}
    try:
        rv = requests.post(url, data=v_xml.encode("utf-8"), timeout=30)
        for vt in ET.fromstring(xml_cleanup(rv.text)).iter():
            if strip_ns(vt.tag).upper() == "VOUCHERTYPE":
                n, p = direct_child_text(vt, "NAME"), direct_child_text(vt, "PARENT")
                if n: vm[n] = p or n
        rg = requests.post(url, data=g_xml.encode("utf-8"), timeout=30)
        for g in ET.fromstring(xml_cleanup(rg.text)).iter():
            if strip_ns(g.tag).upper() == "GROUP":
                n = direct_child_text(g, "NAME")
                if n: gm[n] = {"Parent": direct_child_text(g, "PARENT"), "Nature": direct_child_text(g, "NATURE"), "PrimaryGroup": direct_child_text(g, "_PRIMARYGROUP")}
        for _ in range(5):
            for n, p in vm.items():
                if p and p not in ACCOUNTING_VOUCHER_TYPES and p in vm: vm[n] = vm[p]
            for n, i in gm.items():
                p = i["Parent"]
                if p and not i["Nature"] and p in gm: i["Nature"] = gm[p]["Nature"]
                if p and not i["PrimaryGroup"] and p in gm: i["PrimaryGroup"] = gm[p]["PrimaryGroup"]
    except: pass
    return vm, gm

# Execution
url = f"http://{HOST}:{PORT}"
c_name, s_dt, e_dt = get_company_info(HOST, PORT)
sel_comp = COMPANY or c_name
now = datetime.now()
def_start, def_end = (f"{now.year-1}0401", f"{now.year}0331") if now.month < 4 else (f"{now.year}0401", f"{now.year+1}0331")
f_dt = str(FROM_DATE or s_dt or def_start).strip()
t_dt = str(TO_DATE or e_dt or def_end).strip()
if not re.fullmatch(r"\d{8}", f_dt): f_dt = def_start
if not re.fullmatch(r"\d{8}", t_dt): t_dt = def_end

vtype_map, group_map = fetch_metadata(url, sel_comp)

# Ledger Meta for Vouchers
l_req = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>LM</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"LM\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, PartyGSTIN, MasterID, IncomeTaxNumber</FETCH><COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
l_meta = {}
try:
    rl = requests.post(url, data=l_req.encode("utf-8"), timeout=60)
    for elem in ET.fromstring(xml_cleanup(rl.text)).iter():
        if strip_ns(elem.tag).upper() != "LEDGER": continue
        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name: continue
        parent = direct_child_text(elem, "PARENT"); g_info = group_map.get(parent, {})
        nog = g_info.get("Nature", ""); pg = g_info.get("PrimaryGroup", "") or first_non_empty_text(elem, ["PRIMARYGROUP"])
        nat = "BS" if nog and nog.lower() in ["assets", "liabilities"] else ("PL" if nog and nog.lower() in ["income", "expenses"] else "")
        if not nat and pg: nat, nog = nature_from_pg(pg)
        l_meta[name] = {"MasterID": elem.get("MASTERID") or direct_child_text(elem, "MASTERID"), "Name": name, "PrimaryGroup": pg, "Nature": nat, "NatureOfGroup": nog, "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]), "Parent": parent, "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"])}
except: pass

# Chunking Vouchers
v_rows = []
d1, d2 = datetime.strptime(f_dt, "%Y%m%d"), datetime.strptime(t_dt, "%Y%m%d")
curr = d1
while curr <= d2:
    cs = curr.strftime("%Y%m%d"); ce = min(d2, (curr + timedelta(days=31)).replace(day=1) - timedelta(days=1)); ce_str = ce.strftime("%Y%m%d")
    chunk_sv = f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY><SVFROMDATE TYPE='Date'>{cs}</SVFROMDATE><SVTODATE TYPE='Date'>{ce_str}</SVTODATE></STATICVARIABLES>"
    try:
        rv = requests.post(url, data=f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>V</ID></HEADER><BODY><DESC>{chunk_sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"V\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional, AllLedgerEntries.*</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>".encode("utf-8"), timeout=120)
        for v in ET.fromstring(xml_cleanup(rv.text)).iter():
            if strip_ns(v.tag).upper() != "VOUCHER": continue
            vtype = direct_child_text(v, "VOUCHERTYPENAME")
            if vtype_map.get(vtype, vtype) not in ACCOUNTING_VOUCHER_TYPES: continue
            vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
            for ent in [c for c in list(v) if "LEDGERENTRIES.LIST" in strip_ns(c.tag).upper()]:
                ln = direct_child_text(ent, "LEDGERNAME"); amt = to_decimal(direct_child_text(ent, "AMOUNT"))
                if not ln or amt == 0: continue
                signed = abs(amt) * (Decimal("-1") if direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES" else Decimal("1"))
                m = l_meta.get(ln, {})
                v_rows.append({"Date": vd, "VoucherTypeName": vtype, "BaseVoucherType": vtype_map.get(vtype, vtype), "VoucherNumber": vn, "LedgerName": ln, "MasterID": m.get("MasterID", ""), "Amount": float(signed), "DrCr": "Dr" if signed < 0 else "Cr", "DebitAmount": float(abs(signed)) if signed < 0 else 0.0, "CreditAmount": float(abs(signed)) if signed > 0 else 0.0, "ParentLedger": m.get("Parent", ""), "PrimaryGroup": m.get("PrimaryGroup", ""), "Nature": m.get("Nature", ""), "NatureOfGroup": m.get("NatureOfGroup", ""), "PAN": m.get("PAN", ""), "PartyLedgerName": direct_child_text(v, "PARTYLEDGERNAME"), "PartyGSTIN": direct_child_text(v, "PARTYGSTIN"), "LedgerGSTIN": m.get("PartyGSTIN", ""), "VoucherNarration": v_nar, "IsOptional": direct_child_text(v, "ISOPTIONAL"), "CompanyName": sel_comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
    except: pass
    curr = ce + timedelta(days=1)

dataset = pd.DataFrame(v_rows, columns=VOUCHER_COLUMNS)
for c in VOUCHER_COLUMNS:
    if c not in dataset.columns: dataset[c] = ""
dataset = dataset[VOUCHER_COLUMNS]
