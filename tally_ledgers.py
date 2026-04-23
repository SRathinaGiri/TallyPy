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

# Mapping constants (Exact match with app1.py)
BS_PRIMARY_GROUPS = {"Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans", "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors", "Fixed Assets", "Investments", "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors", "Misc. Expenses (ASSET)", "Suspense Account", "Branch / Divisions"}
PL_PRIMARY_GROUPS = {"Sales Accounts", "Purchase Accounts", "Direct Incomes", "Indirect Incomes", "Direct Expenses", "Indirect Expenses"}
PRIMARY_GROUPS = BS_PRIMARY_GROUPS | PL_PRIMARY_GROUPS
CURRENCY_SYMBOL_FALLBACKS = {"INR": "₹", "INDIAN RUPEE": "₹", "RUPEE": "₹", "RUPEES": "₹", "RS": "₹", "RS.": "₹", "USD": "$", "US DOLLAR": "$", "DOLLAR": "$", "EUR": "€", "EURO": "€", "GBP": "£", "POUND": "£", "POUND STERLING": "£", "AED": "د.إ", "DIRHAM": "د.إ", "": ""}

LEDGER_COLUMNS = ["MasterID", "Name", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "StartingFrom", "CurrencyName", "StateName", "Parent", "PartyGSTIN", "OpeningBalance", "ClosingBalance", "CompanyName", "FromDate", "ToDate"]

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
    g_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>GR</ID></HEADER><BODY><DESC>{sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"GR\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    gm = {}
    try:
        rg = requests.post(url, data=g_xml.encode("utf-8"), timeout=30)
        for g in ET.fromstring(xml_cleanup(rg.text)).iter():
            if strip_ns(g.tag).upper() == "GROUP":
                n = direct_child_text(g, "NAME")
                if n: gm[n] = {"Parent": direct_child_text(g, "PARENT"), "Nature": direct_child_text(g, "NATURE"), "PrimaryGroup": direct_child_text(g, "_PRIMARYGROUP")}
        for _ in range(5):
            for n, i in gm.items():
                p = i["Parent"]
                if p and not i["Nature"] and p in gm: i["Nature"] = gm[p]["Nature"]
                if p and not i["PrimaryGroup"] and p in gm: i["PrimaryGroup"] = gm[p]["PrimaryGroup"]
    except: pass
    return gm

# Execution
url = f"http://{HOST}:{PORT}"
c_name, s_dt, e_dt = get_company_info(HOST, PORT)
sel_comp = COMPANY or c_name
group_map = fetch_metadata(url, sel_comp)

l_req = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>L</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"L\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH><COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE><COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE><COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE><COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
l_rows = []
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
        cur_key = clean_text(direct_child_text(elem, "CURRENCYFORMALNAME") or direct_child_text(elem, "CURRENCYNAME")).upper()
        c_name_sym = CURRENCY_SYMBOL_FALLBACKS.get(cur_key, clean_text(direct_child_text(elem, "CURRENCYSYMBOL") or direct_child_text(elem, "CURRENCYORIGINALSYMBOL")))
        l_rows.append({"MasterID": elem.get("MASTERID") or direct_child_text(elem, "MASTERID"), "Name": name, "PrimaryGroup": pg, "Nature": nat, "NatureOfGroup": nog, "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]), "StartingFrom": direct_child_text(elem, "STARTINGFROM"), "CurrencyName": c_name_sym, "StateName": direct_child_text(elem, "STATENAME"), "Parent": parent, "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]), "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))), "ClosingBalance": float(to_decimal(direct_child_text(elem, "CLOSINGBALANCE"))), "CompanyName": sel_comp, "FromDate": format_tally_date(s_dt), "ToDate": format_tally_date(e_dt)})
except: pass

dataset = pd.DataFrame(l_rows, columns=LEDGER_COLUMNS)
for c in LEDGER_COLUMNS:
    if c not in dataset.columns: dataset[c] = ""
dataset = dataset[LEDGER_COLUMNS]
