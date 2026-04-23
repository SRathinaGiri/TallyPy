import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
import time
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect

LEDGER_COLUMNS = ["MasterID", "Name", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "StartingFrom", "CurrencyName", "StateName", "Parent", "PartyGSTIN", "OpeningBalance", "ClosingBalance", "CompanyName", "FromDate", "ToDate"]
CURRENCY_SYMBOL_FALLBACKS = {"INR": "₹", "INDIAN RUPEE": "₹", "RUPEE": "₹", "RUPEES": "₹", "RS": "₹", "RS.": "₹", "USD": "$", "US DOLLAR": "$", "DOLLAR": "$", "EUR": "€", "EURO": "€", "GBP": "£", "POUND": "£", "POUND STERLING": "£", "AED": "د.இ", "DIRHAM": "د.இ", "": ""}
PRIMARY_GROUPS = {"Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans", "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors", "Fixed Assets", "Investments", "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors", "Misc. Expenses (ASSET)", "Suspense Account", "Branch / Divisions", "Sales Accounts", "Purchase Accounts", "Direct Incomes", "Indirect Incomes", "Direct Expenses", "Indirect Expenses"}

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
    return xml_text

def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper(): return clean_text(child.text)
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

def post_to_tally(url, xml_text):
    return requests.post(url, data=xml_text.encode("utf-8"), headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=120).text

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        root = ET.fromstring(xml_cleanup(post_to_tally(url, xml)))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY": return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

def fetch_gm(url, company):
    sv = f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>" if company else ""
    g_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>GR</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{sv}</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"GR\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    gm = {}
    try:
        rg = ET.fromstring(xml_cleanup(post_to_tally(url, g_xml)))
        for g in rg.iter():
            if strip_ns(g.tag).upper() == "GROUP":
                n, p, nat, pri = direct_child_text(g, "NAME"), direct_child_text(g, "PARENT"), direct_child_text(g, "NATURE"), direct_child_text(g, "_PRIMARYGROUP")
                if n: gm[n] = {"Parent": p, "Nature": nat, "PrimaryGroup": pri}
        for _ in range(5):
            for gi in gm.values():
                p = gi.get("Parent")
                if p and not gi.get("Nature") and p in gm: gi["Nature"] = gm[p].get("Nature")
                if p and not gi.get("PrimaryGroup") and p in gm: gi["PrimaryGroup"] = gm[p].get("PrimaryGroup")
    except: pass
    return gm

# EXECUTION
url = f"http://{HOST}:{PORT}"
c_name, s_dt, e_dt = get_company_info(HOST, PORT)
sel_comp = COMPANY or c_name
group_map = fetch_gm(url, sel_comp)

# Lighter XML request (No COMPUTE fields inside Tally)
l_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>L</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(sel_comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"L\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, OpeningBalance, ClosingBalance, IncomeTaxNumber, CurrencyName</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
root = ET.fromstring(xml_cleanup(post_to_tally(url, l_xml)))
rows = []
for elem in root.iter():
    if strip_ns(elem.tag).upper() != "LEDGER": continue
    name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
    if not name: continue
    p = direct_child_text(elem, "PARENT"); gi = group_map.get(p, {})
    pg = gi.get("PrimaryGroup") or direct_child_text(elem, "PRIMARYGROUP")
    nog = gi.get("Nature", ""); nat = "BS" if nog and nog.lower() in ["assets", "liabilities"] else ("PL" if nog and nog.lower() in ["income", "expenses"] else "")
    c_sym = CURRENCY_SYMBOL_FALLBACKS.get(clean_text(direct_child_text(elem, "CURRENCYNAME")).upper(), "")
    rows.append({"MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"), "Name": name, "PrimaryGroup": pg, "Nature": nat, "NatureOfGroup": nog, "PAN": clean_text(direct_child_text(elem, "PAN")), "StartingFrom": direct_child_text(elem, "STARTINGFROM"), "CurrencyName": c_sym, "StateName": "", "Parent": p, "PartyGSTIN": direct_child_text(elem, "PARTYGSTIN"), "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))), "ClosingBalance": float(to_decimal(direct_child_text(elem, "CLOSINGBALANCE"))), "CompanyName": sel_comp, "FromDate": format_tally_date(s_dt), "ToDate": format_tally_date(e_dt)})

Ledger = pd.DataFrame(rows, columns=LEDGER_COLUMNS)
Ledger = Ledger[LEDGER_COLUMNS]
