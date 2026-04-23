import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
import os
import time
import pickle
import tempfile
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally configuration
HOST = "localhost"
PORT = "9000"
COMPANY = ""      # Leave blank to auto-detect
FROM_DATE = ""    # Leave blank for FY start (YYYYMMDD)
TO_DATE = ""      # Leave blank for FY end (YYYYMMDD)

# Mapping constants
ACCOUNTING_VOUCHER_TYPES = {
    "Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra",
}

BS_PRIMARY_GROUPS = {
    "Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c", "Secured Loans",
    "Unsecured Loans", "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors",
    "Fixed Assets", "Investments", "Current Assets", "Stock-in-hand", "Deposits (Asset)",
    "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors",
    "Misc. Expenses (ASSET)", "Suspense Account", "Branch / Divisions",
}

PL_PRIMARY_GROUPS = {
    "Sales Accounts", "Purchase Accounts", "Direct Incomes", "Indirect Incomes",
    "Direct Expenses", "Indirect Expenses",
}

PRIMARY_GROUPS = BS_PRIMARY_GROUPS | PL_PRIMARY_GROUPS

# Helper functions
def strip_ns(tag):
    if not isinstance(tag, str): return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text)).strip()

def xml_cleanup(xml_text):
    def fix_char_ref(match):
        value = match.group(1)
        try:
            cp = int(value[1:], 16) if value.lower().startswith("x") else int(value)
        except: return ""
        if cp in (9, 10, 13) or (32 <= cp <= 55295) or (57344 <= cp <= 65533) or (65536 <= cp <= 1114111):
            return match.group(0)
        return ""
    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", fix_char_ref, xml_text)
    xml_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", xml_text)
    xml_text = re.sub(r"&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;)", "&amp;", xml_text)
    xml_text = re.sub(r"<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*)", r"<\1\2", xml_text)
    xml_text = re.sub(r'\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*"[^"]*"', "", xml_text)
    return xml_text

def direct_child_text(elem, local_name):
    for child in list(elem):
        if strip_ns(child.tag).upper() == local_name.upper():
            return clean_text(child.text)
    return ""

def first_non_empty_text(elem, names):
    for name in names:
        v = direct_child_text(elem, name)
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

def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE>"
        "<FETCH>Name, StartingFrom, EndingAt</FETCH>"
        "<FILTER>IsActiveCompany</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=10)
        root = ET.fromstring(xml_cleanup(r.text).encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                return direct_child_text(cmp, "NAME"), direct_child_text(cmp, "STARTINGFROM"), direct_child_text(cmp, "ENDINGAT")
    except: pass
    return "", "", ""

def fetch_metadata(url, company):
    sv = f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY></STATICVARIABLES>"
    v_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>VTypes</ID></HEADER><BODY><DESC>{sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"VTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    g_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>Groups</ID></HEADER><BODY><DESC>{sv}<TDL><TDLMESSAGE><COLLECTION NAME=\"Groups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    
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
        
        base_types = {"Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra"}
        for _ in range(5):
            for n, p in vm.items():
                if p and p not in base_types and p in vm: vm[n] = vm[p]
            for n, i in gm.items():
                p = i["Parent"]
                if p and not i["Nature"] and p in gm: i["Nature"] = gm[p]["Nature"]
                if p and not i["PrimaryGroup"] and p in gm: i["PrimaryGroup"] = gm[p]["PrimaryGroup"]
    except: pass
    return vm, gm

# Caching logic to prevent Power BI from hitting Tally multiple times simultaneously
cache_file = os.path.join(tempfile.gettempdir(), f"tally_data_cache_{PORT}.pkl")
cache_expiry = 60 # 60 seconds

def load_cached_data():
    if os.path.exists(cache_file):
        if (time.time() - os.path.getmtime(cache_file)) < cache_expiry:
            try:
                with open(cache_file, 'rb') as f:
                    return pickle.load(f)
            except: return None
    return None

def save_to_cache(data):
    try:
        with open(cache_file, 'wb') as f:
            pickle.dump(data, f)
    except: pass

cached_dfs = load_cached_data()

if cached_dfs:
    Journal = cached_dfs['Journal']
    Ledger = cached_dfs['Ledger']
    StockItem = cached_dfs['StockItem']
    StockVoucher = cached_dfs['StockVoucher']
else:
    # Main extraction logic
    url = f"http://{HOST}:{PORT}"
    company_name, start_dt, end_dt = get_company_info(HOST, PORT)
    COMPANY = COMPANY or company_name
    FROM_DATE = FROM_DATE or start_dt
    TO_DATE = TO_DATE or end_dt

    vtype_map, group_map = fetch_metadata(url, COMPANY)

    # 1. LEDGERS
    time.sleep(1) # Breathe
    l_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>L</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"L\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    l_rows = []
    try:
        rl = requests.post(url, data=l_xml.encode("utf-8"), timeout=60)
        for elem in ET.fromstring(xml_cleanup(rl.text)).iter():
            if strip_ns(elem.tag).upper() != "LEDGER": continue
            name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
            if not name: continue
            parent = direct_child_text(elem, "PARENT")
            pg = group_map.get(parent, {}).get("PrimaryGroup", "") or first_non_empty_text(elem, ["PRIMARYGROUP"])
            nog = group_map.get(parent, {}).get("Nature", "")
            nat = ""
            if nog:
                nv = nog.lower()
                if nv in ["assets", "liabilities"]: nat = "BS"
                elif nv in ["income", "expenses"]: nat = "PL"
            l_rows.append({
                "MasterID": elem.get("MASTERID") or direct_child_text(elem, "MASTERID"),
                "Name": name, "PrimaryGroup": pg, "Nature": nat, "NatureOfGroup": nog,
                "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]),
                "Parent": parent, "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]),
                "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))),
                "ClosingBalance": float(to_decimal(direct_child_text(elem, "CLOSINGBALANCE"))),
                "CompanyName": COMPANY, "FromDate": format_tally_date(FROM_DATE), "ToDate": format_tally_date(TO_DATE)
            })
    except: pass
    Ledger = pd.DataFrame(l_rows)
    ledger_meta = {r["Name"]: r for r in l_rows}

    # 2. JOURNALS (Vouchers)
    time.sleep(2) # Give Tally more time
    v_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>V</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY><SVFROMDATE TYPE='Date'>{escape(FROM_DATE)}</SVFROMDATE><SVTODATE TYPE='Date'>{escape(TO_DATE)}</SVTODATE></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"V\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, AllLedgerEntries.IsDeemedPositive</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    v_rows = []
    try:
        rv = requests.post(url, data=v_xml.encode("utf-8"), timeout=120)
        for v in ET.fromstring(xml_cleanup(rv.text)).iter():
            if strip_ns(v.tag).upper() != "VOUCHER": continue
            vt = direct_child_text(v, "VOUCHERTYPENAME")
            base_vt = vtype_map.get(vt, vt)
            if base_vt not in ACCOUNTING_VOUCHER_TYPES: continue
            
            v_date, v_num = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER")
            v_nar = first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
            
            entries = [c for c in list(v) if "LEDGERENTRIES.LIST" in strip_ns(c.tag).upper()]
            for ent in entries:
                ln = direct_child_text(ent, "LEDGERNAME")
                amt = to_decimal(direct_child_text(ent, "AMOUNT"))
                if not ln or amt == 0: continue
                is_pos = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
                signed = abs(amt) * (Decimal("-1") if is_pos else Decimal("1"))
                meta = ledger_meta.get(ln, {})
                v_rows.append({
                    "Date": v_date, "VoucherTypeName": vt, "BaseVoucherType": base_vt, "VoucherNumber": v_num,
                    "LedgerName": ln, "MasterID": meta.get("MasterID", ""), "Amount": float(signed),
                    "DrCr": "Dr" if signed < 0 else "Cr", "DebitAmount": float(abs(signed)) if signed < 0 else 0.0,
                    "CreditAmount": float(abs(signed)) if signed > 0 else 0.0,
                    "ParentLedger": meta.get("Parent", ""), "PrimaryGroup": meta.get("PrimaryGroup", ""),
                    "Nature": meta.get("Nature", ""), "NatureOfGroup": meta.get("NatureOfGroup", ""),
                    "PAN": meta.get("PAN", ""), "PartyLedgerName": direct_child_text(v, "PARTYLEDGERNAME"),
                    "PartyGSTIN": direct_child_text(v, "PARTYGSTIN"), "LedgerGSTIN": meta.get("PartyGSTIN", ""),
                    "VoucherNarration": v_nar, "CompanyName": COMPANY, "FromDate": format_tally_date(FROM_DATE), "ToDate": format_tally_date(TO_DATE)
                })
    except: pass
    Journal = pd.DataFrame(v_rows)

    # 3. STOCK ITEMS
    time.sleep(1)
    si_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SI</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"SI\"><TYPE>StockItem</TYPE><FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    si_rows = []
    try:
        rsi = requests.post(url, data=si_xml.encode("utf-8"), timeout=60)
        for elem in ET.fromstring(xml_cleanup(rsi.text)).iter():
            if strip_ns(elem.tag).upper() != "STOCKITEM": continue
            si_rows.append({
                "Name": clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME"),
                "Parent": direct_child_text(elem, "PARENT"), "Category": direct_child_text(elem, "CATEGORY"),
                "LedgerName": direct_child_text(elem, "LEDGERNAME"),
                "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))),
                "OpeningValue": float(to_decimal(direct_child_text(elem, "OPENINGVALUE"))),
                "BasicValue": float(to_decimal(direct_child_text(elem, "BASICVALUE"))),
                "BasicQty": float(to_decimal(direct_child_text(elem, "BASICQTY"))),
                "OpeningRate": float(to_decimal(direct_child_text(elem, "OPENINGRATE"))),
                "CompanyName": COMPANY, "FromDate": format_tally_date(FROM_DATE), "ToDate": format_tally_date(TO_DATE)
            })
    except: pass
    StockItem = pd.DataFrame(si_rows)

    # 4. STOCK VOUCHERS
    time.sleep(2)
    sv_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SV</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY><SVFROMDATE TYPE='Date'>{escape(FROM_DATE)}</SVFROMDATE><SVTODATE TYPE='Date'>{escape(TO_DATE)}</SVTODATE></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"SV\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, InventoryEntries.StockItemName, InventoryEntries.Amount, InventoryEntries.BilledQty, InventoryEntries.Rate, InventoryEntries.IsDeemedPositive, InventoryEntries.BatchAllocations.List</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    sv_rows = []
    try:
        rsv = requests.post(url, data=sv_xml.encode("utf-8"), timeout=120)
        for v in ET.fromstring(xml_cleanup(rsv.text)).iter():
            if strip_ns(v.tag).upper() != "VOUCHER": continue
            vt = direct_child_text(v, "VOUCHERTYPENAME")
            if "Order" in vt: continue
            vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
            inv_nodes = [c for c in list(v) if "INVENTORYENTRIES" in strip_ns(c.tag).upper()]
            for ent in inv_nodes:
                inm = direct_child_text(ent, "STOCKITEMNAME")
                if not inm: continue
                is_inward = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
                qty, amt = abs(float(to_decimal(direct_child_text(ent, "BILLEDQTY")))), abs(float(to_decimal(direct_child_text(ent, "AMOUNT"))))
                sv_rows.append({
                    "Date": vd, "VoucherTypeName": vt, "VoucherNumber": vn, "StockItemName": inm,
                    "BilledQty": qty if is_inward else -qty, "Rate": float(to_decimal(direct_child_text(ent, "RATE"))),
                    "Amount": amt if is_inward else -amt, "VoucherNarration": v_nar,
                    "CompanyName": COMPANY, "FromDate": format_tally_date(FROM_DATE), "ToDate": format_tally_date(TO_DATE)
                })
    except: pass
    StockVoucher = pd.DataFrame(sv_rows)

    # Save to cache for subsequent Power BI processes
    save_to_cache({'Journal': Journal, 'Ledger': Ledger, 'StockItem': StockItem, 'StockVoucher': StockVoucher})
