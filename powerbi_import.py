import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
import os
import time
import tempfile
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

# Helper functions
def strip_ns(tag):
    if not isinstance(tag, str): return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean_text(text):
    if text is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text)).strip()

def xml_cleanup(xml_text):
    if not xml_text: return ""
    xml_text = re.sub(r"&#(x[0-9A-Fa-f]+|\d+);", lambda m: chr(int(m.group(1)[1:], 16)) if m.group(1).startswith('x') else chr(int(m.group(1))), xml_text, flags=re.IGNORECASE)
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
    xml = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>MyC</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"MyC\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION><SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=10)
        cleaned = xml_cleanup(r.text)
        root = ET.fromstring(cleaned.encode("utf-8"))
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name: return name, start, end
        
        # Fallback to broad search if filter failed
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name: return name, start, end
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

# Caching & Locking
cache_dir = tempfile.gettempdir()
cache_files = {
    'Journal': os.path.join(cache_dir, f"tally_Journal_{PORT}.csv"),
    'Ledger': os.path.join(cache_dir, f"tally_Ledger_{PORT}.csv"),
    'StockItem': os.path.join(cache_dir, f"tally_StockItem_{PORT}.csv"),
    'StockVoucher': os.path.join(cache_dir, f"tally_StockVoucher_{PORT}.csv")
}
lock_file = os.path.join(cache_dir, f"tally_lock_{PORT}.lock")

def get_data():
    all_exist = all(os.path.exists(f) for f in cache_files.values())
    if all_exist:
        mtimes = [os.path.getmtime(f) for f in cache_files.values()]
        if all((time.time() - mt) < 300 for mt in mtimes):
            try: return {name: pd.read_csv(f) for name, f in cache_files.items()}
            except: pass
    
    for _ in range(120):
        try:
            fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            if (time.time() - os.path.getmtime(lock_file)) > 180:
                try: os.remove(lock_file)
                except: pass
            time.sleep(1)
            if all(os.path.exists(f) for f in cache_files.values()):
                try: return {name: pd.read_csv(f) for name, f in cache_files.items()}
                except: pass
    else:
        if all(os.path.exists(f) for f in cache_files.values()):
             return {name: pd.read_csv(f) for name, f in cache_files.items()}
        return None

    try:
        url = f"http://{HOST}:{PORT}"
        c_name, s_dt, e_dt = get_company_info(HOST, PORT)
        comp = COMPANY or c_name
        
        # Safe Date Fallbacks
        now = datetime.now()
        if now.month < 4:
            def_start, def_end = f"{now.year-1}0401", f"{now.year}0331"
        else:
            def_start, def_end = f"{now.year}0401", f"{now.year+1}0331"
            
        f_dt = str(FROM_DATE or s_dt or def_start).strip()
        t_dt = str(TO_DATE or e_dt or def_end).strip()
        
        # Ensure dates are exactly 8 digits
        if not re.fullmatch(r"\d{8}", f_dt): f_dt = def_start
        if not re.fullmatch(r"\d{8}", t_dt): t_dt = def_end

        v_map, g_map = fetch_metadata(url, comp)

        # 1. LEDGERS
        l_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>L</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"L\"><TYPE>Ledger</TYPE><FETCH>Name, Parent, PartyGSTIN, MasterID, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
        l_rows = []
        rl = requests.post(url, data=l_xml.encode("utf-8"), timeout=60)
        for elem in ET.fromstring(xml_cleanup(rl.text)).iter():
            if strip_ns(elem.tag).upper() != "LEDGER": continue
            n = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
            if not n: continue
            p = direct_child_text(elem, "PARENT")
            pg = g_map.get(p, {}).get("PrimaryGroup", "") or direct_child_text(elem, "PRIMARYGROUP")
            nog = g_map.get(p, {}).get("Nature", "")
            nat = "BS" if nog and nog.lower() in ["assets", "liabilities"] else ("PL" if nog and nog.lower() in ["income", "expenses"] else "")
            l_rows.append({"MasterID": elem.get("MASTERID") or direct_child_text(elem, "MASTERID"), "Name": n, "PrimaryGroup": pg, "Nature": nat, "NatureOfGroup": nog, "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]), "Parent": p, "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]), "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))), "ClosingBalance": float(to_decimal(direct_child_text(elem, "CLOSINGBALANCE"))), "CompanyName": comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
        l_df = pd.DataFrame(l_rows)
        l_meta = {r["Name"]: r for r in l_rows}

        # CHUNKING LOGIC
        v_rows, sv_rows = [], []
        d1 = datetime.strptime(f_dt, "%Y%m%d")
        d2 = datetime.strptime(t_dt, "%Y%m%d")
        curr = d1
        while curr <= d2:
            chunk_start = curr.strftime("%Y%m%d")
            chunk_end = (curr + timedelta(days=31)).replace(day=1) - timedelta(days=1)
            if chunk_end > d2: chunk_end = d2
            chunk_end_str = chunk_end.strftime("%Y%m%d")
            sv_chunk = f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(comp)}</SVCURRENTCOMPANY><SVFROMDATE TYPE='Date'>{chunk_start}</SVFROMDATE><SVTODATE TYPE='Date'>{chunk_end_str}</SVTODATE></STATICVARIABLES>"
            
            # JOURNALS
            v_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>V</ID></HEADER><BODY><DESC>{sv_chunk}<TDL><TDLMESSAGE><COLLECTION NAME=\"V\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, AllLedgerEntries.IsDeemedPositive</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
            rv = requests.post(url, data=v_xml.encode("utf-8"), timeout=120)
            for v in ET.fromstring(xml_cleanup(rv.text)).iter():
                if strip_ns(v.tag).upper() != "VOUCHER": continue
                vt = direct_child_text(v, "VOUCHERTYPENAME")
                if v_map.get(vt, vt) not in ACCOUNTING_VOUCHER_TYPES: continue
                vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
                for ent in [c for c in list(v) if "LEDGERENTRIES.LIST" in strip_ns(c.tag).upper()]:
                    ln = direct_child_text(ent, "LEDGERNAME")
                    amt = to_decimal(direct_child_text(ent, "AMOUNT"))
                    if not ln or amt == 0: continue
                    signed = abs(amt) * (Decimal("-1") if direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES" else Decimal("1"))
                    m = l_meta.get(ln, {})
                    v_rows.append({"Date": vd, "VoucherTypeName": vt, "BaseVoucherType": v_map.get(vt, vt), "VoucherNumber": vn, "LedgerName": ln, "MasterID": m.get("MasterID", ""), "Amount": float(signed), "DrCr": "Dr" if signed < 0 else "Cr", "DebitAmount": float(abs(signed)) if signed < 0 else 0.0, "CreditAmount": float(abs(signed)) if signed > 0 else 0.0, "ParentLedger": m.get("Parent", ""), "PrimaryGroup": m.get("PrimaryGroup", ""), "Nature": m.get("Nature", ""), "NatureOfGroup": m.get("NatureOfGroup", ""), "PAN": m.get("PAN", ""), "PartyLedgerName": direct_child_text(v, "PARTYLEDGERNAME"), "PartyGSTIN": direct_child_text(v, "PARTYGSTIN"), "LedgerGSTIN": m.get("PartyGSTIN", ""), "VoucherNarration": v_nar, "CompanyName": comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
            
            # STOCK VOUCHERS
            sv_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SV</ID></HEADER><BODY><DESC>{sv_chunk}<TDL><TDLMESSAGE><COLLECTION NAME=\"SV\"><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration, InventoryEntries.StockItemName, InventoryEntries.Amount, InventoryEntries.BilledQty, InventoryEntries.Rate, InventoryEntries.IsDeemedPositive</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
            rsv = requests.post(url, data=sv_xml.encode("utf-8"), timeout=120)
            for v in ET.fromstring(xml_cleanup(rsv.text)).iter():
                if strip_ns(v.tag).upper() != "VOUCHER" or "Order" in direct_child_text(v, "VOUCHERTYPENAME"): continue
                vd, vn, v_nar = format_tally_date(direct_child_text(v, "DATE")), direct_child_text(v, "VOUCHERNUMBER"), first_non_empty_text(v, ["NARRATION", "VOUCHERNARRATION"])
                for ent in [c for c in list(v) if "INVENTORYENTRIES" in strip_ns(c.tag).upper()]:
                    inm = direct_child_text(ent, "STOCKITEMNAME")
                    if not inm: continue
                    is_in = direct_child_text(ent, "ISDEEMEDPOSITIVE").upper() == "YES"
                    q, a = abs(float(to_decimal(direct_child_text(ent, "BILLEDQTY")))), abs(float(to_decimal(direct_child_text(ent, "AMOUNT"))))
                    sv_rows.append({"Date": vd, "VoucherTypeName": direct_child_text(v, "VOUCHERTYPENAME"), "VoucherNumber": vn, "StockItemName": inm, "BilledQty": q if is_in else -q, "Rate": float(to_decimal(direct_child_text(ent, "RATE"))), "Amount": a if is_in else -a, "VoucherNarration": v_nar, "CompanyName": comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
            
            curr = chunk_end + timedelta(days=1)
            time.sleep(1)

        # 4. STOCK ITEMS
        si_xml = f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>SI</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>{escape(comp)}</SVCURRENTCOMPANY></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME=\"SI\"><TYPE>StockItem</TYPE><FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
        si_rows = []
        rsi = requests.post(url, data=si_xml.encode("utf-8"), timeout=60)
        for elem in ET.fromstring(xml_cleanup(rsi.text)).iter():
            if strip_ns(elem.tag).upper() != "STOCKITEM": continue
            si_rows.append({"Name": clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME"), "Parent": direct_child_text(elem, "PARENT"), "Category": direct_child_text(elem, "CATEGORY"), "LedgerName": direct_child_text(elem, "LEDGERNAME"), "OpeningBalance": float(to_decimal(direct_child_text(elem, "OPENINGBALANCE"))), "OpeningValue": float(to_decimal(direct_child_text(elem, "OPENINGVALUE"))), "BasicValue": float(to_decimal(direct_child_text(elem, "BASICVALUE"))), "BasicQty": float(to_decimal(direct_child_text(elem, "BASICQTY"))), "OpeningRate": float(to_decimal(direct_child_text(elem, "OPENINGRATE"))), "CompanyName": comp, "FromDate": format_tally_date(f_dt), "ToDate": format_tally_date(t_dt)})
        
        final_dfs = {'Journal': pd.DataFrame(v_rows), 'Ledger': l_df, 'StockItem': pd.DataFrame(si_rows), 'StockVoucher': pd.DataFrame(sv_rows)}
        for name, df in final_dfs.items(): df.to_csv(cache_files[name], index=False)
        return final_dfs
    finally:
        try: os.remove(lock_file)
        except: pass

data = get_data()
Journal, Ledger, StockItem, StockVoucher = data['Journal'], data['Ledger'], data['StockItem'], data['StockVoucher']
