import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET

# Tally Settings
HOST = "localhost"
PORT = "9000"
FROM_DATE = "20000101"
TO_DATE = "20261231"

def xml_cleanup(xml_text):
    if not xml_text: return ""
    def fix_char_ref(match):
        value = match.group(1)
        try:
            cp = int(value[1:], 16) if value.lower().startswith("x") else int(value)
        except Exception: return ""
        if cp in (9, 10, 13) or (32 <= cp <= 55295) or (57344 <= cp <= 65533) or (65536 <= cp <= 1114111):
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

def post_to_tally(url, xml):
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=60)
        r.raise_for_status()
        return r.text
    except Exception as e: return ""

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
            name = (cmp.get("NAME") or cmp.findtext("NAME", "")).strip()
            start = cmp.findtext("STARTINGFROM", "").strip()
            end = cmp.findtext("ENDINGAT", "").strip()
            return name, start, end
    except:
        pass
    return "", "", ""

def fetch_inventory_rows(host, port, company, from_date, to_date):
    if not company or not from_date or not to_date:
        cmp_name, cmp_start, cmp_end = get_company_info(host, port)
        company = company or cmp_name
        from_date = from_date or cmp_start
        to_date = to_date or cmp_end

    xml_req = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>ActualInvExtract</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
        f"<SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>"
        f"<SVFROMDATE TYPE=\"Date\">{from_date}</SVFROMDATE><SVTODATE TYPE=\"Date\">{to_date}</SVTODATE></STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"ActualInvExtract\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, InventoryEntries.*, AllInventoryEntries.*, InventoryEntriesIn.*, InventoryEntriesOut.*</FETCH>"
        "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    
    resp = post_to_tally(f"http://{host}:{port}", xml_req)
    rows = []
    if resp:
        try:
            root = ET.fromstring(xml_cleanup(resp).encode("utf-8"))
            for v in root.findall(".//VOUCHER"):
                v_type = v.findtext("VOUCHERTYPENAME", "")
                if "Order" in v_type: continue
                
                inv_nodes = [child for child in v if "INVENTORYENTRIES" in child.tag.upper()]
                for ent in inv_nodes:
                    item = ent.findtext("STOCKITEMNAME", "")
                    if not item: continue
                    
                    def cn(t):
                        if not t: return 0.0
                        try: return float(re.sub(r"[^0-9.-]", "", t.replace(",","")))
                        except: return 0.0

                    is_pos_val = ent.findtext("ISDEEMEDPOSITIVE", "No")
                    is_inward = (is_pos_val.upper() == "YES")
                    qty = abs(cn(ent.findtext("BILLEDQTY", "0")))
                    amt = abs(cn(ent.findtext("AMOUNT", "0")))
                    
                    rows.append({
                        "Date": v.findtext("DATE", ""),
                        "VoucherTypeName": v_type,
                        "VoucherNumber": v.findtext("VOUCHERNUMBER", ""),
                        "StockItemName": item.strip(),
                        "BilledQty": qty if is_inward else -qty,
                        "Rate": cn(ent.findtext("RATE", "0")),
                        "Amount": amt if is_inward else -amt,
                        "IsDeemedPositive": is_pos_val,
                        "GodownName": ent.findtext(".//GODOWNNAME", ""),
                        "BatchName": ent.findtext(".//BATCHNAME", "")
                    })
        except: pass
    return rows

if __name__ == "__main__":
    rows = fetch_inventory_rows(HOST, PORT, "", "", "")
    dataset = pd.DataFrame(rows)
    if not dataset.empty:
        print(dataset.head())
    else:
        print("No inventory movement found")
