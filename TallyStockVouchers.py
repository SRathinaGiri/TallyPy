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
    return xml_text

def post_to_tally(xml):
    try:
        r = requests.post(f"http://{HOST}:{PORT}", data=xml.encode("utf-8"), timeout=60)
        r.raise_for_status()
        return r.text
    except Exception as e: return ""

# 1. Fetch
xml_req = (
    "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
    "<TYPE>COLLECTION</TYPE><ID>FullFlow</ID></HEADER><BODY><DESC>"
    f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    f"<SVFROMDATE TYPE=\"Date\">{FROM_DATE}</SVFROMDATE><SVTODATE TYPE=\"Date\">{TO_DATE}</SVTODATE></STATICVARIABLES>"
    "<TDL><TDLMESSAGE><COLLECTION NAME=\"FullFlow\"><TYPE>Voucher</TYPE>"
    "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, InventoryEntries.*, AllInventoryEntries.*, InventoryEntriesIn.*, InventoryEntriesOut.*</FETCH>"
    "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
)

resp = post_to_tally(xml_req)
rows = []

if resp:
    try:
        root = ET.fromstring(xml_cleanup(resp).encode("utf-8"))
        for v in root.findall(".//VOUCHER"):
            v_type = v.findtext("VOUCHERTYPENAME", "")
            if "Order" in v_type: continue
            
            # Search for ANY tag with inventory lines
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
                    "IsDeemedPositive": is_pos_val, # Added column
                    "GodownName": ent.findtext(".//GODOWNNAME", ""),
                    "BatchName": ent.findtext(".//BATCHNAME", "")
                })
    except: pass

# 3. Output
if not rows:
    dataset = pd.DataFrame([{"Date": "20000101", "VoucherTypeName": "DEBUG", "VoucherNumber": "No inventory movement found"}])
else:
    dataset = pd.DataFrame(rows)

if __name__ == "__main__":
    if not dataset.empty: print(dataset.head())
