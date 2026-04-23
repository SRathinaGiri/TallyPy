import re
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from xml.sax.saxutils import escape

# Power BI / Tally settings
HOST = "localhost"
PORT = "9000"
COMPANY = ""   # Leave blank to auto-detect first company

LEDGER_OUTPUT_COLUMNS = [
    "MasterID",
    "Name",
    "PrimaryGroup",
    "Nature",
    "NatureOfGroup",
    "PAN",
    "StartingFrom",
    "CurrencyName",
    "StateName",
    "Parent",
    "PartyGSTIN",
    "OpeningBalance",
    "ClosingBalance",
    "CompanyName",
    "FromDate",
    "ToDate",
]

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

PRIMARY_GROUPS = BS_PRIMARY_GROUPS | PL_PRIMARY_GROUPS

CURRENCY_SYMBOL_FALLBACKS = {
    "INR": "₹",
    "INDIAN RUPEE": "₹",
    "RUPEES": "₹",
    "RUPEE": "₹",
    "RS": "₹",
    "RS.": "₹",
    "USD": "$",
    "US DOLLAR": "$",
    "DOLLAR": "$",
    "EUR": "€",
    "EURO": "€",
    "GBP": "£",
    "POUND": "£",
    "POUND STERLING": "£",
    "AED": "د.إ",
    "DIRHAM": "د.إ",
    "": "",
}

CURRENCY_NAME_FALLBACKS = {
    "INR": "INR",
    "INDIAN RUPEE": "INR",
    "USD": "USD",
    "US DOLLAR": "USD",
    "EUR": "EUR",
    "EURO": "EUR",
    "GBP": "GBP",
    "POUND": "GBP",
    "POUND STERLING": "GBP",
    "AED": "AED",
    "DIRHAM": "AED",
}


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


def first_non_empty_text(elem, names):
    for name in names:
        value = direct_child_text(elem, name)
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


def detect_company_name(root):
    for elem in root.iter():
        if strip_ns(elem.tag).upper() == "COMPANY":
            return clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
    return ""


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


def ledger_primary_group(ledger_name, ledger_meta):
    seen = set()
    current = clean_text(ledger_name)
    while current and current not in seen:
        seen.add(current)
        meta = ledger_meta.get(current, {})
        parent = clean_text(meta.get("Parent", ""))
        if not parent:
            return ""
        if parent in PRIMARY_GROUPS:
            return parent
        current = parent
    return ""


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


def fetch_tally_metadata(url, company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")
    
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MetadataFetch</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllGroups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    
    group_map = {}
    try:
        resp = post_to_tally(url, xml)
        root = parse_xml_root(resp)
        for g in root.iter():
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
    return group_map


def build_ledger_request_xml(company):
    static_vars = ["<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"]
    if company:
        static_vars.append(f"<SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY>")

    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyLedgers</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES>{''.join(static_vars)}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyLedgers\"><TYPE>Ledger</TYPE>"
        "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH>"
        "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE>"
        "<COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def parse_ledgers(root, group_map=None):
    ledger_rows = []
    ledger_lookup = {}
    if group_map is None:
        group_map = {}
        
    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "LEDGER":
            continue

        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name:
            continue

        parent = direct_child_text(elem, "PARENT")
        g_info = group_map.get(parent, {})
        nature_of_group = g_info.get("Nature", "")
        primary_group = g_info.get("PrimaryGroup", "") or first_non_empty_text(elem, ["PRIMARYGROUP"]) or first_descendant_text(elem, "PRIMARYGROUP")

        row = {
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"),
            "Name": name,
            "PrimaryGroup": primary_group,
            "Nature": "",
            "NatureOfGroup": nature_of_group,
            "PAN": first_non_empty_text(elem, ["INCOMETAXNUMBER", "PAN"]) or first_descendant_text(elem, "INCOMETAXNUMBER"),
            "StartingFrom": first_non_empty_text(elem, ["STARTINGFROM"]) or first_descendant_text(elem, "STARTINGFROM"),
            "CurrencyNameRaw": first_non_empty_text(elem, ["CURRENCYNAME"]) or first_descendant_text(elem, "CURRENCYNAME"),
            "CurrencySymbolRaw": first_non_empty_text(elem, ["CURRENCYSYMBOL"]) or first_descendant_text(elem, "CURRENCYSYMBOL"),
            "CurrencyOriginalSymbolRaw": first_non_empty_text(elem, ["CURRENCYORIGINALSYMBOL"]) or first_descendant_text(elem, "CURRENCYORIGINALSYMBOL"),
            "CurrencyFormalNameRaw": first_non_empty_text(elem, ["CURRENCYFORMALNAME"]) or first_descendant_text(elem, "CURRENCYFORMALNAME"),
            "StateName": first_non_empty_text(elem, ["STATENAME"]) or first_descendant_text(elem, "STATENAME"),
            "Parent": parent,
            "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]) or first_descendant_text(elem, "PARTYGSTIN"),
            "OpeningBalance": to_float(first_non_empty_text(elem, ["OPENINGBALANCE"]) or first_descendant_text(elem, "OPENINGBALANCE")),
            "ClosingBalance": to_float(first_non_empty_text(elem, ["CLOSINGBALANCE"]) or first_descendant_text(elem, "CLOSINGBALANCE")),
        }
        ledger_rows.append(row)
        ledger_lookup[name] = row

    for row in ledger_rows:
        if not row["PrimaryGroup"]:
            row["PrimaryGroup"] = ledger_primary_group(row["Name"], ledger_lookup)
        
        pg = row["PrimaryGroup"]
        if not row["NatureOfGroup"] and pg:
            row["NatureOfGroup"] = group_map.get(pg, {}).get("Nature", "")
        
        if row["NatureOfGroup"]:
            n_val = row["NatureOfGroup"].lower()
            if n_val in ["assets", "liabilities"]: row["Nature"] = "BS"
            elif n_val in ["income", "expenses"]: row["Nature"] = "PL"
            
        if not row["Nature"] and pg:
            bs_pl, nog = nature_from_primary_group(pg)
            row["Nature"] = bs_pl
            row["NatureOfGroup"] = nog

    return ledger_rows


def build_ledger_rows(ledger_rows):
    rows = []
    for source_row in sorted(ledger_rows, key=lambda r: (int(r.get("MasterID") or 0), r.get("Name", ""))):
        row = dict(source_row)
        currency_name_raw = clean_text(str(row.get("CurrencyNameRaw", "")))
        currency_symbol_raw = clean_text(str(row.get("CurrencySymbolRaw", "")))
        currency_original_symbol_raw = clean_text(str(row.get("CurrencyOriginalSymbolRaw", "")))
        currency_formal_name_raw = clean_text(str(row.get("CurrencyFormalNameRaw", "")))

        stable_currency_key = currency_formal_name_raw or currency_name_raw
        stable_currency_key_upper = stable_currency_key.upper()

        currency_name = CURRENCY_NAME_FALLBACKS.get(stable_currency_key_upper, stable_currency_key)
        symbol = currency_symbol_raw or currency_original_symbol_raw
        fallback_key = stable_currency_key_upper or currency_name.upper()
        if not symbol or symbol == "?":
            symbol = CURRENCY_SYMBOL_FALLBACKS.get(fallback_key, symbol)
        row["CurrencyName"] = symbol
        rows.append({column: row.get(column, "") for column in LEDGER_OUTPUT_COLUMNS})
    return rows


def get_company_info(host, port):
    url = f"http://{host}:{port}"
    xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE>"
        "<FETCH>Name, StartingFrom, EndingAt, Guid</FETCH>"
        "<FILTER>IsActiveCompany</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=10)
        cleaned = xml_cleanup(r.text)
        root = ET.fromstring(cleaned.encode("utf-8"))
        
        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name:
                    return name, start, end

        for cmp in root.iter():
            if strip_ns(cmp.tag).upper() == "COMPANY":
                name = clean_text(cmp.get("NAME")) or direct_child_text(cmp, "NAME")
                start = direct_child_text(cmp, "STARTINGFROM")
                end = direct_child_text(cmp, "ENDINGAT")
                if name:
                    return name, start, end
    except:
        pass
    return "", "", ""


def fetch_ledger_rows(host, port, company):
    url = f"http://{host}:{port}"

    cmp_name, cmp_start, cmp_end = get_company_info(host, port)
    if not company:
        company = cmp_name

    group_map = fetch_tally_metadata(url, company)
    
    ledger_root = parse_xml_root(post_to_tally(url, build_ledger_request_xml(company)))
    status = clean_text(first_descendant_text(ledger_root, "STATUS"))
    if status == "0":
        error_text = first_descendant_text(ledger_root, "LINEERROR") or "Tally returned STATUS=0 for ledger extract"
        raise ValueError(error_text)

    ledger_rows = parse_ledgers(ledger_root, group_map)
    return build_ledger_rows(ledger_rows), company, cmp_start, cmp_end


rows, final_company, final_start, final_end = fetch_ledger_rows(
    host=HOST,
    port=PORT,
    company=COMPANY,
)

dataset = pd.DataFrame(rows)
dataset["CompanyName"] = final_company
dataset["FromDate"] = format_tally_date(final_start)
dataset["ToDate"] = format_tally_date(final_end)

for column in LEDGER_OUTPUT_COLUMNS:
    if column not in dataset.columns:
        dataset[column] = ""
dataset = dataset[LEDGER_OUTPUT_COLUMNS]
