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
    "StartingFrom",
    "CurrencyName",
    "StateName",
    "Parent",
    "PartyGSTIN",
    "OpeningBalance",
    "ClosingBalance",
]

BS_PRIMARY_GROUPS = {
    "Bank Accounts",
    "Bank OD A/c",
    "Capital Account",
    "Cash-in-Hand",
    "Current Assets",
    "Deposits (Asset)",
    "Fixed Assets",
    "Investments",
    "Loans & Advances (Asset)",
    "Loans (Liability)",
    "Provisions",
    "Secured Loans",
    "Sundry Creditors",
    "Sundry Debtors",
}

PL_PRIMARY_GROUPS = {
    "Direct Expenses",
    "Duties & Taxes",
    "Indirect Expenses",
    "Indirect Incomes",
    "Purchase Accounts",
    "Sales Accounts",
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
        "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance</FETCH>"
        "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE>"
        "<COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )


def parse_ledgers(root):
    ledger_rows = []
    ledger_lookup = {}
    for elem in root.iter():
        if strip_ns(elem.tag).upper() != "LEDGER":
            continue

        name = clean_text(elem.get("NAME")) or direct_child_text(elem, "NAME")
        if not name:
            continue

        row = {
            "MasterID": clean_text(elem.get("MASTERID")) or direct_child_text(elem, "MASTERID"),
            "Name": name,
            "StartingFrom": first_non_empty_text(elem, ["STARTINGFROM"]) or first_descendant_text(elem, "STARTINGFROM"),
            "CurrencyNameRaw": first_non_empty_text(elem, ["CURRENCYNAME"]) or first_descendant_text(elem, "CURRENCYNAME"),
            "CurrencySymbolRaw": first_non_empty_text(elem, ["CURRENCYSYMBOL"]) or first_descendant_text(elem, "CURRENCYSYMBOL"),
            "CurrencyOriginalSymbolRaw": first_non_empty_text(elem, ["CURRENCYORIGINALSYMBOL"]) or first_descendant_text(elem, "CURRENCYORIGINALSYMBOL"),
            "CurrencyFormalNameRaw": first_non_empty_text(elem, ["CURRENCYFORMALNAME"]) or first_descendant_text(elem, "CURRENCYFORMALNAME"),
            "StateName": first_non_empty_text(elem, ["STATENAME"]) or first_descendant_text(elem, "STATENAME"),
            "Parent": direct_child_text(elem, "PARENT"),
            "PartyGSTIN": first_non_empty_text(elem, ["PARTYGSTIN", "GSTIN"]) or first_descendant_text(elem, "PARTYGSTIN"),
            "OpeningBalance": to_float(first_non_empty_text(elem, ["OPENINGBALANCE"]) or first_descendant_text(elem, "OPENINGBALANCE")),
            "ClosingBalance": to_float(first_non_empty_text(elem, ["CLOSINGBALANCE"]) or first_descendant_text(elem, "CLOSINGBALANCE")),
            "PrimaryGroup": first_non_empty_text(elem, ["PRIMARYGROUP"]) or first_descendant_text(elem, "PRIMARYGROUP"),
        }
        ledger_rows.append(row)
        ledger_lookup[name] = row

    for row in ledger_rows:
        if not row["PrimaryGroup"]:
            row["PrimaryGroup"] = ledger_primary_group(row["Name"], ledger_lookup)

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
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE><FETCH>Name, StartingFrom, EndingAt</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        r = requests.post(url, data=xml.encode("utf-8"), timeout=15)
        cleaned = xml_cleanup(r.text)
        root = ET.fromstring(cleaned.encode("utf-8"))
        cmp = root.find(".//COMPANY")
        if cmp is not None:
            name = clean_text(cmp.get("NAME")) or clean_text(cmp.findtext("NAME", ""))
            start = clean_text(cmp.findtext("STARTINGFROM", ""))
            end = clean_text(cmp.findtext("ENDINGAT", ""))
            return name, start, end
    except:
        pass
    return "", "", ""


def fetch_ledger_rows(host, port, company):
    url = f"http://{host}:{port}"

    if not company:
        cmp_name, _, _ = get_company_info(host, port)
        company = cmp_name

    ledger_root = parse_xml_root(post_to_tally(url, build_ledger_request_xml(company)))
    status = clean_text(first_descendant_text(ledger_root, "STATUS"))
    if status == "0":
        error_text = first_descendant_text(ledger_root, "LINEERROR") or "Tally returned STATUS=0 for ledger extract"
        raise ValueError(error_text)

    ledger_rows = parse_ledgers(ledger_root)
    return build_ledger_rows(ledger_rows)


rows = fetch_ledger_rows(
    host=HOST,
    port=PORT,
    company=COMPANY,
)

dataset = pd.DataFrame(rows)
for column in LEDGER_OUTPUT_COLUMNS:
    if column not in dataset.columns:
        dataset[column] = ""
dataset = dataset[LEDGER_OUTPUT_COLUMNS]
