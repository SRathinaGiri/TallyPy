import json

with open(r'C:\Users\acer\OneDrive\MyProjects\Python\TallyXML\CaseStudy01.SemanticModel\model.bim.bak', 'r', encoding='utf-8') as f:
    model = json.load(f)

# 1. Update PBI_QueryOrder annotation
for annotation in model['model']['annotations']:
    if annotation['name'] == 'PBI_QueryOrder':
        annotation['value'] = json.dumps(["Journal","Ledger","StockItem","Lakhs","StockVoucher","Schedule3Layout","Schedule3Mapping"])

# 2. Add Measure "Schedule3 Value" to "_Measures" table
schedule3_measure = {
  "name": "Schedule3 Value",
  "expression": [
    "VAR CurrentLineItem = SELECTEDVALUE ( Schedule3Layout[LineItem] )",
    "VAR IsReservesAndSurplus = ( CurrentLineItem = \"Reserves and Surplus\" )",
    "VAR MappedGroups = CALCULATETABLE ( VALUES ( Schedule3Mapping[PrimaryGroup] ), KEEPFILTERS ( Schedule3Layout ) )",
    "VAR StandardBalance = CALCULATE ( SUM ( Ledger[ClosingBalance] ), TREATAS ( MappedGroups, Ledger[PrimaryGroup] ) )",
    "VAR NetProfitLoss = IF ( IsReservesAndSurplus, CALCULATE ( SUM ( Ledger[ClosingBalance] ), Ledger[Type] IN { \"Incomes\", \"Expenses\" } ), 0 )",
    "VAR TotalValue = StandardBalance + NetProfitLoss",
    "RETURN IF ( TotalValue = 0, BLANK (), ABS ( TotalValue ) )"
  ],
  "formatString": "#,##0.00",
  "lineageTag": "a905545b-d363-4876-8804-d538e12d46e2"
}

for table in model['model']['tables']:
    if table['name'] == '_Measures':
        if 'measures' not in table:
            table['measures'] = []
        table['measures'].append(schedule3_measure)
        break

# 3. Add Schedule3Layout table
layout_table = {
  "name": "Schedule3Layout",
  "lineageTag": "f905545b-d363-4876-8804-d538e12d46e2",
  "columns": [
    {
      "name": "SortOrder",
      "dataType": "int64",
      "formatString": "0",
      "lineageTag": "37f4044b-f363-4552-b804-d538e12d46e2",
      "sourceColumn": "[Value1]",
      "summarizeBy": "sum"
    },
    {
      "name": "Heading",
      "dataType": "string",
      "lineageTag": "22f7744b-f363-4552-b804-d538e12d46e2",
      "sourceColumn": "[Value2]",
      "summarizeBy": "none"
    },
    {
      "name": "SubHeading",
      "dataType": "string",
      "lineageTag": "11f5544b-f363-4552-b804-d538e12d46e2",
      "sourceColumn": "[Value3]",
      "summarizeBy": "none"
    },
    {
      "name": "LineItem",
      "dataType": "string",
      "lineageTag": "00f4444b-f363-4552-b804-d538e12d46e2",
      "sourceColumn": "[Value4]",
      "sortByColumn": "SortOrder",
      "summarizeBy": "none"
    }
  ],
  "partitions": [
    {
      "name": "Schedule3Layout",
      "mode": "import",
      "source": {
        "expression": [
          "DATATABLE ( \"SortOrder\", INTEGER, \"Heading\", STRING, \"SubHeading\", STRING, \"LineItem\", STRING, { { 10, \"I. EQUITY AND LIABILITIES\", \"(1) Shareholders' Funds\", \"Share Capital\" }, { 20, \"I. EQUITY AND LIABILITIES\", \"(1) Shareholders' Funds\", \"Reserves and Surplus\" }, { 30, \"I. EQUITY AND LIABILITIES\", \"(2) Non-Current Liabilities\", \"Long-term borrowings\" }, { 40, \"I. EQUITY AND LIABILITIES\", \"(2) Non-Current Liabilities\", \"Deferred tax liabilities (Net)\" }, { 50, \"I. EQUITY AND LIABILITIES\", \"(2) Non-Current Liabilities\", \"Other Long term liabilities\" }, { 60, \"I. EQUITY AND LIABILITIES\", \"(2) Non-Current Liabilities\", \"Long-term provisions\" }, { 70, \"I. EQUITY AND LIABILITIES\", \"(3) Current Liabilities\", \"Short-term borrowings\" }, { 80, \"I. EQUITY AND LIABILITIES\", \"(3) Current Liabilities\", \"Trade payables\" }, { 90, \"I. EQUITY AND LIABILITIES\", \"(3) Current Liabilities\", \"Other current liabilities\" }, { 100, \"I. EQUITY AND LIABILITIES\", \"(3) Current Liabilities\", \"Short-term provisions\" }, { 110, \"II. ASSETS\", \"(1) Non-current assets\", \"Property, Plant and Equipment\" }, { 120, \"II. ASSETS\", \"(1) Non-current assets\", \"Intangible assets\" }, { 130, \"II. ASSETS\", \"(1) Non-current assets\", \"Non-current investments\" }, { 140, \"II. ASSETS\", \"(1) Non-current assets\", \"Deferred tax assets (net)\" }, { 150, \"II. ASSETS\", \"(1) Non-current assets\", \"Long-term loans and advances\" }, { 160, \"II. ASSETS\", \"(1) Non-current assets\", \"Other non-current assets\" }, { 170, \"II. ASSETS\", \"(2) Current assets\", \"Current investments\" }, { 180, \"II. ASSETS\", \"(2) Current assets\", \"Inventories\" }, { 190, \"II. ASSETS\", \"(2) Current assets\", \"Trade receivables\" }, { 200, \"II. ASSETS\", \"(2) Current assets\", \"Cash and cash equivalents\" }, { 210, \"II. ASSETS\", \"(2) Current assets\", \"Short-term loans and advances\" }, { 220, \"II. ASSETS\", \"(2) Current assets\", \"Other current assets\" } } )"
        ],
        "type": "calculated"
      }
    }
  ]
}
model['model']['tables'].append(layout_table)

# 4. Add Schedule3Mapping table
mapping_table = {
  "name": "Schedule3Mapping",
  "lineageTag": "e905545b-d363-4876-8804-d538e12d46e2",
  "columns": [
    {
      "name": "PrimaryGroup",
      "dataType": "string",
      "lineageTag": "c905545b-d363-4876-8804-d538e12d46e2",
      "sourceColumn": "[Value1]",
      "summarizeBy": "none"
    },
    {
      "name": "LineItem",
      "dataType": "string",
      "lineageTag": "b905545b-d363-4876-8804-d538e12d46e2",
      "sourceColumn": "[Value2]",
      "summarizeBy": "none"
    }
  ],
  "partitions": [
    {
      "name": "Schedule3Mapping",
      "mode": "import",
      "source": {
        "expression": [
          "DATATABLE ( \"PrimaryGroup\", STRING, \"LineItem\", STRING, { { \"Capital Account\", \"Share Capital\" }, { \"Reserves & Surplus\", \"Reserves and Surplus\" }, { \"Secured Loans\", \"Long-term borrowings\" }, { \"Unsecured Loans\", \"Long-term borrowings\" }, { \"Bank OD A/c\", \"Short-term borrowings\" }, { \"Loans (Liability)\", \"Short-term borrowings\" }, { \"Sundry Creditors\", \"Trade payables\" }, { \"Duties & Taxes\", \"Other current liabilities\" }, { \"Provisions\", \"Short-term provisions\" }, { \"Fixed Assets\", \"Property, Plant and Equipment\" }, { \"Investments\", \"Non-current investments\" }, { \"Sundry Debtors\", \"Trade receivables\" }, { \"Bank Accounts\", \"Cash and cash equivalents\" }, { \"Cash-in-Hand\", \"Cash and cash equivalents\" }, { \"Loans & Advances (Asset)\", \"Short-term loans and advances\" }, { \"Current Assets\", \"Other current assets\" }, { \"Deposits (Asset)\", \"Other current assets\" } } )"
        ],
        "type": "calculated"
      }
    }
  ]
}
model['model']['tables'].append(mapping_table)

# 5. Add StockVoucher table
with open('escaped_python.txt', 'r', encoding='utf-8') as f:
    escaped_python = f.read()

stock_voucher_table = {
  "name": "StockVoucher",
  "annotations": [
    {
      "name": "PBI_NavigationStepName",
      "value": "Navigation"
    },
    {
      "name": "PBI_ResultType",
      "value": "Table"
    }
  ],
  "columns": [
    {
      "name": "Date",
      "dataType": "dateTime",
      "formatString": "Short Date",
      "lineageTag": "f1d831b6-6e11-416a-9d33-2ea36c7f2525",
      "sourceColumn": "Date",
      "summarizeBy": "none"
    },
    {
      "name": "VoucherTypeName",
      "dataType": "string",
      "lineageTag": "e1312316-7e2c-4f1d-869e-89dc18ebc29b",
      "sourceColumn": "VoucherTypeName",
      "summarizeBy": "none"
    },
    {
      "name": "VoucherNumber",
      "dataType": "string",
      "lineageTag": "d518bc26-1f5d-4458-8ffb-525d0d27db5e",
      "sourceColumn": "VoucherNumber",
      "summarizeBy": "none"
    },
    {
      "name": "StockItemName",
      "dataType": "string",
      "lineageTag": "c2ecf448-5c62-4b8f-a03f-92af9ae34ad1",
      "sourceColumn": "StockItemName",
      "summarizeBy": "none"
    },
    {
      "name": "BilledQty",
      "dataType": "double",
      "lineageTag": "b025e796-d773-4b08-83ec-a3dce217c300",
      "sourceColumn": "BilledQty",
      "summarizeBy": "sum"
    },
    {
      "name": "Rate",
      "dataType": "double",
      "lineageTag": "a1d831b6-6e11-416a-9d33-2ea36c7f2526",
      "sourceColumn": "Rate",
      "summarizeBy": "sum"
    },
    {
      "name": "Amount",
      "dataType": "double",
      "lineageTag": "9b986408-13ec-4dd1-a41a-383171ca1fb8",
      "sourceColumn": "Amount",
      "summarizeBy": "sum"
    },
    {
      "name": "GodownName",
      "dataType": "string",
      "lineageTag": "82a71e52-cbd4-4450-91d5-9a226878056d",
      "sourceColumn": "GodownName",
      "summarizeBy": "none"
    },
    {
      "name": "BatchName",
      "dataType": "string",
      "lineageTag": "7c5eb3ed-0f4e-4804-822e-33f15b525c36",
      "sourceColumn": "BatchName",
      "summarizeBy": "none"
    },
    {
      "name": "VoucherNarration",
      "dataType": "string",
      "lineageTag": "63d4cbe6-1ad8-43fd-860a-903f0e340fd5",
      "sourceColumn": "VoucherNarration",
      "summarizeBy": "none"
    },
    {
      "name": "CompanyName",
      "dataType": "string",
      "lineageTag": "5c7a0530-d62d-4872-a154-d851e259eeaf",
      "sourceColumn": "CompanyName",
      "summarizeBy": "none"
    }
  ],
  "lineageTag": "d905545b-d363-4876-8804-d538e12d46e2",
  "partitions": [
    {
      "name": "StockVoucher",
      "mode": "import",
      "source": {
        "expression": [
          "let",
          "    Source = Python.Execute(\"" + escaped_python + "\"),",
          "    dataset = Source{[Name=\"dataset\"]}[Value],",
          "    #\"Changed Type\" = Table.TransformColumnTypes(dataset,{{\"Date\", type date}, {\"VoucherTypeName\", type text}, {\"VoucherNumber\", type text}, {\"StockItemName\", type text}, {\"BilledQty\", type number}, {\"Rate\", type number}, {\"Amount\", type number}, {\"GodownName\", type text}, {\"BatchName\", type text}, {\"VoucherNarration\", type text}, {\"CompanyName\", type text}})",
          "in",
          "    #\"Changed Type\""
        ],
        "type": "m"
      }
    }
  ]
}
model['model']['tables'].append(stock_voucher_table)

with open(r'C:\Users\acer\OneDrive\MyProjects\Python\TallyXML\CaseStudy01.SemanticModel\model.bim', 'w', encoding='utf-8') as f:
    json.dump(model, f, indent=2)
