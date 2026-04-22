# TallyPy Project Context

## Project State as of April 22, 2026
This project is a Python and Power BI suite for Tally XML data extraction and statutory auditing.

## Foundational Mandates
1. **100% XML Extraction**: Do NOT use ODBC or external TDL files. All extraction must use dynamic TDL injection via XML requests to `localhost:9000`.
2. **Standardized Cleanup**: All XML responses must pass through the `xml_cleanup` function (found in `tally_vouchers.py`) to handle illegal control characters and unescaped ampersands.
3. **Auto-Detection**: Use the `get_company_info` logic to dynamically detect Company Name and Financial Year dates if not provided by the user.

## Data Logic
- **Inventory Signs**: Inward movements (Receipts/Purchases) are Positive (+). Outward movements (Sales/Deliveries) are Negative (-).
- **Inventory Source**: Use "Greedy Search" to scan all XML tags containing "INVENTORYENTRIES" to capture Receipt/Delivery notes and Stock Journals.
- **Balance Sheet**: Follow Schedule 3 statutory hierarchy. Use the `Schedule3 Value` measure which handles automatic re-classification (e.g., Credit Debtors move to Liabilities).

## Key Files
- `app1.py`: Streamlit Dashboard.
- `tally_stock_vouchers.py`: Latest robust inventory extractor.
- `Schedule3_Implementation.md`: Reference for all DAX measures and table schemas.
- `ClosingBalance_Audit.md`: Audit confirmation of the dynamic closing balance logic.
