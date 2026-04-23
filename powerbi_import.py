import sys
import os

# Add the project directory to sys.path so we can import our local extractor scripts
project_dir = r"C:\Users\srath\OneDrive\MyProjects\Python\TallyXML"
if project_dir not in sys.path:
    sys.path.append(project_dir)

# Import the individual extractor scripts.
# Note: These scripts execute their data fetching logic upon being imported.
import tally_vouchers
import tally_ledgers
import tally_stock_vouchers
import tally_stock_items

# Assign the 'dataset' from each module to the specific names requested for Power BI.
# Power BI's Python connector will detect these 4 DataFrames in the global scope.
Journal = tally_vouchers.dataset
Ledger = tally_ledgers.dataset
StockVoucher = tally_stock_vouchers.dataset
StockItem = tally_stock_items.dataset

# When importing this script into Power BI:
# 1. Use the "Python script" data connector.
# 2. Paste the contents of this script (or use the one-liner pointing to this file).
# 3. Power BI will present a navigator window showing Journal, Ledger, StockVoucher, and StockItem tables.
