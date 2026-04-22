# TallyPy: XML-Based Data Extractor & Explorer

TallyPy is a robust Python-based suite designed to extract accounting and inventory data from Tally using dynamic TDL over XML. It eliminates the need for external TDL files or ODBC drivers, providing a seamless bridge between Tally and data analysis tools like Power BI and Streamlit.

## 🚀 Key Features

- **100% XML Extraction**: Uses dynamic TDL injection via XML requests—no installation required in Tally.
- **Auto-Detection**: Automatically detects the active company name and financial period (Starting From/Ending At) from Tally.
- **Accounting Data**: Robust extraction of Ledgers and Vouchers (Sales, Purchase, Journal, Receipt, Payment, etc.).
- **Inventory Data**: Specialized extraction of Stock Items and Stock Vouchers, including Receipt Notes, Delivery Notes, and Stock Journals.
- **Streamlit Dashboard**: A built-in web interface (`app1.py`) to explore, filter, and download Tally data as CSV or Excel.
- **Power BI Ready**: Optimized Python scripts designed to be used directly within Power BI's `Python.Execute` data source.

## 🛠️ Components

- `app1.py`: The main Streamlit dashboard application.
- `tally_ledgers.py`: Extractor for Ledger masters.
- `tally_vouchers.py`: Extractor for Accounting transactions.
- `tally_stock_items.py`: Extractor for Stock Item masters.
- `tally_stock_vouchers.py`: Extractor for Inventory transactions (Stock Ledgers).
- `requirements.txt`: List of Python dependencies.

## 📋 Prerequisites

1. **Tally Running**: Tally must be open with at least one company loaded.
2. **Connectivity**: Ensure the Tally ODBC/XML port is enabled (Default: 9000).
   - Verify in Tally: `F1` > `Settings` > `Connectivity`.
3. **Python Environment**: Python 3.8+ installed.

## 🚀 Getting Started

1. **Clone the repository**:
   ```bash
   git clone https://github.com/SRathinaGiri/TallyPy.git
   cd TallyPy
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Explorer**:
   ```bash
   streamlit run app1.py
   ```

## 📊 Power BI Integration

You can use the scripts in this project directly in Power BI Desktop:
1. Go to **Get Data** > **Python Script**.
2. Copy the contents of the desired `.py` file.
3. Replace the `HOST` and `PORT` variables if your Tally is not on `localhost:9000`.
4. Power BI will generate a table named `dataset` containing your Tally data.

## ⚖️ License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
