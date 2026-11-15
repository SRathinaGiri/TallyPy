# Tally Exporter

Tally Exporter is a standalone Windows application for extracting accounting
and inventory data from Tally through dynamic TDL over XML. Download and run
the single `TallyExporter.exe`; no installer or additional application runtime
is required.

[Download the latest Windows release](https://github.com/SRathinaGiri/TallyPy/releases/latest)

## 🚀 Key Features

- **100% XML Extraction**: Uses dynamic TDL injection via XML requests—no installation required in Tally.
- **Auto-Detection**: Automatically detects the active company name and financial period (Starting From/Ending At) from Tally.
- **Accounting Data**: Robust extraction of Ledgers and Vouchers (Sales, Purchase, Journal, Receipt, Payment, etc.).
- **Inventory Data**: Specialized extraction of Stock Items and Stock Vouchers, including Receipt Notes, Delivery Notes, and Stock Journals.
- **Standalone Windows App**: Native HMG/Harbour interface with virtual data grids and CSV export.
- **Power BI Ready**: Optimized Python scripts designed to be used directly within Power BI's `Python.Execute` data source.

## 🛠️ Components

- `HMG/`: Source for the main standalone Windows application.
- `tally_xml_exporter.py`: Optional Python desktop utility with CSV-backed pagination for large exports.
- `tally_ledgers.py`: Extractor for Ledger masters.
- `tally_vouchers.py`: Extractor for Accounting transactions.
- `tally_stock_items.py`: Extractor for Stock Item masters.
- `tally_stock_vouchers.py`: Extractor for Inventory transactions (Stock Ledgers).
- `requirements.txt`: List of Python dependencies.

## 📋 Prerequisites

1. **Tally Running**: Tally must be open with at least one company loaded.
2. **Connectivity**: Ensure the Tally ODBC/XML port is enabled (Default: 9000).
   - Verify in Tally: `F1` > `Settings` > `Connectivity`.
3. **No application runtime required**: The standalone Windows executable does
   not require an installer or an additional application runtime. Python is
   needed only when using the optional scripts directly.

## 🚀 Getting Started

1. Download `TallyExporter.exe` from the
   [latest release](https://github.com/SRathinaGiri/TallyPy/releases/latest).
2. Start Tally and open the required company.
3. Run `TallyExporter.exe`, connect to Tally, load the tables, and export CSVs.

To build or inspect the source:

1. **Clone the repository**:
   ```bash
   git clone https://github.com/SRathinaGiri/TallyPy.git
   cd TallyPy
   ```

2. Open [`HMG/README.md`](HMG/README.md) for the native build instructions.

## 📊 Power BI Integration

You can use the scripts in this project directly in Power BI Desktop:
1. Go to **Get Data** > **Python Script**.
2. Copy the contents of the desired `.py` file.
3. Replace the `HOST` and `PORT` variables if your Tally is not on `localhost:9000`.
4. Power BI will generate a table named `dataset` containing your Tally data.

## ⚖️ License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
