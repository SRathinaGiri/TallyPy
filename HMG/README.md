# Tally HMG Exporter

This folder contains the native HMG/Harbour port of **Tally Qt Exporter**.

- `../qt_tally_exporter/` contains the Qt implementation from
  [SRathinaGiri/TallyPy](https://github.com/SRathinaGiri/TallyPy/tree/main/qt_tally_exporter).
- `TallyExporter.prg` is the HMG user interface.
- `tally_service.prg` is the native Tally HTTP, XML/TDL, parsing, and CSV layer.
- `TallyExporter.hbp` is also the HMG IDE project: `TallyExporter.prg` is listed
  first and contains `Main()` and the main window.

The HMG application uses Windows `WinHttp.WinHttpRequest.5.1` for HTTP and
`MSXML2.DOMDocument.6.0` for XML parsing. Both are Windows components, so the
deployed application does not need Qt, Python, plug-ins, or application DLLs.

## Build

The project targets the latest 64-bit toolchain installed at `D:\hmg5`.

```bat
build.bat
```

The build produces `TallyExporter.exe` in this folder. `TallyExporter.hbp`
contains the Harbour project definition and `TallyExporter.rc` embeds the
upstream application icon plus the **Tally XML Exporter** startup splash.

To work in HMG IDE, open `TallyExporter.hbp`. Its companion
`TallyExporter.hbc` contains the IDE/build configuration; the first project
module, `TallyExporter.prg`, is the main program.

## Use

1. Start Tally and enable its HTTP/XML server (normally port `9000`).
2. Open the required company in Tally.
3. Run `TallyExporter.exe` and select **Connect**. Empty company and date fields
   are populated from the active company.
4. Select **Load tables**. The program injects temporary TDL collections and
   reads ledger, voucher, stock-item, and inventory-entry XML.
5. Choose an export folder and select **Export all CSVs**.

Generated files:

- `vouchers.csv`
- `allvouchers.csv`
- `ledgers.csv`
- `stock_items.csv`
- `stock_vouchers.csv`

The UI uses virtual grids so it can display every loaded row without copying
the complete datasets into grid controls. Voucher and inventory requests are
split by calendar month to keep Tally responses manageable. CSV files are
UTF-8 with a BOM and are streamed to disk row-by-row.

## Native deployment

Copy only `TallyExporter.exe` to another Windows 10/11 computer. The executable
links the Harbour/HMG application and compiler runtimes statically; the target
computer supplies the standard Windows UCRT, WinHTTP, and MSXML components.

