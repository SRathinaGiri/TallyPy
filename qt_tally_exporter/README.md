# Tally Qt Exporter

Native Qt Widgets desktop app for exporting Tally data to CSV using the same XML request and parsing flow as `app1.py`.

Current release version: `0.1.1`

## Build

From this folder:

```bat
build_release.bat
```

This uses:

- `D:\Qt\6.11.0\mingw_64`
- `D:\Qt\Tools\CMake_64`
- `D:\Qt\Tools\Ninja`
- `D:\Qt\Tools\mingw1310_64`

## Deploy standalone app

After building:

```bat
deploy_release.bat
```

This creates a portable distributable folder at:

`dist\TallyQtExporter`

That folder contains the `.exe`, Qt runtime DLLs, plugins, and MinGW runtime DLLs so end users can run it without installing Qt or Python.

## Build installer

After `deploy_release.bat`:

```bat
build_installer.bat
```

This creates:

`dist\TallyQtExporter_Setup_v0.1.1.exe`

The installer copies the app to `C:\Program Files\Tally Qt Exporter` and creates Desktop and Start Menu shortcuts.

## Notes

- Connection defaults are `localhost:9000`
- If company or dates are left blank, the app asks Tally for the active company and period
- Exports available:
  - `vouchers.csv`
  - `allvouchers.csv`
  - `ledgers.csv`
  - `stock_items.csv`
  - `stock_vouchers.csv`
